import streamlit as st
import pandas as pd
import io
import re
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


def clean_dataframe(df: pd.DataFrame,
                    delete_enabled: bool = False,
                    custom_chars: str = "",
                    match_sub_toggle: bool = False,
                    drop_treppe_sub: bool = False,
                    config: dict | None = None,
                    group_col: str | None = None) -> pd.DataFrame:
    """
    Bereinigt die mehrschichtigen Daten.
    Konfigurator steuert pro (Gruppe, Feld) ob Wert aus Mutter oder Sub übernommen wird – nur für Felder mit '... Sub'-Paar.
    """
    # -----------------------------
    # 0) Grundsetup / Hilfen
    # -----------------------------
    def has_value(x) -> bool:
        return pd.notna(x) and str(x).strip() != ""

    # Masterspalten (werden von Mutter an Subs vererbt)
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Paare: nur Felder, die eine " Sub"-Spalte besitzen UND das Basisfeld existiert
    sub_pairs = sorted({
        base for col in df.columns
        if col.endswith(" Sub") and (base := col[:-4]) in df.columns
    })

    # -----------------------------
    # 1) Flag mehrschichtig / Vererbung Masterspalten
    # -----------------------------
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(col)) for col in master_cols), axis=1
    )

    # Werte aus Mutterzeile in Subs vererben (nur master_cols)
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                j += 1
            i = j
        else:
            i += 1

    # -----------------------------
    # 2) Nicht klassifizierte Hauptelemente ohne Subs entfernen
    # -----------------------------
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1
            if not sub_idxs and "eBKP-H" in df.columns and df.at[i, "eBKP-H"] == "Nicht klassifiziert":
                drop_idx.append(i)
            i = j
        else:
            i += 1
    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)

    # -----------------------------
    # 3) Aufschlüsseln mehrschichtiger Elemente (bestehende Logik)
    # -----------------------------
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            if sub_idxs and match_sub_toggle and "eBKP-H Sub" in df.columns:
                mother_val = df.at[i, "eBKP-H"] if "eBKP-H" in df.columns else pd.NA
                for idx in sub_idxs:
                    sub_val = df.at[idx, "eBKP-H Sub"]
                    # 1) Treppe-Subs droppen
                    if "Treppe" in str(mother_val) and "Treppe" in str(sub_val):
                        if drop_treppe_sub:
                            drop_idx.append(idx)
                    # 2) andere klassifizierte Subs zu neuen Mutterzeilen
                    elif has_value(sub_val) and sub_val not in ["Nicht klassifiziert", "Keine Zuordnung"]:
                        new = df.loc[idx].copy()
                        for c in master_cols:
                            new[c] = df.at[i, c]
                        new["Mehrschichtiges Element"] = False
                        new_rows.append(new)
                i = j
                continue

            # Standardfall ohne Toggle
            if sub_idxs:
                valid = any(
                    has_value(df.at[idx, "eBKP-H Sub"])
                    and df.at[idx, "eBKP-H Sub"] not in ["Nicht klassifiziert", "Keine Zuordnung"]
                    for idx in sub_idxs
                ) if "eBKP-H Sub" in df.columns else False
                if valid:
                    drop_idx.append(i)  # Mutter droppen
                    for idx in sub_idxs:
                        sub_val = df.at[idx, "eBKP-H Sub"]
                        if has_value(sub_val) and sub_val not in ["Nicht klassifiziert", "Keine Zuordnung"]:
                            new = df.loc[idx].copy()
                            new["Mehrschichtiges Element"] = True
                            new_rows.append(new)
                else:
                    # Subs ohne verwertbare Klassifikation entfernen; Mutter bleibt
                    drop_idx.extend(sub_idxs)
                    df.at[i, "Mehrschichtiges Element"] = False
            else:
                df.at[i, "Mehrschichtiges Element"] = False
            i = j
        else:
            i += 1

    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # -----------------------------
    # 4) KONFIGURATOR anwenden (nur für Paare mit "... Sub")
    #    -> Entscheidung pro Gruppe & Feld: "Mutter" oder "Sub"
    # -----------------------------
    if config and group_col and group_col in df.columns:
        i = 0
        while i < len(df):
            # Nur Mutterzeilen als Anker
            if not df.at[i, "Mehrschichtiges Element"]:
                j = i + 1
                sub_idxs = []
                while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                    sub_idxs.append(j)
                    j += 1
                grp = df.at[i, group_col]
                grp_cfg = config.get(grp, {})
                # Für jede Sub-Zeile die gewählten Quellen setzen
                for idx in sub_idxs:
                    for base in sub_pairs:
                        choice = grp_cfg.get(base, st.session_state.global_sources_per_pair.get(base, "Auto"))
                        if choice == "Mutter":
                            df.at[idx, base] = df.at[i, base]
                        elif choice == "Sub":
                            sub_col = f"{base} Sub"
                            if sub_col in df.columns and has_value(df.at[idx, sub_col]):
                                df.at[idx, base] = df.at[idx, sub_col]
                        # Auto => Standardlogik unverändert

                i = j
            else:
                i += 1
    else:
        # Fallback ohne Konfigurator: Standard – Sub-Werte ins Basisfeld, wenn vorhanden
        for idx, row in df.iterrows():
            if row.get("Mehrschichtiges Element", False):
                for base in sub_pairs:
                    sub_col = f"{base} Sub"
                    if sub_col in df.columns and has_value(row[sub_col]):
                        df.at[idx, base] = row[sub_col]

    # -----------------------------
    # 5) Sub-Spalten jetzt entfernen (nach angewandter Entscheidung)
    # -----------------------------
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # -----------------------------
    # 6) Restbereinigung
    # -----------------------------
    # Unnütze Spalten entfernen
    for c in ["Einzelteile", "Farbe"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    # 'oi' normalisieren
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA

    if "eBKP-H" in df.columns:
        df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])]

    df.reset_index(drop=True, inplace=True)

    # Echte Duplikate (identische GUID-Rekorde) entfernen
    def remove_exact_duplicates(d: pd.DataFrame) -> pd.DataFrame:
        if "GUID" not in d.columns:
            return d
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)

    df = remove_exact_duplicates(df)

    # Standardisieren & Werte bereinigen
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled, custom_chars)

    return df


def app(supplement_name, delete_enabled, custom_chars):
    st.header("Mehrschichtig Bereinigen")
    st.markdown("""
    **Einleitung:**  
    1) Excel laden  
    2) Konfiguration vornehmen (nur Paare mit " Sub")  
    3) Verarbeitung starten  
    """)

    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="bereinigen_file_uploader")
    if not uploaded_file:
        return

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        return

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df.head(15), use_container_width=True)

    # -----------------------------
    # Konfigurator (kompakt via Data Editor)
    # Nur Basisfelder, die eine "... Sub"-Spalte haben
    # -----------------------------
    # --- Globale Einstellungen je Werte-Paar ---
    options = ["Auto", "Mutter", "Sub"]
    
    pair_bases = sorted({
        base for col in df.columns
        if col.endswith(" Sub") and (base := col[:-4]) in df.columns and base != "GUID"
    })
    
    if "global_sources_per_pair" not in st.session_state:
        # Default: Auto für jedes Paar
        st.session_state.global_sources_per_pair = {b: "Auto" for b in pair_bases}
    
    st.markdown("### Globale Einstellungen je Paar")
    global_cfg_df = pd.DataFrame({
        "Feld": pair_bases,
        "Globale Quelle": [st.session_state.global_sources_per_pair.get(b, "Auto") for b in pair_bases]
    })
    
    global_cfg_df = st.data_editor(
        global_cfg_df,
        num_rows="fixed",
        disabled=["Feld"],
        use_container_width=True,
        column_config={
            "Feld": st.column_config.TextColumn("Feld (Basis)"),
            "Globale Quelle": st.column_config.SelectboxColumn(
                "Globale Quelle", options=options,
                help="Globale Standardquelle für dieses Feld (Auto/Mutter/Sub)"
            )
        },
        key="global_cfg_table"
    )
    
    # in Session speichern
    for _, row in global_cfg_df.iterrows():
        st.session_state.global_sources_per_pair[row["Feld"]] = row["Globale Quelle"]
    
    # -----------------------------
    # Gruppen-Konfigurator (Abweichungen von den globalen Defaults)
    # -----------------------------
    st.markdown("### Gruppen-Konfigurator (Abweichungen)")
    config = {}
    groups = sorted(df[group_col].dropna().unique()) if group_col in df.columns else []
    cols_layout = st.columns(2) if len(groups) > 1 else [st]
    
    for gi, group in enumerate(groups):
        with cols_layout[gi % 2].expander(f"{group_col} = {group}", expanded=False):
            defaults = [
                st.session_state.config_sources.get((group, b), st.session_state.global_sources_per_pair[b])
                for b in pair_bases
            ]
            cfg_df = pd.DataFrame({"Feld": pair_bases, "Quelle": defaults})
            cfg_df = st.data_editor(
                cfg_df,
                num_rows="fixed",
                disabled=["Feld"],
                use_container_width=True,
                column_config={
                    "Feld": st.column_config.TextColumn("Feld (Basis)"),
                    "Quelle": st.column_config.SelectboxColumn("Quelle", options=options)
                },
                key=f"cfg_table_{group}"
            )
            choices = dict(zip(cfg_df["Feld"], cfg_df["Quelle"]))
            for c, v in choices.items():
                st.session_state.config_sources[(group, c)] = v
            config[group] = choices          
          else:
              st.write("Keine Paare mit ' Sub' vorhanden.")

    st.caption("Hinweis: *GUID bleibt immer aus der jeweiligen Zeile. Sie wird nie überschrieben.*")

    # Weitere Optionen
    use_match = st.checkbox("Sub-eBKP-H aufschlüsseln (Speziallogik)", value=False)
    drop_treppe = st.checkbox("Bei 'Treppe' Sub-Zeilen droppen (falls Speziallogik aktiv)", value=False)

    # -----------------------------
    # Verarbeitung starten
    # -----------------------------
    if st.button("Verarbeitung starten"):
        with st.spinner("Daten werden bereinigt ..."):
            df_clean = clean_dataframe(
                df.copy(),
                delete_enabled=delete_enabled,
                custom_chars=custom_chars,
                match_sub_toggle=use_match,
                drop_treppe_sub=drop_treppe,
                config=config,
                group_col=group_col
            )

        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_clean.head(15), use_container_width=True)

        # Export mit Float-Spalten (Mengen)
        output = io.BytesIO()
        df_export = convert_quantity_columns(df_clean.copy())
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False)
        output.seek(0)

        file_name = f"{(supplement_name or '').strip() or 'default'}_bereinigt.xlsx"
        st.download_button(
            "Bereinigte Datei herunterladen",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
