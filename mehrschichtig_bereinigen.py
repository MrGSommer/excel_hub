import streamlit as st
import pandas as pd
import io
import re
from typing import Tuple, Dict, Any
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


def clean_dataframe(
    df: pd.DataFrame,
    delete_enabled: bool = False,
    custom_chars: str = "",
    match_sub_toggle: bool = False,
    drop_treppe_sub: bool = False,
    config: Dict[str, Dict[str, str]] | None = None,
    group_col: str | None = None,
    inherit_mother_ebkph_if_sub_missing: bool = False,
    global_sources_per_pair: Dict[str, str] | None = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Bereinigt die mehrschichtigen Daten.
    - Vererbung eBKP-H von Mutter an Subs (optional, nur wenn eBKP-H Sub fehlt/unklassifiziert).
    - Subs als Hauptzeilen aufschlüsseln; Mutter nur droppen, wenn mind. ein Sub nutzbar ist.
    - Spezialfall 'Treppe': Mutter nie droppen; Subs je nach Einstellung droppen.
    - Konfigurator steuert je Gruppe/Feldpaar (nur für '... Sub'-Paare): Auto/Mutter/Sub.
    """
    stats = {
        "inherited_ebkph": 0,
        "mothers_dropped": 0,
        "treppe_subs_dropped": 0,
    }

    def has_value(x) -> bool:
        return pd.notna(x) and str(x).strip() != ""

    # Masterspalten (werden von Mutter an Subs vererbt)
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Feldpaare (nur Basisspalten, die eine "... Sub" besitzen)
    sub_pairs = sorted({
        base for col in df.columns
        if col.endswith(" Sub") and (base := col[:-4]) in df.columns
    })

    # 1) Flag mehrschichtig
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(col)) for col in master_cols), axis=1
    )

    # 2) Werte aus Mutter in Subs für master_cols
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

    # 3) Nicht klassifizierte Hauptelemente ohne Subs entfernen
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

    # 4) Vererbung eBKP-H Mutter -> Subs (nur wenn Sub fehlt/unklassifiziert)
    if inherit_mother_ebkph_if_sub_missing and "eBKP-H" in df.columns:
        i = 0
        while i < len(df):
            if not df.at[i, "Mehrschichtiges Element"]:
                j = i + 1
                while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                    mother_ebkph = df.at[i, "eBKP-H"]
                    sub_ebkph = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if pd.notna(mother_ebkph) and (
                        pd.isna(sub_ebkph) or str(sub_ebkph).strip() in ["", "Nicht klassifiziert", "Keine Zuordnung"]
                    ):
                        df.at[j, "eBKP-H"] = mother_ebkph
                        stats["inherited_ebkph"] += 1
                    j += 1
                i = j
            else:
                i += 1

    # 5) Aufschlüsseln: Subs zu Hauptzeilen
    #    - nutzbarer Sub: hat gültiges eBKP-H Sub ODER (durch Vererbung) eBKP-H
    #    - Treppe: Mutter nie droppen; Sub ggf. droppen (stats)
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

            # Treppe-Erkennung (in Mutter oder in deren Subs)
            mother_txt = str(df.at[i, "eBKP-H"]) if "eBKP-H" in df.columns else ""
            treppe_case = "Treppe" in mother_txt
            for idx in sub_idxs:
                if "eBKP-H Sub" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H Sub"]):
                    treppe_case = True
                if "eBKP-H" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H"]):
                    treppe_case = True

            # Bestimme, ob es min. einen nutzbaren Sub gibt
            usable_subs = []
            for idx in sub_idxs:
                ebkph_sub = df.at[idx, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                ebkph = df.at[idx, "eBKP-H"] if "eBKP-H" in df.columns else pd.NA
                is_valid_sub = (
                    (has_value(ebkph_sub) and str(ebkph_sub) not in ["Nicht klassifiziert", "Keine Zuordnung"]) or
                    (has_value(ebkph) and str(ebkph) not in ["Nicht klassifiziert", "Keine Zuordnung"])
                )
                if is_valid_sub:
                    usable_subs.append(idx)

            if sub_idxs:
                if treppe_case:
                    # Mutter NIE droppen
                    if drop_treppe_sub:
                        # Subs mit Treppe droppen
                        for idx in sub_idxs:
                            if (
                                ("eBKP-H Sub" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H Sub"])) or
                                ("eBKP-H" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H"]))
                            ):
                                drop_idx.append(idx)
                                stats["treppe_subs_dropped"] += 1
                    # Keine weitere Aufschlüsselung für Treppe (Mutter bleibt Anker)
                    i = j
                    continue

                if usable_subs:
                    # Mutter droppen und nutzbare Subs zu neuen (Haupt-)Zeilen erheben
                    drop_idx.append(i)
                    stats["mothers_dropped"] += 1
                    for idx in usable_subs:
                        new = df.loc[idx].copy()
                        # Markiere als NICHT mehrschichtig (wird Hauptzeile)
                        new["Mehrschichtiges Element"] = False
                        # Masterwerte der Mutter in neue Hauptzeile übernehmen
                        for c in master_cols:
                            new[c] = df.at[i, c]
                        new_rows.append(new)
                else:
                    # Keine nutzbaren Subs -> Subs entfernen, Mutter bleibt
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

    # 6) Konfigurator anwenden (nur Paare mit "... Sub")
    if config and group_col and group_col in df.columns:
        i = 0
        while i < len(df):
            # Mutter als Anker
            if not df.at[i, "Mehrschichtiges Element"]:
                j = i + 1
                sub_idxs = []
                while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                    sub_idxs.append(j)
                    j += 1
                grp = df.at[i, group_col]
                grp_cfg = config.get(grp, {})
                for idx in sub_idxs:
                    for base in sub_pairs:
                        # Reihenfolge: Gruppen-Override -> globaler Default -> Auto
                        choice = grp_cfg.get(base, (global_sources_per_pair or {}).get(base, "Auto"))
                        sub_col = f"{base} Sub"
                        if choice == "Mutter":
                            df.at[idx, base] = df.at[i, base]
                        elif choice == "Sub":
                            if sub_col in df.columns and has_value(df.at[idx, sub_col]):
                                df.at[idx, base] = df.at[idx, sub_col]
                        else:
                            # Auto: Standardverhalten -> Sub wenn vorhanden, sonst lassen
                            if sub_col in df.columns and has_value(df.at[idx, sub_col]):
                                df.at[idx, base] = df.at[idx, sub_col]
                i = j
            else:
                i += 1
    else:
        # Fallback: Sub-Werte ins Basisfeld, wenn vorhanden
        for idx, row in df.iterrows():
            if row.get("Mehrschichtiges Element", False):
                for base in sub_pairs:
                    sub_col = f"{base} Sub"
                    if sub_col in df.columns and has_value(row[sub_col]):
                        df.at[idx, base] = row[sub_col]

    # 7) Sub-Spalten entfernen
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # 8) Restbereinigung
    for c in ["Einzelteile", "Farbe"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA

    if "eBKP-H" in df.columns:
        df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])]

    df.reset_index(drop=True, inplace=True)

    # 9) Echte Duplikate (identische GUID-Rekorde) entfernen
    def remove_exact_duplicates(d: pd.DataFrame) -> pd.DataFrame:
        if "GUID" not in d.columns:
            return d
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)

    df = remove_exact_duplicates(df)

    # 10) Standardisieren & Werte bereinigen
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled, custom_chars)

    return df, stats


def app(supplement_name, delete_enabled, custom_chars):
    st.header("Mehrschichtig Bereinigen")
    st.markdown("""
    **Ablauf**  
    1) Excel laden  
    2) Globale Defaults je Paar + Gruppen-Abweichungen setzen  
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
    st.dataframe(df.head(15), width="stretch")

    # --- Gruppierungsspalte wählen ---
    group_col = st.selectbox(
        "Spalte für Gruppierung wählen",
        [c for c in df.columns if c != "GUID"],
        index=0,
        key="group_col_select"
    )

    # --- Paare ermitteln: nur Basisspalten mit '... Sub' ---
    options = ["Auto", "Mutter", "Sub"]
    pair_bases = sorted({
        base for col in df.columns
        if col.endswith(" Sub") and (base := col[:-4]) in df.columns and base != "GUID"
    })

    if "config_sources" not in st.session_state:
        st.session_state.config_sources = {}
    if "global_sources_per_pair" not in st.session_state:
        st.session_state.global_sources_per_pair = {b: "Auto" for b in pair_bases}

    # --- Globale Defaults je Paar ---
    st.markdown("### Globale Einstellungen je Paar")
    if pair_bases:
        global_cfg_df = pd.DataFrame({
            "Feld": pair_bases,
            "Globale Quelle": [st.session_state.global_sources_per_pair.get(b, "Auto") for b in pair_bases]
        })
        global_cfg_df = st.data_editor(
            global_cfg_df,
            num_rows="fixed",
            disabled=["Feld"],
            width="stretch",
            column_config={
                "Feld": st.column_config.TextColumn("Feld (Basis)"),
                "Globale Quelle": st.column_config.SelectboxColumn(
                    "Globale Quelle", options=options,
                    help="Standardquelle für dieses Feld (Auto/Mutter/Sub)"
                )
            },
            key="global_cfg_table"
        )
        for _, row in global_cfg_df.iterrows():
            st.session_state.global_sources_per_pair[row["Feld"]] = row["Globale Quelle"]
    else:
        st.info("Keine konfigurierbaren Paare gefunden (keine '... Sub'-Spalten).")

    # --- Gruppen (Abweichungen) ---
    st.markdown("### Gruppen-Konfigurator (Abweichungen von global)")
    config: Dict[str, Dict[str, str]] = {}
    groups = sorted(df[group_col].dropna().unique()) if group_col in df.columns else []
    cols_layout = st.columns(2) if len(groups) > 1 else [st]

    for gi, group in enumerate(groups):
        with cols_layout[gi % 2].expander(f"{group_col} = {group}", expanded=False):
            if pair_bases:
                defaults = [
                    st.session_state.config_sources.get((group, b), st.session_state.global_sources_per_pair[b])
                    for b in pair_bases
                ]
                cfg_df = pd.DataFrame({"Feld": pair_bases, "Quelle": defaults})
                cfg_df = st.data_editor(
                    cfg_df,
                    num_rows="fixed",
                    disabled=["Feld"],
                    width="stretch",
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

    # Abweichungen ggü. globalen Defaults (für Log)
    overrides = []
    for g in groups:
        for b in pair_bases:
            grp_val = st.session_state.config_sources.get((g, b), st.session_state.global_sources_per_pair[b])
            glob_val = st.session_state.global_sources_per_pair[b]
            if grp_val != glob_val:
                overrides.append({"Gruppe": g, "Feld": b, "Global": glob_val, "Override": grp_val})

    st.caption("Hinweis: GUID bleibt immer aus der jeweiligen Zeile (nie überschrieben).")

    # --- Spezielle Optionen ---
    use_match = st.checkbox("Sub-eBKP-H aufschlüsseln (Speziallogik Alt)", value=False)
    drop_treppe = st.checkbox("Bei 'Treppe' Sub-Zeilen droppen (Mutter bleibt)", value=True)
    inherit_mother_ebkph = st.checkbox(
        "eBKP-H der Mutter an Subs vererben, wenn eBKP-H Sub fehlt/nicht klassifiziert",
        value=True
    )

    # --- Verarbeitung starten ---
    if st.button("Verarbeitung starten"):
        with st.spinner("Daten werden bereinigt ..."):
            df_clean, stats = clean_dataframe(
                df.copy(),
                delete_enabled=delete_enabled,
                custom_chars=custom_chars,
                match_sub_toggle=use_match,
                drop_treppe_sub=drop_treppe,
                config=config,
                group_col=group_col,
                inherit_mother_ebkph_if_sub_missing=inherit_mother_ebkph,
                global_sources_per_pair=st.session_state.global_sources_per_pair,
            )

        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_clean.head(15), width="stretch")

        # Log / KPIs
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Vererbte eBKP-H an Subs", stats["inherited_ebkph"])
        col_b.metric("Gedroppte Mütter", stats["mothers_dropped"])
        col_c.metric("Gedroppte Treppen-Subs", stats["treppe_subs_dropped"])
        if overrides:
            st.markdown("**Abweichungen von globalen Defaults:**")
            st.dataframe(pd.DataFrame(overrides), width="stretch")

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
