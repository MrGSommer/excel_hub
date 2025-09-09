import streamlit as st
import pandas as pd
import io
from typing import Optional, List
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# ========= Hilfen =========
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _is_undef(val: any) -> bool:
    """eBKP-H-Wert gilt als 'nicht definiert'."""
    if not _has_value(val):
        return True
    txt = str(val).strip().lower()
    return (
        txt == ""
        or "icht klassifiziert" in txt
        or "eine zuordnung" in txt
        or "icht verfügbar" in txt
    )


def _is_na(x) -> bool:
    """Sicherer NA-Check (verhindert bool-Kontext von pd.NA)."""
    try:
        return pd.isna(x)
    except Exception:
        return False


# ========= Kernbereinigung =========
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: Optional[List[str]] = None,  # eBKP-H-Werte: Sub-Zeilen mit diesen Texten ignorieren (droppen)
) -> pd.DataFrame:
    """
    Logik:
    - Master-Kontext an Subs vererben. eBKP-H der Mutter vererben, wenn 'eBKP-H Sub' undefiniert.
    - Generisch: fuer Basis/'... Sub'-Paare gilt 'Sub bevorzugen, sonst Mutter'.
    - Subs zu Hauptzeilen promoten; Mutter wird gedropt, wenn mind. 1 verwertbarer Sub existiert.
    - GUID bleibt die der Sub. Neue Spalte 'GUID Gruppe' enthaelt immer die GUID der Mutter.
    """
    drop_set = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}

    def _matches_drop_values(val: any) -> bool:
        if _is_na(val):
            return False
        return str(val).strip().lower() in drop_set if drop_set else False

    # Master-Kontextspalten (falls vorhanden)
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Paare Basis / Sub (GUID bewusst ausschliessen)
    sub_pairs = sorted({
        base for c in df.columns
        if c.endswith(" Sub") and (base := c[:-4]) in df.columns and base != "GUID"
    })

    # Zeilentypen markieren
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(c)) for c in master_cols), axis=1
    )
    df["Promoted"] = False
    df["GUID Gruppe"] = pd.NA  # immer GUID der Mutter

    # ===== 1) Mutter-Kontext + eBKP-H vererben =====
    i = 0
    while i < len(df):
        # Mutterzeile = alle master_cols belegt
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            mother_guid = df.at[i, "GUID"] if "GUID" in df.columns else pd.NA
            # Mutter kennt ihre Gruppe
            df.at[i, "GUID Gruppe"] = mother_guid

            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                # a) Master-Kontext vererben
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                # b) eBKP-H vererben, falls 'eBKP-H Sub' nicht definiert
                if "eBKP-H" in df.columns:
                    mother_ebkp = df.at[i, "eBKP-H"]
                    sub_ebkp_sub = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if _has_value(mother_ebkp) and _is_undef(sub_ebkp_sub):
                        df.at[j, "eBKP-H"] = mother_ebkp
                j += 1
            i = j
        else:
            i += 1

    # ===== 2) Sub bevorzugen, sonst Mutter =====
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1
            for idx in sub_idxs:
                for base in sub_pairs:
                    sub_col = f"{base} Sub"
                    if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                        df.at[idx, base] = df.at[idx, sub_col]
                    else:
                        df.at[idx, base] = df.at[i, base]
            i = j
        else:
            i += 1

    # ===== 3) Subs promoten; Mutter ggf. droppen =====
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            mother_guid = df.at[i, "GUID"] if "GUID" in df.columns else pd.NA
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            # a) Subs anhand eBKP-H Ignorierliste verwerfen
            if drop_set:
                for idx in list(sub_idxs):
                    ebkp_sub = df.at[idx, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else None
                    ebkp     = df.at[idx, "eBKP-H"] if "eBKP-H" in df.columns else None
                    if _matches_drop_values(ebkp_sub) or _matches_drop_values(ebkp):
                        drop_idx.append(idx)
                        sub_idxs.remove(idx)

            # b) Wenn nutzbare Subs existieren → Mutter droppen und promoten
            if sub_idxs:
                drop_idx.append(i)
                for idx in sub_idxs:
                    new = df.loc[idx].copy()

                    # GUID bleibt Sub (oder Fallback)
                    if "GUID Sub" in df.columns and _has_value(df.at[idx, "GUID Sub"]):
                        new["GUID"] = df.at[idx, "GUID Sub"]
                    elif "GUID" in df.columns:
                        new["GUID"] = df.at[idx, "GUID"]

                    # Sub/Basis-Paare finalisieren
                    for base in sub_pairs:
                        sub_col = f"{base} Sub"
                        if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                            new[base] = df.at[idx, sub_col]
                        else:
                            new[base] = df.at[i, base]

                    # Metadaten
                    new["Mehrschichtiges Element"] = False
                    new["Promoted"] = True
                    new["GUID Gruppe"] = mother_guid
                    for c in master_cols:
                        new[c] = df.at[i, c]

                    new_rows.append(new)
            else:
                # Mutter bleibt alleine bestehen
                df.at[i, "Mehrschichtiges Element"] = False
                df.at[i, "Promoted"] = False
                df.at[i, "GUID Gruppe"] = mother_guid

            i = j
        else:
            i += 1

    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # ===== 4) ... Sub-Spalten entfernen (Inhalte sind uebernommen) =====
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # ===== 5) Restbereinigung =====
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA
    if "eBKP-H" in df.columns:
        mask_invalid = df["eBKP-H"].astype(str).str.lower().str.contains(
            "nicht klassifiziert|keine zuordnung|nicht verfügbar", na=True
        )
        df = df[~mask_invalid]
    for c in ["Einzelteile", "Farbe"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    df.reset_index(drop=True, inplace=True)

    # ===== 6) Deduplizieren (exact duplicates je GUID) =====
    def _remove_exact_duplicates(d: pd.DataFrame) -> pd.DataFrame:
        if "GUID" not in d.columns:
            return d
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)

    df = _remove_exact_duplicates(df)

    # ===== 7) Standardisieren & Werte bereinigen =====
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# ========= Auswertung =========
def summarize_column(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Werte in Spalte zaehlen und Anteil ausweisen; leere Werte ausblenden."""
    if col not in df.columns:
        return pd.DataFrame()
    s = df[col].astype(str).str.strip()
    s = s[s != ""]
    summary = (
        s.value_counts(dropna=True)
         .rename_axis(col)
         .reset_index(name="Anzahl")
    )
    total = int(summary["Anzahl"].sum()) if not summary.empty else 0
    summary["Anteil %"] = (summary["Anzahl"] / total * 100).round(2) if total else 0.0
    return summary


# ========= Streamlit App =========
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.set_page_config(page_title="Vererbung & Mengenuebernahme", layout="wide")
    st.header("Vererbung & Mengenuebernahme")

    st.markdown("""
    **Ablauf**
    1) Verarbeitung: eBKP-H Auswahl → Subs ignorieren (droppen) **waehrend** der Promotion.  
    2) Filter (optional): **nach** der Verarbeitung eine Spalte und Werte waehlen und Zeilen entfernen.  
    3) Finalisieren: Spalte fuer Zusammenfassung waehlen → Tabellen & Export.  
    """)

    # --- Datei laden ---
    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="vererbung_file_uploader")
    if not uploaded_file:
        st.stop()

    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        st.stop()

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df_raw.head(15), use_container_width=True)

    # Kandidaten fuer eBKP-H Dropdown
    ebkp_candidates = []
    if "eBKP-H" in df_raw.columns:
        ebkp_candidates.extend(df_raw["eBKP-H"].dropna().astype(str).str.strip().tolist())
    if "eBKP-H Sub" in df_raw.columns:
        ebkp_candidates.extend(df_raw["eBKP-H Sub"].dropna().astype(str).str.strip().tolist())
    ebkp_options = sorted({v for v in ebkp_candidates if v})

    # ===== Formular 1: Verarbeitung steuern =====
    with st.form(key="form_process"):
        sel_drop_values = st.multiselect(
            "Subs ignorieren (droppen), wenn eBKP-H gleich einem dieser Werte ist",
            options=ebkp_options,
            default=[],
            help="Exakter Textvergleich; wirkt waehrend der Promotion der Subs."
        )
        btn_process = st.form_submit_button("Verarbeitung starten")

    if btn_process:
        with st.spinner("Verarbeitung laeuft ..."):
            df_clean = _process_df(df_raw.copy(), drop_sub_values=sel_drop_values)
            df_clean = convert_quantity_columns(df_clean)

        st.session_state["df_clean"] = df_clean.copy()
        st.session_state["df_active"] = df_clean.copy()  # aktive Tabelle fuer nachgelagerte Schritte

    # Wenn noch nicht verarbeitet wurde, nichts weiter anzeigen
    if "df_active" not in st.session_state:
        st.info("Bitte zuerst die Verarbeitung starten.")
        st.stop()

    df_active = st.session_state["df_active"]

    st.subheader("Bereinigte Daten (15 Zeilen)")
    st.dataframe(df_active.head(15), use_container_width=True)

    # Export bereinigte Daten (aktueller aktiver Stand)
    out_clean = io.BytesIO()
    with pd.ExcelWriter(out_clean, engine="openpyxl") as writer:
        df_active.to_excel(writer, index=False, sheet_name="Bereinigt")
    out_clean.seek(0)
    st.download_button(
        "Aktuelle bereinigte Datei herunterladen",
        data=out_clean,
        file_name=f"{(supplement_name or '').strip() or 'default'}_vererbung_mengen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===== Formular 2: Nachgelagerter Filter (eine Spalte) =====
    st.markdown("---")
    st.subheader("Filter nach Verarbeitung (optional)")
    with st.form(key="form_filter"):
        col_options = ["-- Spalte waehlen --"] + list(df_active.columns)
        sel_col = st.selectbox("Spalte fuer Filter", options=col_options, index=0)
        value_options = []
        if sel_col and sel_col != "-- Spalte waehlen --":
            # Auswahlwerte aus aktueller aktiver Tabelle
            value_options = sorted({
                v for v in df_active[sel_col].dropna().astype(str).str.strip().tolist() if v
            })
        sel_values = st.multiselect(
            "Zeilen entfernen, wenn Wert gleich ist",
            options=value_options,
            default=[],
            help="Exakter Textvergleich; mehrere Eintraege moeglich."
        )
        btn_apply_filter = st.form_submit_button("Filter anwenden")

    if btn_apply_filter:
        if sel_col and sel_col != "-- Spalte waehlen --" and sel_values:
            mask_drop = df_active[sel_col].astype(str).str.strip().isin(set(sel_values))
            df_filtered = df_active.loc[~mask_drop].reset_index(drop=True)
            st.session_state["df_active"] = df_filtered.copy()
            df_active = df_filtered
            st.success(f"Filter angewandt: {len(mask_drop[mask_drop].index)} Zeilen entfernt.")
        else:
            st.warning("Bitte Spalte und mindestens einen Wert waehlen.")

    # Vorschau nach Filter
    st.dataframe(df_active.head(15), use_container_width=True)

    # ===== Formular 3: Finalisieren =====
    st.markdown("---")
    st.subheader("Finalisieren (Zusammenfassung & Export)")
    with st.form(key="form_finalize"):
        final_col = st.selectbox(
            "Spalte fuer Zusammenfassung",
            options=["-- Spalte waehlen --"] + list(df_active.columns),
            index=0,
            help="Werte werden gezaehlt und prozentual dargestellt."
        )
        btn_finalize = st.form_submit_button("Finalisieren")

    if btn_finalize:
        if final_col and final_col != "-- Spalte waehlen --":
            summary_df = summarize_column(df_active, final_col)
            if summary_df.empty:
                st.info("Keine gueltigen Werte fuer die Zusammenfassung gefunden.")
            else:
                st.markdown(f"**Zusammenfassung: {final_col}**")
                st.dataframe(summary_df, use_container_width=True)

                # Export: aktive Tabelle + Summary
                out_final = io.BytesIO()
                with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
                    df_active.to_excel(writer, index=False, sheet_name="Bereinigt_Aktiv")
                    summary_df.to_excel(writer, index=False, sheet_name="Auswertung")
                out_final.seek(0)
                st.download_button(
                    "Bereinigte aktive Datei + Auswertung herunterladen",
                    data=out_final,
                    file_name=f"{(supplement_name or '').strip() or 'default'}_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Bitte eine Spalte fuer die Zusammenfassung waehlen.")
