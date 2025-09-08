import streamlit as st
import pandas as pd
import io
import re
from typing import Tuple
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# --------- Hilfen ---------
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _is_undef(val: any) -> bool:
    """Hilfsfunktion: prueft, ob eBKP-H-Wert als 'nicht definiert' gilt."""
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
    """Sicherer Check, ob Wert NA ist (verhindert bool-Kontext von pd.NA)."""
    try:
        return pd.isna(x)
    except Exception:
        return False


# --------- Kernbereinigung ---------
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: list[str] | None = None,  # Liste exakter eBKP-H-Werte zum Droppen (nur Sub-Zeilen)
) -> pd.DataFrame:
    """
    Grundregel:
    Wenn eine Sub-Zeile zur Hauptzeile befoerdert wird, werden ALLE vorhandenen '... Sub'
    Werte verwendet und ersetzen die Basiswerte. Fehlt ein '... Sub' Wert, wird der Mutterwert uebernommen.
    GUID bleibt fuer alle Elemente erhalten; beim Promoten gilt 'GUID Sub' > 'GUID'.
    """
    drop_sub_values = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}

    def _matches_drop_values(val: any) -> bool:
        if _is_na(val):
            return False
        return str(val).strip().lower() in drop_sub_values if drop_sub_values else False

    # Master-Kontextspalten (vom Mutterelement vererben)
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Paare Basis / Sub ermitteln (GUID bewusst ausschliessen)
    sub_pairs = sorted({
        base for c in df.columns
        if c.endswith(" Sub") and (base := c[:-4]) in df.columns and base != "GUID"
    })

    # Mehrschichtiges Element flaggen: wenn alle master_cols leer → Sub-Zeile
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(c)) for c in master_cols), axis=1
    )

    # 1) Mutter-Kontext an Subs vererben (inkl. eBKP-H, wenn Sub fehlt/unklassifiziert/nicht verfügbar)
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                # a) Master-Kontext vererben
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                # b) eBKP-H vererben, falls eBKP-H Sub nicht definiert
                if "eBKP-H" in df.columns:
                    mother_ebkp = df.at[i, "eBKP-H"]
                    sub_ebkp_sub = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if _has_value(mother_ebkp) and _is_undef(sub_ebkp_sub):
                        df.at[j, "eBKP-H"] = mother_ebkp
                j += 1
            i = j
        else:
            i += 1

    # 2) Generische Regel fuer Sub-Zeilen anwenden:
    #    Fuer jedes Basis/Sub-Paar gilt: Sub bevorzugen, sonst Mutter
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
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

    # 3) Subs zu Hauptzeilen promoten; Mutter droppen, wenn mind. 1 nutzbarer Sub
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            # Sub-Zeilen anhand exakter eBKP-H Auswahlliste droppen
            if drop_sub_values:
                for idx in list(sub_idxs):
                    ebkp_sub = df.at[idx, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else None
                    ebkp     = df.at[idx, "eBKP-H"] if "eBKP-H" in df.columns else None
                    cond_sub = _matches_drop_values(ebkp_sub)
                    cond_mom = _matches_drop_values(ebkp)
                    if cond_sub or cond_mom:
                        drop_idx.append(idx)
                        sub_idxs.remove(idx)

            if sub_idxs:
                # Nutzbare Subs vorhanden: Mutter droppen und Subs promoten
                drop_idx.append(i)
                for idx in sub_idxs:
                    new = df.loc[idx].copy()

                    # GUID der Sub bevorzugen; Fallback auf GUID
                    if "GUID Sub" in df.columns and _has_value(df.at[idx, "GUID Sub"]):
                        new["GUID"] = df.at[idx, "GUID Sub"]
                    elif "GUID" in df.columns:
                        new["GUID"] = df.at[idx, "GUID"]

                    # Fuer alle Basis/Sub-Paare: Sub bevorzugen, sonst Mutter
                    for base in sub_pairs:
                        sub_col = f"{base} Sub"
                        if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                            new[base] = df.at[idx, sub_col]
                        else:
                            new[base] = df.at[i, base]

                    new["Mehrschichtiges Element"] = False
                    for c in master_cols:
                        new[c] = df.at[i, c]

                    new_rows.append(new)
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

    # 4) ... Sub-Spalten entfernen
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # 5) Restbereinigung
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

    # 6) Deduplizieren (GUID-identische Voll-Duplikate)
    def _remove_exact_duplicates(d: pd.DataFrame) -> pd.DataFrame:
        if "GUID" not in d.columns:
            return d
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)

    df = _remove_exact_duplicates(df)

    # 7) Standardisieren & Werte bereinigen
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# --------- Streamlit Tab ---------
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.header("Vererbung & Mengenuebernahme")
    st.markdown("""
    **Logik:**  
    1) eBKP-H der Mutter → an Subs vererben, wenn `eBKP-H Sub` nicht definiert ist.  
    2) Generisch: Fuer jedes Basis/`... Sub`-Paar gilt **Sub bevorzugen, sonst Mutter**.  
    3) Subs als Hauptzeilen promoten; dabei alle vorhandenen `... Sub`-Werte uebernehmen; Mutter droppen.  
    4) 'Nicht klassifiziert', 'Keine Zuordnung', 'Nicht verfügbar' gelten als nicht definiert.

    **Steuerung:**  
    - Auswahl unten: Sub-Zeilen droppen, deren eBKP-H exakt in der Liste ist.  
    - Verarbeitung startet erst beim Button-Klick (Formular).
    """)

    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="vererbung_file_uploader")
    if not uploaded_file:
        return

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        return

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df.head(15), use_container_width=True)

    # Kandidaten fuer Dropdown: union aus eBKP-H und eBKP-H Sub
    ebkp_candidates = []
    if "eBKP-H" in df.columns:
        ebkp_candidates.extend(df["eBKP-H"].dropna().astype(str).str.strip().tolist())
    if "eBKP-H Sub" in df.columns:
        ebkp_candidates.extend(df["eBKP-H Sub"].dropna().astype(str).str.strip().tolist())
    ebkp_options = sorted({v for v in ebkp_candidates if v})

    with st.form(key="vererbung_form"):
        sel_drop_values = st.multiselect(
            "Sub-Zeilen droppen, wenn eBKP-H exakt gleich einem der folgenden Werte ist",
            options=ebkp_options,
            default=[],
            help="Mehrfachauswahl moeglich. Es wird exakt auf den Text verglichen."
        )
        run = st.form_submit_button("Verarbeitung starte")

    if run:
        with st.spinner("Verarbeitung laeuft ..."):
            df_clean = _process_df(df.copy(), drop_sub_values=sel_drop_values)
            df_clean = convert_quantity_columns(df_clean)

        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_clean.head(15), use_container_width=True)

        # Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False)
        output.seek(0)
        file_name = f"{(supplement_name or '').strip() or 'default'}_vererbung_mengen.xlsx"

        st.download_button(
            "Bereinigte Datei herunterladen",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
