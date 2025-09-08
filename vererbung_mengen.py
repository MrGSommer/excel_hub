import streamlit as st
import pandas as pd
import io
import re
from typing import Dict, Any, Tuple
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# --------- Hilfen ---------
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _to_float_try(s: Any) -> bool:
    """Grobe Heuristik, ob ein String zahlenaehnlich ist (inkl. 2.3e-4 usw.)."""
    if pd.isna(s):
        return False
    s = str(s)
    if s.strip() == "":
        return False
    # Einheit / Leerzeichen / Tausender trennen
    s1 = s.replace("\xa0", " ").strip()
    s1 = re.sub(r"\s*(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m|stk|Stk|pcs)\s*$", "", s1, flags=re.IGNORECASE)
    s1 = s1.replace("’", "").replace("'", "").replace(" ", "").replace(",", ".")
    # harte Bereinigung auf zulaessige Zeichen
    cleaned = re.sub(r"[^0-9eE\.\+\-]", "", s1)
    try:
        float(cleaned)
        return True
    except Exception:
        return False


# --------- Kernbereinigung ---------

def _process_df(
    df: pd.DataFrame,
    drop_treppe_sub: bool,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Neue Grundregel:
    Wenn eine Sub-Zeile zur Hauptzeile befoerdert wird, werden ALLE vorhandenen '... Sub'
    Werte verwendet und ersetzen die Basiswerte. Fehlt ein '... Sub' Wert, wird der Mutterwert uebernommen.
    GUID bleibt fuer alle Elemente erhalten; beim Promoten gilt 'GUID Sub' > 'GUID'.
    """
    stats = {"inherited_ebkph": 0, "mothers_dropped": 0, "treppe_subs_dropped": 0}

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

    # 1) Mutter-Kontext an Subs vererben (inkl. eBKP-H, wenn Sub fehlt/unklassifiziert)
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                # a) Master-Kontext vererben
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                # b) eBKP-H vererben, falls eBKP-H Sub fehlt/unklassifiziert
                if "eBKP-H" in df.columns:
                    mother_ebkp = df.at[i, "eBKP-H"]
                    sub_ebkp_sub = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if _has_value(mother_ebkp) and (
                        pd.isna(sub_ebkp_sub) or str(sub_ebkp_sub).strip() in ["", "Nicht klassifiziert", "Keine Zuordnung"]
                    ):
                        df.at[j, "eBKP-H"] = mother_ebkp
                        stats["inherited_ebkph"] += 1
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

            # Treppe-Sonderfall: Mutter nie droppen
            treppe_case = False
            mother_txt = str(df.at[i, "eBKP-H"]) if "eBKP-H" in df.columns else ""
            if "Treppe" in mother_txt:
                treppe_case = True
            for idx in sub_idxs:
                if "eBKP-H Sub" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H Sub"]):
                    treppe_case = True
                if "eBKP-H" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H"]):
                    treppe_case = True

            if sub_idxs:
                if treppe_case:
                    if drop_treppe_sub:
                        for idx in sub_idxs:
                            if (
                                ("eBKP-H Sub" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H Sub"])) or
                                ("eBKP-H" in df.columns and "Treppe" in str(df.at[idx, "eBKP-H"]))
                            ):
                                drop_idx.append(idx)
                                stats["treppe_subs_dropped"] += 1
                    i = j
                    continue

                # Nutzbare Subs vorhanden: Mutter droppen und Subs promoten
                drop_idx.append(i)
                stats["mothers_dropped"] += 1
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

                    # Flag setzen und Master-Kontext (bereits vererbt) sichern
                    new["Mehrschichtiges Element"] = False
                    for c in master_cols:
                        new[c] = df.at[i, c]

                    new_rows.append(new)
            else:
                # keine Subs
                df.at[i, "Mehrschichtiges Element"] = False

            i = j
        else:
            i += 1

    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # 4) ... Sub-Spalten entfernen (nachdem alle Werte uebernommen wurden)
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # 5) Restbereinigung
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA
    if "eBKP-H" in df.columns:
        df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])]
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

    return df, stats


# --------- Streamlit Tab ---------
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.header("Vererbung & Mengenuebernahme")
    st.markdown("""
    **Logik:**  
    1) eBKP-H der Mutter → an Subs vererben, wenn `eBKP-H Sub` fehlt/unklassifiziert.  
    2) Generisch: Fuer jedes Basis/`... Sub`-Paar gilt **Sub bevorzugen, sonst Mutter**.  
    3) Subs als Hauptzeilen promoten; dabei alle vorhandenen `... Sub`-Werte uebernehmen; Mutter droppen.  
       **Treppe**: Mutter bleibt; Subs optional droppen.
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

    drop_treppe_sub = st.checkbox("Bei 'Treppe' Sub-Zeilen droppen (Mutter bleiben)", value=True)

    if st.button("Verarbeitung starte"):
        with st.spinner("Verarbeitung laeuft ..."):
            df_clean, stats = _process_df(df.copy(), drop_treppe_sub=drop_treppe_sub)

        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_clean.head(15), use_container_width=True)

        # Export: Mengen als float
        output = io.BytesIO()
        df_export = convert_quantity_columns(df_clean.copy())
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False)
        output.seek(0)
        file_name = f"{(supplement_name or '').strip() or 'default'}_vererbung_mengen.xlsx"

        c1, c2, c3 = st.columns(3)
        c1.metric("Vererbte eBKP-H", stats["inherited_ebkph"])
        c2.metric("Gedroppte Muetter", stats["mothers_dropped"])
        c3.metric("Gedroppte Treppen-Subs", stats["treppe_subs_dropped"])

        st.download_button(
            "Bereinigte Datei herunterladen",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
