import streamlit as st
import pandas as pd
import io
from excel_utils import clean_columns_values, rename_columns_to_standard

def clean_dataframe(df, delete_enabled=False, custom_chars="", match_sub_toggle=False):
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain"]
    master_cols = [col for col in master_cols if col in df.columns]
    
    df["Mehrschichtiges Element"] = df.apply(lambda row: all(pd.isna(row[col]) for col in master_cols), axis=1)

    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            while j < len(df) and all(pd.isna(df.at[j, col]) for col in master_cols):
                for col in master_cols:
                    df.at[j, col] = df.at[i, col]
                j += 1
            i = j
        else:
            i += 1

    drop_indices = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            sub_indices = []
            while j < len(df) and all(pd.isna(df.at[j, col]) for col in master_cols):
                if df.at[j, "Mehrschichtiges Element"]:
                    sub_indices.append(j)
                j += 1
            if not sub_indices and df.at[i, "eBKP-H"] == "Nicht klassifiziert":
                drop_indices.append(i)
            i = j
        else:
            i += 1
    if drop_indices:
        df.drop(index=drop_indices, inplace=True)
        df.reset_index(drop=True, inplace=True)

    new_rows = []
    drop_indices = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            sub_indices = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_indices.append(j)
                j += 1


            # Aufschlüsselung mehrschichtiger Elemente
            if sub_indices and match_sub_toggle and "eBKP-H Sub" in df.columns:
                mother_val = df.at[i, "eBKP-H"]
                for idx in sub_indices:
                    sub_val = df.at[idx, "eBKP-H Sub"]
                    if sub_val == mother_val:
                        drop_indices.append(idx)
                    else:
                        drop_indices.append(i)
                        new = df.loc[idx].copy()
                        new["Mehrschichtiges Element"] = True
                        new_rows.append(new)
                i = j
                continue
            
            # ursprüngliches Verhalten, wenn Toggle aus oder keine Sub-Spalte
            if sub_indices:
                valid_sub_found = any(
                    pd.notna(df.at[idx, "eBKP-H Sub"]) and
                    df.at[idx, "eBKP-H Sub"] not in ["Nicht klassifiziert", "", "Keine Zuordnung"]
                    for idx in sub_indices
                )
                if valid_sub_found:
                    drop_indices.append(i)
                    for idx in sub_indices:
                        sub_val = df.at[idx, "eBKP-H Sub"]
                        if pd.notna(sub_val) and sub_val not in ["Nicht klassifiziert", "", "Keine Zuordnung"]:
                            new = df.loc[idx].copy()
                            new["Mehrschichtiges Element"] = True
                            new_rows.append(new)
                else:
                    drop_indices.extend(sub_indices)
                    df.at[i, "Mehrschichtiges Element"] = False
            else:
                df.at[i, "Mehrschichtiges Element"] = False
            i = j
        else:
            i += 1
            
    if drop_indices:
        df.drop(index=drop_indices, inplace=True)
        df.reset_index(drop=True, inplace=True)
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    for idx, row in df.iterrows():
        if row["Mehrschichtiges Element"]:
            for col in df.columns:
                if col.endswith(" Sub"):
                    main_col = col.replace(" Sub", "")
                    if pd.notna(row[col]) and row[col] != "":
                        df.at[idx, main_col] = row[col]

    sub_cols = [col for col in df.columns if col.endswith(" Sub")]
    df.drop(columns=sub_cols, inplace=True, errors="ignore")
    for col in ["Einzelteile", "Farbe"]:
        if col in df.columns:
            df.drop(columns=col, inplace=True)

    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = ""

    df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])]
    df.reset_index(drop=True, inplace=True)

    def remove_exact_duplicates(df):
        indices_to_drop = []
        for guid, group in df.groupby("GUID"):
            if len(group) > 1 and all(n <= 1 for n in group.nunique().values):
                indices_to_drop.extend(group.index.tolist()[1:])
        return df.drop(index=indices_to_drop).reset_index(drop=True)
    df = remove_exact_duplicates(df)

    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled, custom_chars)

    return df

def app(supplement_name, delete_enabled, custom_chars):
    st.header("Mehrschichtig Bereinigen")
    st.markdown("""
    **Einleitung:**  
    Bereinigt Excel-Dateien in mehreren Schritten:
    1. Mehrschichtigkeits-Flag setzen.
    2. Pre-Mapping der Masterspalten.
    3. Entfernen nicht klassifizierter Hauptelemente.
    4. Aufschlüsselung mehrschichtiger Elemente.
    5. Mapping der Sub-Spalten.
    6. Entfernen überflüssiger Spalten.
    7. Entfernen exakter Duplikate basierend auf GUID.
    8. Textersetzung in Mengenspalten.  
    Neuer Schritt: Entfernt den Wert "oi" in "Unter Terrain".
    """)

    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="bereinigen_file_uploader")
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")
            return

        st.subheader("Originale Daten (10 Zeilen)")
        st.dataframe(df.head(10))

        use_match_sub = st.checkbox(
            "Sub-eBKP-H identisch zur Mutter ignorieren (Toggle)",
            value=False
        )
        with st.spinner("Daten werden bereinigt ..."):
            df_clean = clean_dataframe(
                df,
                delete_enabled=delete_enabled,
                custom_chars=custom_chars,
                match_sub_toggle=use_match_sub
            )

        st.subheader("Bereinigte Daten (10 Zeilen)")
        st.dataframe(df_clean.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False)
        output.seek(0)
        file_name = f"{supplement_name.strip() or 'default'}_bereinigt.xlsx"
        st.download_button("Bereinigte Datei herunterladen", data=output,
                           file_name=file_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
