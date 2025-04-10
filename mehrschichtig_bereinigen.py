import streamlit as st
import pandas as pd
import io
import re

def clean_dataframe(df, delete_enabled=False, custom_chars=""):
    master_cols = ["Teilprojekt", "Geschoss", "Unter Terrain"]
    
    # Schritt 1: Mehrschichtigkeits-Flag setzen
    df["Mehrschichtiges Element"] = df.apply(lambda row: all(pd.isna(row[col]) for col in master_cols), axis=1)
    
    # Schritt 1.5: Pre-Mapping der Masterspaltenwerte
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
    
    # Schritt 2: Entferne nicht klassifizierte Hauptelemente
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
    
    # Schritt 3: Aufschlüsselung mehrschichtiger Elemente
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
            if sub_indices:
                valid_sub_found = any(
                    pd.notna(df.at[idx, "eBKP-H Sub"]) and 
                    df.at[idx, "eBKP-H Sub"] not in ["Nicht klassifiziert", "", "Keine Zuordnung"]
                    for idx in sub_indices
                )
                if valid_sub_found:
                    drop_indices.append(i)
                    for idx in sub_indices:
                        if pd.notna(df.at[idx, "eBKP-H Sub"]) and df.at[idx, "eBKP-H Sub"] not in ["Nicht klassifiziert", "", "Keine Zuordnung"]:
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
    
    # Schritt 4: Mapping der Sub-Spalten in Hauptspalten
    for idx, row in df.iterrows():
        if row["Mehrschichtiges Element"]:
            for col in df.columns:
                if col.endswith(" Sub"):
                    main_col = col.replace(" Sub", "")
                    if pd.notna(row[col]) and row[col] != "":
                        df.at[idx, main_col] = row[col]
    
    # Schritt 5: Entferne überflüssige Spalten
    sub_cols = [col for col in df.columns if col.endswith(" Sub")]
    df.drop(columns=sub_cols, inplace=True, errors="ignore")
    for col in ["Einzelteile", "Farbe"]:
        if col in df.columns:
            df.drop(columns=col, inplace=True)
    
    # Neuer Schritt: In "Unter Terrain" den Wert "oi" entfernen
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = ""
    
    # Schritt 6: Entferne Zeilen mit "Keine Zuordnung" oder "Nicht klassifiziert" in eBKP-H
    df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])].reset_index(drop=True)
    
    # Schritt 7: Entferne exakte Duplikate basierend auf GUID
    def remove_exact_duplicates(df):
        indices_to_drop = []
        for guid, group in df.groupby("GUID"):
            if len(group) > 1:
                if all(n <= 1 for n in group.nunique().values):
                    indices_to_drop.extend(group.index.tolist()[1:])
        return df.drop(index=indices_to_drop).reset_index(drop=True)
    df = remove_exact_duplicates(df)
    
    # Schritt 8: Textersetzung in "Flaeche", "Volumen" und "Laenge" (immer aktiv)
    pattern = r'\s*m2|\s*m3|\s*m'
    for col in ["Flaeche", "Volumen", "Laenge"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(pattern, "", regex=True).str.replace(",", ".")
            df[col] = pd.to_numeric(df[col], errors="coerce")
    
    # Optional: Weitere Zeichen entfernen (z. B. "kg")
    if delete_enabled:
        delete_chars = [" kg"]
        if custom_chars:
            delete_chars += [c.strip() for c in custom_chars.split(",") if c.strip()]
        for col in df.columns:
            if df[col].dtype == object:
                for char in delete_chars:
                    df[col] = df[col].str.replace(char, "", regex=False)

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
    8. Textersetzung in "Flaeche", "Volumen" und "Laenge".  
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
        
        with st.spinner("Daten werden bereinigt ..."):
            df_clean = clean_dataframe(df, delete_enabled=delete_enabled, custom_chars=custom_chars)

        
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

