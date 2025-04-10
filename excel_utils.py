import pandas as pd
import re
import streamlit as st

# Gemeinsames zentrales Preset für Hierarchiezuordnung und Standardnamen
COLUMN_PRESET = {
    "Fläche (m2)": ["Fläche", "Flaeche", "Fläche BQ", "Fläche Total", "Fläche Solibri"],
    "Volumen (m3)": ["Volumen", "Volumen BQ", "Volumen Total", "Volumen Solibri"],
    "Länge (m)": ["Länge", "Laenge", "Länge BQ", "Länge Solibri"],
    "Dicke (m)": ["Dicke", "Dicke BQ", "Stärke", "Dicke Solibri"],
    "Höhe (m)": ["Höhe", "Hoehe", "Höhe BQ", "Höhe Solibri"]
}

def clean_columns_values(df, delete_enabled=False, custom_chars=""):
    """
    Entfernt Einheiten aus typischen Mengenspalten, konvertiert in float,
    entfernt optionale Zeichen, gibt Warnungen aus.
    """
    pattern = r'\s*m2|\s*m3|\s*m'
    mengenspalten = list(COLUMN_PRESET.keys())
    nicht_numerisch = []

    for col in mengenspalten:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(pattern, "", regex=True).str.replace(",", ".")
            df[col] = pd.to_numeric(df[col], errors="coerce")
            if df[col].isna().all():
                nicht_numerisch.append(col)

    if delete_enabled:
        delete_chars = [" kg"]
        if custom_chars:
            delete_chars += [c.strip() for c in custom_chars.split(",") if c.strip()]
        for col in df.columns:
            if df[col].dtype == object:
                for char in delete_chars:
                    df[col] = df[col].str.replace(char, "", regex=False)

    if nicht_numerisch:
        st.warning(f"Die folgenden Spalten enthalten keine gültigen Zahlen und wurden mit NaN ersetzt: {', '.join(nicht_numerisch)}")

    return df

def detect_header_row(df_raw, suchbegriff="Teilprojekt"):
    """
    Gibt den Index der ersten Zeile zurück, die den Suchbegriff enthält.
    Fallback ist Zeile 0.
    """
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains(suchbegriff, case=False, na=False).any():
            return idx
    return 0

def apply_preset_hierarchy(df, existing_hierarchy, preset=None):
    if preset is None:
        preset = {
            "Flaeche": COLUMN_PRESET["Fläche (m2)"],
            "Volumen": COLUMN_PRESET["Volumen (m3)"],
            "Laenge": COLUMN_PRESET["Länge (m)"],
            "Dicke": COLUMN_PRESET["Dicke (m)"],
            "Hoehe": COLUMN_PRESET["Höhe (m)"]
        }

    if all(not val for val in existing_hierarchy.values()):
        for measure, keywords in preset.items():
            detected = []
            # 1. Spalten mit 'BQ' im Namen
            detected += [col for col in df.columns if "bq" in col.lower() and any(k.lower() in col.lower() for k in keywords)]
            # 2. Exakte Matches aus Preset
            detected += [col for col in keywords if col in df.columns and col not in detected]
            # 3. Weitere Spalten mit passenden Keywords
            detected += [col for col in df.columns if any(k.lower() in col.lower() for k in keywords) and col not in detected]
            # 4. Spalten mit 'Solibri' im Namen
            solibri = [col for col in detected if "solibri" in col.lower()]
            detected = [col for col in detected if col not in solibri] + solibri

            if detected:
                existing_hierarchy[measure] = detected

    return existing_hierarchy

def rename_columns_to_standard(df):
    """
    Sucht in DataFrame nach alternativen Spaltennamen und ersetzt sie durch den definierten Standardnamen.
    Gibt eine Warnung aus, wenn mehrere Matches für denselben Standardnamen vorhanden sind.
    """
    renamed_cols = {}
    for standard_col, alternatives in COLUMN_PRESET.items():
        matches = [col for col in alternatives if col in df.columns]
        if matches:
            if len(matches) > 1:
                st.warning(f"Mehrere Spalten für '{standard_col}' gefunden: {matches}. Nur '{matches[0]}' wird verwendet.")
            renamed_cols[matches[0]] = standard_col

    df = df.rename(columns=renamed_cols)
    return df
