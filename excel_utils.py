import pandas as pd
import re
import streamlit as st

def clean_columns_values(df, delete_enabled=False, custom_chars=""):
    """
    Entfernt Einheiten aus typischen Mengenspalten, konvertiert in float,
    entfernt optionale Zeichen, gibt Warnungen aus.
    """
    pattern = r'\s*m2|\s*m3|\s*m'
    mengenspalten = ["Flaeche (m2)", "Volumen (m3)", "Laenge (m)", "Dicke (m)", "Hoehe (m)"]
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
    """
    Setzt die hierarchischen Spalten automatisch anhand eines Presets,
    wenn noch keine Auswahl vorhanden ist.
    """
    if preset is None:
        preset = {
            "Flaeche": ["Fläche", "Fläche BQ", "Flaeche", "Fläche Total", "Fläche Solibri"],
            "Volumen": ["Volumen", "Volumen BQ", "Volumen Total", "Volumen Solibri"],
            "Laenge": ["Länge", "Länge BQ", "Laenge", "Länge Solibri"],
            "Dicke": ["Dicke", "Dicke BQ", "Stärke", "Dicke Solibri"],
            "Hoehe": ["Höhe", "Höhe BQ", "Hoehe", "Höhe Solibri"]
        }

    if all(not val for val in existing_hierarchy.values()):
        for measure, possible_cols in preset.items():
            matched_cols = [col for col in possible_cols if col in df.columns]
            ordered_matches = [col for col in possible_cols if col in matched_cols]
            if ordered_matches:
                existing_hierarchy[measure] = ordered_matches
    return existing_hierarchy
