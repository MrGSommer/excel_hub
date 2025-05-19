import pandas as pd
import streamlit as st
import re

# Gemeinsames zentrales Preset für Hierarchie-Zuordnung und Standardnamen
COLUMN_PRESET = {
    "Fläche (m2)": ["Fläche", "Flaeche", "Fläche BQ", "Fläche Total", "Fläche Solibri"],
    "Volumen (m3)": ["Volumen", "Volumen BQ", "Volumen Total", "Volumen Solibri"],
    "Länge (m)":   ["Länge", "Laenge", "Länge BQ", "Länge Solibri"],
    "Dicke (m)":   ["Dicke", "Dicke BQ", "Stärke", "Dicke Solibri"],
    "Höhe (m)":    ["Höhe", "Hoehe", "Höhe BQ", "Höhe Solibri"]
}

def convert_size_to_m(x):
    """
    Wandelt Strings mit Einheiten (mm, cm, dm, m sowie mm2, cm2, dm2, m2, mm3, cm3, dm3, m3)
    korrekt in Meter (bzw. m2, m3) um.
    """
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    m = re.match(r"^([\d\.,]+)\s*(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m)$",
                 s, flags=re.IGNORECASE)
    if m:
        num, unit = m.groups()
        num = float(num.replace(",", "."))
        unit = unit.lower()
        if unit.endswith(("2", "3")):
            base, exp = unit[:-1], int(unit[-1])
        else:
            base, exp = unit, 1
        factor = {"mm": 0.001, "cm": 0.01, "dm": 0.1, "m": 1}[base] ** exp
        return num * factor

    # Fallback: alle Nicht-Numerischen entfernen
    cleaned = re.sub(r"[^\d\.,-]", "", s)
    try:
        return float(cleaned.replace(",", "."))
    except:
        st.warning(f"Ungültiges Format in Zelle: '{s}'")
        return pd.NA

def clean_columns_values(df, delete_enabled=False, custom_chars=""):
    """
    1) 'Nicht klassifiziert' ⇒ None
    2) Unit-Erkennung & Konvertierung in allen Spalten
    3) Ganze 0-Werte ⇒ None
    4) Zusätzliche Zeichen (kommagetrennt) löschen, falls aktiviert
    5) Warnung bei komplett leeren vordefinierten Mengenspalten
    """
    df = df.replace("Nicht klassifiziert", pd.NA)

    def try_convert(x):
        if pd.isna(x):
            return pd.NA
        if isinstance(x, str):
            return convert_size_to_m(x)
        return x

    # Anwenden auf alle Spalten
    for col in df.columns:
        df[col] = df[col].apply(try_convert)

    # Null-Werte in numerischen Spalten
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].mask(df[col] == 0, pd.NA)

    # Zusätzliche Zeichen löschen
    if delete_enabled and custom_chars:
        delete_chars = [c.strip() for c in custom_chars.split(",") if c.strip()]
        for col in df.columns:
            if df[col].dtype == object:
                for ch in delete_chars:
                    df[col] = df[col].str.replace(ch, "", regex=False)
                df[col] = df[col].mask(
                    df[col].str.strip().isin(["", "0", "0.0", "0.00"]),
                    pd.NA
                )

    # Warnung für komplett leere Mengenspalten
    empty = [c for c in COLUMN_PRESET if c in df.columns and df[c].isna().all()]
    if empty:
        st.warning("Folgende Mengenspalten leer nach Bereinigung: " + ", ".join(empty))

    return df

def detect_header_row(df_raw, suchbegriff="Teilprojekt"):
    """
    Gibt Index der ersten Zeile mit dem Suchbegriff zurück. Fallback 0.
    """
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains(suchbegriff, case=False, na=False).any():
            return idx
    return 0

def apply_preset_hierarchy(df, existing_hierarchy, preset=None):
    if preset is None:
        preset = {
            "Flaeche":  COLUMN_PRESET["Fläche (m2)"],
            "Volumen":  COLUMN_PRESET["Volumen (m3)"],
            "Laenge":   COLUMN_PRESET["Länge (m)"],
            "Dicke":    COLUMN_PRESET["Dicke (m)"],
            "Hoehe":    COLUMN_PRESET["Höhe (m)"]
        }
    for measure, defaults in existing_hierarchy.items():
        existing_hierarchy[measure] = [c for c in defaults if c in df.columns]

    if all(not v for v in existing_hierarchy.values()):
        for measure, keywords in preset.items():
            detected = []
            detected += [col for col in df.columns
                         if "bq" in col.lower() and any(k.lower() in col.lower() for k in keywords)]
            detected += [col for col in keywords if col in df.columns and col not in detected]
            detected += [col for col in df.columns
                         if any(k.lower() in col.lower() for k in keywords) and col not in detected]
            sol = [c for c in detected if "solibri" in c.lower()]
            detected = [c for c in detected if c not in sol] + sol
            if detected:
                existing_hierarchy[measure] = detected

    return existing_hierarchy

def rename_columns_to_standard(df):
    """
    Ersetzt alternative Spaltennamen durch Standardnamen laut COLUMN_PRESET.
    """
    renamed = {}
    for standard, alts in COLUMN_PRESET.items():
        matches = [c for c in df.columns
                   if any(alt.lower() in c.lower() for alt in alts)]
        if matches:
            if len(matches) > 1:
                st.warning(f"Mehrere Spalten für '{standard}': {matches}. Nutze '{matches[0]}'.")
            renamed[matches[0]] = standard
    return df.rename(columns=renamed)

def prepend_values_cleaning(df, delete_enabled=False, custom_chars=""):
    """
    1) Spalten umbenennen
    2) clean_columns_values anwenden
    """
    df = rename_columns_to_standard(df)
    return clean_columns_values(df, delete_enabled, custom_chars)
