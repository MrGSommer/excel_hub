import pandas as pd
import streamlit as st
import re

# Gemeinsames zentrales Preset für Hierarchiezuordnung und Standardnamen
COLUMN_PRESET = {
    "Fläche (m2)": ["Fläche", "Flaeche", "Fläche BQ", "Fläche Total", "Fläche Solibri"],
    "Volumen (m3)": ["Volumen", "Volumen BQ", "Volumen Total", "Volumen Solibri"],
    "Länge (m)": ["Länge", "Laenge", "Länge BQ", "Länge Solibri"],
    "Dicke (m)": ["Dicke", "Dicke BQ", "Stärke", "Dicke Solibri"],
    "Höhe (m)": ["Höhe", "Hoehe", "Höhe BQ", "Höhe Solibri"]
}


def convert_size_to_m(x):
    """
    Wandelt Strings mit Einheiten (mm, cm, dm, m sowie mm2, cm2, dm2, m2, mm3, cm3, dm3, m3)
    korrekt in Meter (bzw. m2, m3) um.
    """
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    m = re.match(r"^([\d\.,]+)\s*(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m)$", s, flags=re.IGNORECASE)
    if m:
        num, unit = m.groups()
        num = float(num.replace(",", "."))
        unit = unit.lower()
        # Exponent ermitteln
        if unit.endswith(("2", "3")):
            base, exp = unit[:-1], int(unit[-1])
        else:
            base, exp = unit, 1
        factor = {"mm": 0.001, "cm": 0.01, "dm": 0.1, "m": 1}[base] ** exp
        
        return num * factor

    # Fallback: alles Nicht-Numerische entfernen
    cleaned = re.sub(r"[^\d\.,-]", "", s)
    try:
        val = float(cleaned.replace(",", "."))
    except:
        val = pd.NA

    # WARNUNG für komplett nicht erkannte Formate
    if not m and val is pd.NA:
        st.warning(f"Ungültiges Format in Zelle: '{s}'")

    return val


def clean_columns_values(df, delete_enabled=False, custom_chars=""):
    """
    1. Ersetzt 'Nicht klassifiziert' durch None
    2. Entfernt Mengeneinheiten (inkl. mm) und konvertiert in float
    3. Wandelt alle 0-Werte (ganzer Zelleninhalt) in None um
    4. Entfernt optionale Zeichen (z.B. ' kg')
    5. Gibt Warnungen aus, falls Spalten komplett unbestückt sind
    """
    # 1) String-Cleanup: 'Nicht klassifiziert' => None
    df = df.replace("Nicht klassifiziert", pd.NA)

    # 2) Mengenspalten bereinigen
    mengenspalten = list(COLUMN_PRESET.keys())
    nicht_numerisch = []
    for col in mengenspalten:
        if col in df.columns:
            df[col] = df[col].apply(convert_size_to_m)
            # 3) Ganze 0-Werte -> None
            df[col] = df[col].mask(df[col] == 0, pd.NA)
            if df[col].isna().all():
                nicht_numerisch.append(col)


    # 4) Optionale Zeichen löschen (nur für Textspalten)
    if delete_enabled:
        delete_chars = [" kg"]
        if custom_chars:
            delete_chars += [c.strip() for c in custom_chars.split(",") if c.strip()]
        for col in df.columns:
            if df[col].dtype == object:
                # Zeichen entfernen
                for char in delete_chars:
                    df[col] = df[col].str.replace(char, "", regex=False)
                # Ganze Null-Strings zu None
                df[col] = df[col].mask(
                    df[col].str.strip().isin(["0", "0.0", "0.00", "0 mm"]),
                    pd.NA
                )

    # 5) Warnung bei komplett leeren Mengenspalten
    if nicht_numerisch:
        st.warning(
            "Die folgenden Mengenspalten wurden komplett in None umgewandelt: "
            f"{', '.join(nicht_numerisch)}"
        )

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

    # Vorhandene Defaults filtern
    for measure, defaults in existing_hierarchy.items():
        existing_hierarchy[measure] = [c for c in defaults if c in df.columns]

    # Nur wenn komplett leer: Preset anwenden
    if all(not vals for vals in existing_hierarchy.values()):
        for measure, keywords in preset.items():
            detected = []
            detected += [
                col for col in df.columns
                if "bq" in col.lower() and any(k.lower() in col.lower() for k in keywords)
            ]
            detected += [
                col for col in keywords
                if col in df.columns and col not in detected
            ]
            detected += [
                col for col in df.columns
                if any(k.lower() in col.lower() for k in keywords) and col not in detected
            ]
            sol = [c for c in detected if "solibri" in c.lower()]
            detected = [c for c in detected if c not in sol] + sol

            if detected:
                existing_hierarchy[measure] = detected

    return existing_hierarchy


def rename_columns_to_standard(df):
    """
    Ersetzt alternative Spaltennamen durch Standardnamen laut COLUMN_PRESET.
    Warnt bei Mehrfachmatches.
    """
    renamed = {}
    for standard, alts in COLUMN_PRESET.items():
        matches = [c for c in alts if c in df.columns]
        if matches:
            if len(matches) > 1:
                st.warning(
                    f"Mehrere Spalten für '{standard}' gefunden: {matches}. Verwende '{matches[0]}'."
                )
            renamed[matches[0]] = standard

    return df.rename(columns=renamed)


def prepend_values_cleaning(df, delete_enabled=False, custom_chars=""):
    """
    Hilfsfunktion: Standardisiert Spaltennamen und bereinigt Werte.
    1. Spalten umbenennen
    2. Werte mit clean_columns_values bereinigen
    """
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled, custom_chars)
    return df
