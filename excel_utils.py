import pandas as pd
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
    1. Ersetzt 'Nicht klassifiziert' durch None
    2. Entfernt Mengeneinheiten (inkl. mm) und konvertiert in float
    3. Wandelt alle 0-Werte (ganzer Zelleninhalt) in None um
    4. Entfernt optionale Zeichen (z.B. ' kg')
    5. Gibt Warnungen aus, falls Spalten komplett unbestückt sind
    """
    # 1) String-Cleanup: 'Nicht klassifiziert' => None
    df = df.replace("Nicht klassifiziert", pd.NA)

    # 2) Mengenspalten bereinigen
    # Reihenfolge: mm zuerst, dann m3, m2, m
    pattern = r"\s*mm|\s*m3|\s*m2|\s*m"
    mengenspalten = list(COLUMN_PRESET.keys())
    nicht_numerisch = []

    for col in mengenspalten:
        if col in df.columns:
            # Einheit entfernen, Komma->Punkt
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(pattern, "", regex=True)
                .str.replace(",", ".")
            )
            # in numeric wandeln
            df[col] = pd.to_numeric(df[col], errors="coerce")
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
