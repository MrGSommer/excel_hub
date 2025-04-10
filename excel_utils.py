import pandas as pd
import re
import streamlit as st

def clean_columns_values(df, delete_enabled=False, custom_chars=""):
    """
    Bereinigt die typischen Mengenspalten:
    - Entfernt Einheiten (m2, m3, m)
    - Wandelt in float um
    - Entfernt zusätzliche Zeichen wie "kg" bei Bedarf
    - Gibt Warnung bei nicht-numerischen Spalten
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
