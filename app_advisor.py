import streamlit as st
import pandas as pd
import io
import openpyxl
from collections import Counter
from excel_utils import clean_columns_values, COLUMN_PRESET


def clean_value(value, delete_enabled, custom_chars):
    if isinstance(value, str):
        unwanted = [" m2", " m3", " m", "Nicht klassifiziert", "---"]
        if delete_enabled and custom_chars.strip():
            custom_list = [x.strip() for x in custom_chars.split(",") if x.strip()]
            unwanted.extend(custom_list)
        for u in unwanted:
            value = value.replace(u, "")
    return value


def detect_tool_suggestion(df: pd.DataFrame, sheetnames: list) -> tuple[str, str, str]:
    cols = df.columns.str.lower()
    confidence = "Mittel"
    reasons = []
    tools = []

    sample_rows = df.head(50).fillna("").astype(str).apply(lambda row: " ".join(row).lower(), axis=1)
    full_text = " ".join(sample_rows)

    lower_cols = df.columns.str.lower().tolist()
    has_teilprojekt = "teilprojekt" in lower_cols
    has_ebkph = "ebkp-h" in lower_cols
    has_ebkph_sub = any("ebkp-h sub" == col.lower() for col in df.columns)
    has_master_cols = any(col in lower_cols for col in ["teilprojekt", "geschoss", "unter terrain"])
    has_mengenspalten = any(term in lower_cols for term in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"])
    menge_count = sum(1 for term in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"] if term in lower_cols)

    if has_ebkph and has_ebkph_sub and has_master_cols:
        subrows = df[[c for c in df.columns if c.lower() == "ebkp-h sub"][0]].notna().sum()
        if subrows >= 1:
            tools.append("Mehrschichtig Bereinigen")
            reasons.append("eBKP-H, eBKP-H Sub und Masterspalten deuten auf eine mehrschichtige Struktur hin.")

    if has_teilprojekt and menge_count >= 2:
        tools.append("Spalten Mengen Merger")
        reasons.append("Mehrere Mengenspalten mit 'Teilprojekt' erkannt – geeignet zum Zusammenführen.")

    if len(sheetnames) > 1:
        tools.append("Master Table")
        reasons.append("Mehrere Arbeitsblätter vorhanden – ideal für die Zusammenführung in einer Master-Tabelle.")

    if not tools:
        tools.append("Merge to Table")
        reasons.append("Einfache Struktur – geeignet für das Zusammenführen mehrerer Dateien zu einer Tabelle.")
        confidence = "Niedrig"
    elif len(tools) == 1:
        confidence = "Hoch"
    else:
        confidence = "Mittel"

    return tools[0], reasons[0], confidence


def app_advisor():
    st.header("Tool-Beratung basierend auf Ihrer Excel-Datei")
    st.markdown("Laden Sie eine Beispiel-Datei hoch. Wir analysieren die Struktur und schlagen Ihnen das passende Tool vor.")

    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"], key="advisor_upload")
    if not uploaded_file:
        return

    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
        suggested_tool, reason, confidence = detect_tool_suggestion(df, xls.sheet_names)

        st.success(f"**Empfohlenes Tool:** {suggested_tool}")
        st.info(f"{reason} (Vertrauenswürdigkeit: **{confidence}**)")

        if confidence in ["Mittel", "Niedrig"]:
            st.markdown("---")
            st.markdown("### Zusätzliche Klärung durch Fragen")
            if st.checkbox("Enthält Ihre Datei Subzeilen mit 'eBKP-H Sub'?"):
                st.info("➡️ Das Tool **Mehrschichtig Bereinigen** könnte sinnvoll sein.")
            if st.checkbox("Sind Mengenspalten wie Fläche oder Volumen enthalten?"):
                st.info("➡️ Das Tool **Spalten Mengen Merger** könnte geeignet sein.")
            if st.checkbox("Enthält die Datei mehrere Arbeitsblätter mit ähnlicher Struktur?"):
                st.info("➡️ Das Tool **Master Table** könnte eine Alternative sein.")

        st.subheader("Vorschau der ersten 10 Zeilen")
        st.dataframe(df.head(10))

        with st.expander("Analyse der Strukturmerkmale"):
            lower_cols = [col.lower() for col in df.columns]
            checks = {
                "Teilprojekt": "✅" if "teilprojekt" in lower_cols else "❌",
                "eBKP-H": "✅" if "ebkp-h" in lower_cols else "❌",
                "eBKP-H Sub": "✅" if any("ebkp-h sub" == col.lower() for col in df.columns) else "❌",
                "Mengenspalten (z.B. Fläche, Volumen...)": "✅" if any(term in lower_cols for term in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]) else "❌"
            }
            for label, symbol in checks.items():
                st.write(f"{symbol} {label}")

        with st.expander("Weitere Informationen zur Datei"):
            st.markdown(f"**Anzahl Blätter:** {len(xls.sheet_names)}")
            st.markdown(f"**Spaltennamen:** {', '.join(df.columns.astype(str))}")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
