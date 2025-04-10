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
    lower_cols = cols.tolist()
    full_text = " ".join(df.head(50).fillna("").astype(str).apply(lambda row: " ".join(row).lower(), axis=1))

    tool_flags = {
        "Spalten Mengen Merger": False,
        "Mehrschichtig Bereinigen": False,
        "Master Table": False,
        "Merge to Table": False
    }
    reasons = []

    menge_terms = ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]
    menge_count = sum(1 for term in menge_terms if term in lower_cols)

    has_teilprojekt = "teilprojekt" in lower_cols
    has_ebkph = "ebkp-h" in lower_cols
    has_ebkph_sub = any("ebkp-h sub" == col.lower() for col in df.columns)
    has_master_cols = all(col in lower_cols for col in ["teilprojekt", "geschoss", "unter terrain"])
    num_sheets = len(sheetnames)

    if has_ebkph and has_ebkph_sub and has_master_cols:
        sub_col = [col for col in df.columns if col.lower() == "ebkp-h sub"]
        if sub_col and df[sub_col[0]].notna().sum() > 0:
            tool_flags["Mehrschichtig Bereinigen"] = True
            reasons.append("Struktur mit eBKP-H Sub und Masterspalten erkannt.")

    if has_teilprojekt and menge_count >= 2:
        tool_flags["Spalten Mengen Merger"] = True
        reasons.append("Teilprojekt und mehrere Mengenspalten vorhanden.")

    if num_sheets > 1:
        tool_flags["Master Table"] = True
        reasons.append("Mehrere Arbeitsblätter erkannt.")

    if not any(tool_flags.values()):
        tool_flags["Merge to Table"] = True
        reasons.append("Keine komplexe Struktur erkannt.")

    confidence = "Hoch" if sum(tool_flags.values()) == 1 else ("Mittel" if sum(tool_flags.values()) == 2 else "Niedrig")

    primary_tool = [tool for tool, valid in tool_flags.items() if valid][0]
    return primary_tool, "; ".join(reasons), confidence


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

        color = {"Hoch": "🟢", "Mittel": "🟡", "Niedrig": "🔴"}[confidence]
        st.success(f"**Empfohlenes Tool:** {suggested_tool}")
        st.info(f"{reason} (Vertrauenswürdigkeit: {color} **{confidence}**) ")

        if confidence in ["Mittel", "Niedrig"]:
            st.markdown("---")
            st.markdown("### Zusätzliche Klärung durch Fragen")
            follow_ups = {
                "Enthält Ihre Datei Subzeilen mit 'eBKP-H Sub'?": "Mehrschichtig Bereinigen",
                "Sind Mengenspalten wie Fläche oder Volumen enthalten?": "Spalten Mengen Merger",
                "Enthält die Datei mehrere Arbeitsblätter mit ähnlicher Struktur?": "Master Table"
            }
            confirmed_tools = []
            for question, tool in follow_ups.items():
                if st.checkbox(question):
                    confirmed_tools.append(tool)

            if confirmed_tools:
                st.markdown(f"🔍 Basierend auf Ihren Antworten könnte eines dieser Tools passend sein: **{', '.join(confirmed_tools)}**")

        st.subheader("Vorschau der ersten 10 Zeilen")
        st.dataframe(df.head(10))

        with st.expander("Analyse der Strukturmerkmale"):
            lower_cols = [col.lower() for col in df.columns]
            checks = {
                "Teilprojekt": "✅" if "teilprojekt" in lower_cols else "❌",
                "eBKP-H": "✅" if "ebkp-h" in lower_cols else "❌",
                "eBKP-H Sub": "✅" if any("ebkp-h sub" == col.lower() for col in df.columns) else "❌",
                "Mengenspalten (z.B. Fläche, Volumen...)": "✅" if any(term in lower_cols for term in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]) else "❌",
                "Anzahl Arbeitsblätter > 1": "✅" if len(xls.sheet_names) > 1 else "❌"
            }
            for label, symbol in checks.items():
                st.write(f"{symbol} {label}")

        with st.expander("Weitere Informationen zur Datei"):
            st.markdown(f"**Anzahl Blätter:** {len(xls.sheet_names)}")
            st.markdown(f"**Spaltennamen:** {', '.join(df.columns.astype(str))}")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
