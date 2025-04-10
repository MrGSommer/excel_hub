import streamlit as st
import pandas as pd
import io
import openpyxl
from collections import Counter
from excel_utils import clean_columns_values, rename_columns_to_standard, COLUMN_PRESET


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

    # Zeilen analysieren (bis zu 30)
    sample_rows = df.head(30).fillna("").astype(str).apply(lambda row: " ".join(row).lower(), axis=1)
    full_text = " ".join(sample_rows)

    # Kriterien für mehrschichtige Daten (höchste Priorität)
    if any(col.lower() == "ebkp-h" for col in df.columns) and any(col.lower() == "ebkp-h sub" for col in df.columns):
        master_cols = ["teilprojekt", "geschoss", "unter terrain"]
        if any(col in cols for col in master_cols):
            subrows = df[[c for c in df.columns if c.lower() == "ebkp-h sub"][0]].notna().sum()
            if subrows >= 1:
                confidence = "Hoch"
                return (
                    "Mehrschichtig Bereinigen",
                    "eBKP-H und eBKP-H Sub sind vorhanden, ebenso Masterspalten – spricht für mehrschichtige Hierarchie.",
                    confidence
                )

    # Kriterien für Spalten Mengen Merger
    mengenbegriffe = ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]
    if "teilprojekt" in cols:
        if any(term in cols for term in mengenbegriffe) or any(term in full_text for term in mengenbegriffe):
            confidence = "Hoch"
            return (
                "Spalten Mengen Merger",
                "Die Datei enthält 'Teilprojekt' und Hinweise auf Mengenspalten wie Fläche oder Volumen. Diese eignen sich zum Zusammenführen.",
                confidence
            )

    # Master Table bei mehreren Sheets
    if len(sheetnames) > 1:
        confidence = "Hoch"
        return (
            "Master Table",
            "Mehrere Arbeitsblätter in einer Datei erkannt – diese lassen sich zu einer Master-Tabelle zusammenführen.",
            confidence
        )

    # Default-Fallback
    confidence = "Niedrig"
    return (
        "Merge to Table",
        "Einzelblattstruktur ohne spezielle Merkmale – ideal zum Zusammenführen mehrerer Dateien auf Tabellenebene.",
        confidence
    )


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
            if st.checkbox("Sind mehrere Arbeitsblätter mit gleichem Aufbau in Ihrer Datei vorhanden?"):
                st.info("➡️ Das Tool **Master Table** könnte eine passende Alternative sein.")
            if st.checkbox("Enthält Ihre Datei Subzeilen mit 'eBKP-H Sub'?"):
                st.info("➡️ Das Tool **Mehrschichtig Bereinigen** könnte sinnvoll sein.")
            if st.checkbox("Sind Mengenspalten wie Fläche oder Volumen enthalten?"):
                st.info("➡️ Das Tool **Spalten Mengen Merger** ist wahrscheinlich sinnvoll.")

        st.subheader("Vorschau der ersten 10 Zeilen")
        st.dataframe(df.head(10))

        with st.expander("Analyse der Strukturmerkmale"):
            lower_cols = [col.lower() for col in df.columns]
            checks = {
                "Teilprojekt": "✅" if "teilprojekt" in lower_cols else "❌",
                "eBKP-H": "✅" if "ebkp-h" in lower_cols else "❌",
                "eBKP-H Sub": "✅" if "ebkp-h sub" in lower_cols else "❌",
                "Mengenspalten (z.B. Fläche, Volumen...)": "✅" if any(term in lower_cols for term in [
                    "fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]) else "❌"
            }
            for label, symbol in checks.items():
                st.write(f"{symbol} {label}")

        with st.expander("Weitere Informationen zur Datei"):
            st.markdown(f"**Anzahl Blätter:** {len(xls.sheet_names)}")
            st.markdown(f"**Spaltennamen:** {', '.join(df.columns.astype(str))}")

        with st.expander("Mögliche Umbenennung zu Standardnamen"):
            renamed = rename_columns_to_standard(df.copy())
            st.dataframe(renamed.head(5))
            st.caption("Spaltennamen wurden gemäss Preset-Namenskonvention angepasst (z. B. 'Fläche BQ' → 'Fläche (m2)')")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
