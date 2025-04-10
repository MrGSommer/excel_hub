import streamlit as st
import pandas as pd
import io
import openpyxl
from collections import Counter
from excel_utils import clean_columns_values, rename_columns_to_standard, PRESET_RENAME


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

    if "teilprojekt" in cols:
        if any(name in cols for name in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]):
            confidence = "Hoch"
            return (
                "Spalten Mengen Merger",
                "Die Datei enthält 'Teilprojekt' und Mengenspalten wie Fläche oder Volumen. Diese eignen sich zum Zusammenführen.",
                confidence
            )

    if "ebkp-h" in cols and "ebkp-h sub" in cols:
        master_cols = ["teilprojekt", "geschoss", "unter terrain"]
        if any(col in cols for col in master_cols):
            confidence = "Hoch"
            return (
                "Mehrschichtig Bereinigen",
                "eBKP-H und eBKP-H Sub sind vorhanden, ebenso leere Masterspalten – spricht für mehrschichtige Daten.",
                confidence
            )

    if len(sheetnames) > 1:
        confidence = "Hoch"
        return (
            "Master Table",
            "Mehrere Arbeitsblätter in einer Datei erkannt – diese lassen sich zu einer Master-Tabelle zusammenführen.",
            confidence
        )

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

        st.subheader("Vorschau der ersten 10 Zeilen")
        st.dataframe(df.head(10))

        with st.expander("Analyse der Strukturmerkmale"):
            def highlight_found(colname):
                return "✅" if colname.lower() in df.columns.str.lower().tolist() else "❌"

            checks = {
                "Teilprojekt": highlight_found("Teilprojekt"),
                "eBKP-H": highlight_found("eBKP-H"),
                "eBKP-H Sub": highlight_found("eBKP-H Sub"),
                "Mengenspalten (z.B. Fläche, Volumen...)": any(col.lower() in df.columns.str.lower().tolist() for col in [
                    "fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"])
            }
            for label, found in checks.items():
                symbol = "✅" if found else "❌"
                st.write(f"{symbol} {label}")

        with st.expander("Weitere Informationen zur Datei"):
            st.markdown(f"**Anzahl Blätter:** {len(xls.sheet_names)}")
            st.markdown(f"**Spaltennamen:** {', '.join(df.columns.astype(str))}")

        with st.expander("Mögliche Umbenennung zu Standardnamen"):
            renamed = rename_columns_to_standard(df.copy())
            st.dataframe(renamed.head(5))
            st.caption("Spaltennamen wurden gemäss Preset-Namenskonvention angepasst (z.B. 'Fläche BQ' → 'Flaeche (m2)')")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
