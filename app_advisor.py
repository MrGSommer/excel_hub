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

    # === Mehrschichtig Bereinigen ===
    if "ebkp-h" in cols and "ebkp-h sub" in cols:
        master_cols = ["teilprojekt", "geschoss", "unter terrain"]
        if any(col in cols for col in master_cols):
            subrows = df[df["eBKP-H Sub"].notna()].shape[0] if "eBKP-H Sub" in df.columns else 0
            if subrows >= 1:
                return (
                    "Mehrschichtig Bereinigen",
                    "eBKP-H und eBKP-H Sub sind vorhanden, ebenso Masterspalten – spricht für mehrschichtige Hierarchie.",
                    "Hoch"
                )

    # === Spalten Mengen Merger ===
    mengenbegriffe = ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"]
    mengen_spalten = [col for col in mengenbegriffe if any(col in c for c in cols)]
    if "teilprojekt" in cols and len(mengen_spalten) >= 2:
        return (
            "Spalten Mengen Merger",
            "Mehrere gleichartige Mengenspalten (z. B. Fläche BQ, Fläche Total) erkannt – Zusammenführung empfohlen.",
            "Hoch"
        )

    # === Master Table ===
    if len(sheetnames) > 1:
        return (
            "Master Table",
            "Mehrere Arbeitsblätter in einer Datei erkannt – geeignet für einen Masterzusammenzug.",
            "Hoch"
        )

    # === Merge to Table ===
    return (
        "Merge to Table",
        "Einzelblatt ohne spezifische Strukturen – geeignet für allgemeine Tabellenzusammenführung.",
        "Mittel"
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
            st.caption("Spaltennamen wurden gemäss Preset-Namenskonvention angepasst (z. B. 'Fläche BQ' → 'Fläche (m2)')")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
