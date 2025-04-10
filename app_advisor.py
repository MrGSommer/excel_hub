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


def detect_tool_suggestion(df: pd.DataFrame, sheetnames: list, confirmed_answers: list[str] = []) -> tuple[str, str, str]:
    cols = df.columns.str.lower()
    lower_cols = cols.tolist()
    full_text = " ".join(df.head(50).fillna("").astype(str).apply(lambda row: " ".join(row).lower(), axis=1))

    checks = {
        "has_teilprojekt": "teilprojekt" in lower_cols,
        "has_ebkp_h": "ebkp-h" in lower_cols,
        "has_ebkp_h_sub": any("ebkp-h sub" == col.lower() for col in df.columns),
        "has_master_cols": all(col in lower_cols for col in ["teilprojekt", "geschoss", "unter terrain"]),
        "menge_spalten_count": sum(1 for term in ["fläche", "flaeche", "volumen", "dicke", "länge", "laenge", "höhe", "hoehe"] if term in lower_cols),
        "num_sheets": len(sheetnames)
    }

    tool_scores = {
        "Mehrschichtig Bereinigen": 0,
        "Spalten Mengen Merger": 0,
        "Master Table": 0,
        "Merge to Table": 0
    }
    reason_list = []

    if checks["has_ebkp_h"] and checks["has_ebkp_h_sub"] and checks["has_master_cols"]:
        ebkp_sub_col = [col for col in df.columns if col.lower() == "ebkp-h sub"]
        if ebkp_sub_col and df[ebkp_sub_col[0]].notna().sum() > 0:
            tool_scores["Mehrschichtig Bereinigen"] += 3
            reason_list.append("Struktur mit eBKP-H Sub und Masterspalten erkannt.")

    if checks["has_teilprojekt"] and checks["menge_spalten_count"] >= 2:
        tool_scores["Spalten Mengen Merger"] += 2
        reason_list.append("Teilprojekt und mehrere Mengenspalten vorhanden.")

    if checks["num_sheets"] > 1:
        tool_scores["Master Table"] += 3
        reason_list.append("Mehrere Arbeitsblätter erkannt.")

    if all(score == 0 for score in tool_scores.values()):
        tool_scores["Merge to Table"] += 1
        reason_list.append("Keine komplexe Struktur erkannt.")

    question_weights = {
        "Mehrschichtig Bereinigen": ["Mehrschichtig Bereinigen"],
        "Spalten Mengen Merger": ["Spalten Mengen Merger"],
        "Master Table": ["Master Table"]
    }

    for tool, keywords in question_weights.items():
        if any(answer in keywords for answer in confirmed_answers):
            tool_scores[tool] += 2  # Mehr Einfluss durch Antwortkombinationen

    best_tool = max(tool_scores, key=tool_scores.get)
    max_score = tool_scores[best_tool]

    if max_score >= 4:
        confidence = "Hoch"
    elif max_score >= 2:
        confidence = "Mittel"
    else:
        confidence = "Niedrig"

    return best_tool, "; ".join(reason_list), confidence


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

        confirmed_tools = []
        suggested_tool, reason, confidence = detect_tool_suggestion(df, xls.sheet_names)

        if confidence in ["Mittel", "Niedrig"]:
            st.markdown("---")
            st.markdown("### Zusätzliche Klärung durch Fragen")
            follow_ups = {
                "Enthält Ihre Datei Subzeilen mit 'eBKP-H Sub'?": "Mehrschichtig Bereinigen",
                "Sind Mengenspalten wie Fläche oder Volumen enthalten?": "Spalten Mengen Merger",
                "Enthält die Datei mehrere Arbeitsblätter mit ähnlicher Struktur?": "Master Table"
            }
            for question, tool in follow_ups.items():
                if st.checkbox(question):
                    confirmed_tools.append(tool)

            suggested_tool, reason, confidence = detect_tool_suggestion(df, xls.sheet_names, confirmed_tools)

        color = {"Hoch": "🟢", "Mittel": "🟡", "Niedrig": "🔴"}[confidence]
        st.success(f"**Empfohlenes Tool:** {suggested_tool} (Vertrauenswürdigkeit: {color} **{confidence}**) ")
        st.info(f"{reason}")

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
