import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Operation Tools", layout="wide")
st.title("Excel Operation Tools 🚀")
st.markdown("Willkommen! Wählen Sie einen Tab für verschiedene Excel-Operationen.")

tabs = st.tabs(["Excel Merger", "Weitere Tools"])

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    # Liefert die Arbeitsblattnamen der Excel-Datei
    excel_file = pd.ExcelFile(file_bytes, engine="openpyxl")
    return excel_file.sheet_names

@st.cache_data(show_spinner=False)
def load_sheet(file_bytes, sheet_name, nrows=None, header=0):
    # Lädt ein Arbeitsblatt, optional mit nrows und Header
    return pd.read_excel(file_bytes, sheet_name=sheet_name, nrows=nrows, header=header, engine="openpyxl")

@st.cache_data(show_spinner=False)
def perform_merge(file_bytes, selected_sheet, hierarchies):
    excel_file = pd.ExcelFile(file_bytes, engine="openpyxl")
    sheet_names = excel_file.sheet_names
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet in sheet_names:
            sheet_df = pd.read_excel(file_bytes, sheet_name=sheet, engine="openpyxl")
            if sheet == selected_sheet:
                for measure in hierarchies:
                    hierarchy = hierarchies[measure]
                    if hierarchy:
                        new_col = sheet_df[hierarchy[0]]
                        for col in hierarchy[1:]:
                            new_col = new_col.combine_first(sheet_df[col])
                        if measure == "Flaeche":
                            new_name = "Flaeche (m2)"
                        elif measure == "Laenge":
                            new_name = "Laenge (m)"
                        elif measure == "Dicke":
                            new_name = "Dicke (m)"
                        elif measure == "Hoehe":
                            new_name = "Hoehe (m)"
                        elif measure == "Volumen":
                            new_name = "Volumen (m3)"
                        else:
                            new_name = measure
                        sheet_df[new_name] = new_col
                # Entferne die in den Hierarchien verwendeten Spalten
                used_columns = []
                for measure in hierarchies:
                    used_columns.extend(hierarchies[measure])
                used_columns = list(set(used_columns))
                sheet_df.drop(columns=[col for col in used_columns if col in sheet_df.columns], inplace=True)
            sheet_df.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    return output

with tabs[0]:
    st.header("Excel Merger Tool ✨")
    st.markdown(
        """
        **Einleitung:**  
        Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt und prüfen Sie den automatisch erkannten Header (Zeile mit dem Zellwert „Teilprojekt“).  
        Legen Sie per Mehrfachauswahl die Hierarchie der Hauptmengenspalten (Dicke, Flaeche, Volumen, Laenge, Hoehe) fest.  
        Nach dem Merge werden die verwendeten Spalten entfernt.  
        **Schritte:**  
        1. Excel-Datei hochladen  
        2. Arbeitsblatt auswählen und Vorschau (5 Zeilen) prüfen  
        3. Header-Erkennung bestätigen  
        4. Hierarchie der Hauptmengenspalten festlegen  
        5. Merge durchführen, Datei herunterladen und Vorschau betrachten  
        """
    )

    # Session-State initialisieren
    if "df" not in st.session_state:
        st.session_state["df"] = None
    if "all_columns" not in st.session_state:
        st.session_state["all_columns"] = []
    if "hierarchies" not in st.session_state:
        st.session_state["hierarchies"] = {
            "Dicke": [],
            "Flaeche": [],
            "Volumen": [],
            "Laenge": [],
            "Hoehe": []
        }
    if "selected_sheet" not in st.session_state:
        st.session_state["selected_sheet"] = None
    if "uploaded_file" not in st.session_state:
        st.session_state["uploaded_file"] = None

    # Schritt 1: Excel-Datei hochladen
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"])
    if uploaded_file:
        # Umwandeln in Bytes für Caching
        file_bytes = uploaded_file.getvalue()
        st.session_state["uploaded_file"] = file_bytes

        # Arbeitsblätter ermitteln (aus Cache)
        sheet_names = get_sheet_names(file_bytes)
        selected_sheet = st.selectbox("Arbeitsblatt auswählen", sheet_names, key="sheet
