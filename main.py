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
        selected_sheet = st.selectbox("Arbeitsblatt auswählen", sheet_names, key="sheet_select")
        st.session_state["selected_sheet"] = selected_sheet

        # Vorschau des Arbeitsblatts (5 Zeilen, aus Cache)
        preview_df = load_sheet(file_bytes, selected_sheet, nrows=5)
        st.subheader("Vorschau des Arbeitsblatts (5 Zeilen)")
        st.dataframe(preview_df)

        # Header-Erkennung anhand "Teilprojekt"
        df_raw = load_sheet(file_bytes, selected_sheet, nrows=10, header=None)
        header_row = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains("Teilprojekt", case=False, na=False).any():
                header_row = idx
                break
        if header_row is None:
            st.info("Kein Header mit 'Teilprojekt' gefunden. Erste Zeile wird als Header genutzt.")
            header_row = 0
        df = load_sheet(file_bytes, selected_sheet, header=header_row)
        st.markdown(f"**Erkannter Header:** Zeile {header_row + 1}")
        st.dataframe(df.head(5))
        st.session_state["df"] = df
        st.session_state["all_columns"] = list(df.columns)

        # Toggle: Dynamisches Laden der Spaltennamen
        st.markdown("### Optionen zur Spaltenauswahl")
        dynamic_loading = st.checkbox("Dynamisches Laden der verfügbaren Spaltennamen aktivieren", 
                                      value=True, key="dynamic_loading")
        with st.expander("ℹ️ Info zur Spaltenauswahl"):
            st.write("Wenn aktiviert, werden in anderen Hierarchien bereits ausgewählte Spalten nicht mehr angezeigt. "
                     "Ist diese Option deaktiviert, werden alle Spalten angezeigt – besser für die Performance.")

        # Schritt 4: Hierarchie der Hauptmengenspalten festlegen
        st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
        for measure in st.session_state["hierarchies"]:
            if dynamic_loading:
                used_in_other = []
                for m, sel in st.session_state["hierarchies"].items():
                    if m != measure:
                        used_in_other.extend(sel)
                available_options = [col for col in st.session_state["all_columns"] if col not in used_in_other]
            else:
                available_options = st.session_state["all_columns"]
            current_selection = st.multiselect(
                f"Spalten für **{measure}** auswählen (Reihenfolge = Auswahlreihenfolge)",
                options=available_options,
                default=st.session_state["hierarchies"][measure],
                key=f"{measure}_multiselect"
            )
            st.session_state["hierarchies"][measure] = current_selection
            st.markdown(f"Ausgewählte Spalten für **{measure}**: {current_selection}")

        # Schritt 5: Merge durchführen
        if st.button("Merge und Excel herunterladen"):
            output = perform_merge(file_bytes, selected_sheet, st.session_state["hierarchies"])
            # Download-Button oberhalb der Vorschau
            st.download_button("Download Excel", data=output, file_name="merged_excel.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.subheader("Vorschau der gemergten Datei (jeweils 5 Zeilen)")
            merged_excel = pd.ExcelFile(output, engine="openpyxl")
            for sheet in merged_excel.sheet_names:
                df_preview = pd.read_excel(merged_excel, sheet_name=sheet, nrows=5, engine="openpyxl")
                st.markdown(f"**Sheet: {sheet}**")
                st.dataframe(df_preview)
