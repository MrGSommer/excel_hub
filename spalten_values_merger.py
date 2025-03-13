import streamlit as st
import pandas as pd
import io

def app():
    st.header("Spalten Values Merger")
    st.markdown("""
    **Einleitung:**  
    Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt und prüfen Sie den erkannten Header.
    Legen Sie eine Hierarchie für die Hauptmengenspalten fest. Nach dem Merge werden die benutzten Spalten entfernt.
    """)
    
    # Session-State initialisieren
    if "df" not in st.session_state:
        st.session_state["df"] = None
    if "all_columns" not in st.session_state:
        st.session_state["all_columns"] = []
    if "hierarchies" not in st.session_state:
        st.session_state["hierarchies"] = {"Dicke": [], "Flaeche": [], "Volumen": [], "Laenge": [], "Hoehe": []}
    if "selected_sheet" not in st.session_state:
        st.session_state["selected_sheet"] = None
    if "uploaded_file" not in st.session_state:
        st.session_state["uploaded_file"] = None

    # Datei-Upload
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"])
    if uploaded_file:
        st.session_state["uploaded_file"] = uploaded_file
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_names = excel_file.sheet_names
            
            # Arbeitsblatt-Auswahl
            selected_sheet = st.selectbox("Arbeitsblatt wählen", sheet_names, key="sheet_select")
            st.session_state["selected_sheet"] = selected_sheet
            
            # Vorschau
            preview_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5, engine="openpyxl")
            st.subheader("Vorschau (5 Zeilen)")
            st.dataframe(preview_df)
            
            # Header-Erkennung anhand "Teilprojekt"
            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, engine="openpyxl")
            header_row = None
            for idx, row in df_raw.iterrows():
                if row.astype(str).str.contains("Teilprojekt", case=False, na=False).any():
                    header_row = idx
                    break
            if header_row is None:
                st.info("Kein Header mit 'Teilprojekt' gefunden. Erste Zeile wird als Header genutzt.")
                header_row = 0
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine="openpyxl")
            st.markdown(f"**Erkannter Header:** Zeile {header_row+1}")
            st.dataframe(df.head(5))
            st.session_state["df"] = df
            st.session_state["all_columns"] = list(df.columns)
            
            # Hierarchie der Hauptmengenspalten festlegen
            dynamic_loading = st.checkbox("Dynamisches Laden der Spalten aktivieren", value=True, key="dynamic_loading")
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
                current_selection = st.multiselect(f"Spalten für {measure}", 
                                                   options=available_options,
                                                   default=st.session_state["hierarchies"][measure],
                                                   key=f"{measure}_multiselect")
                st.session_state["hierarchies"][measure] = current_selection
                
            # Merge durchführen und Download anbieten
            if st.button("Merge und Download"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet in sheet_names:
                        sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                        if sheet == selected_sheet:
                            for measure, hierarchy in st.session_state["hierarchies"].items():
                                if hierarchy:
                                    new_col = sheet_df[hierarchy[0]]
                                    for col in hierarchy[1:]:
                                        new_col = new_col.combine_first(sheet_df[col])
                                    new_name = {
                                        "Flaeche": "Flaeche (m2)",
                                        "Laenge": "Laenge (m)",
                                        "Dicke": "Dicke (m)",
                                        "Hoehe": "Hoehe (m)",
                                        "Volumen": "Volumen (m3)"
                                    }.get(measure, measure)
                                    sheet_df[new_name] = new_col
                            used_columns = []
                            for hierarchy in st.session_state["hierarchies"].values():
                                used_columns.extend(hierarchy)
                            used_columns = list(set(used_columns))
                            sheet_df.drop(columns=[col for col in used_columns if col in sheet_df.columns], inplace=True)
                        sheet_df.to_excel(writer, sheet_name=sheet, index=False)
                output.seek(0)
                st.download_button("Download Excel", data=output, file_name="merged_excel.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.subheader("Merge-Vorschau")
                output.seek(0)
                merged_excel = pd.ExcelFile(output, engine="openpyxl")
                for sheet in merged_excel.sheet_names:
                    df_preview = pd.read_excel(merged_excel, sheet_name=sheet, nrows=5, engine="openpyxl")
                    st.markdown(f"**Sheet: {sheet}**")
                    st.dataframe(df_preview)
        except Exception as e:
            st.error(f"Fehler: {e}")
