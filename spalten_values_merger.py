import streamlit as st
import pandas as pd
import io

def app(supplement_name, delete_enabled, custom_chars):
    st.header("Spalten Mengen Merger")
    st.markdown("""
    **Einleitung:**  
    Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt und definieren Sie eine Hierarchie für die Mengenspalten.
    Nach dem Merge werden die benutzten Spalten entfernt.
    """)
    
    # Session-State initialisieren
    if "df_values" not in st.session_state:
        st.session_state["df_values"] = None
    if "all_columns_values" not in st.session_state:
        st.session_state["all_columns_values"] = []
    if "hierarchies_values" not in st.session_state:
        st.session_state["hierarchies_values"] = {"Dicke": [], "Flaeche": [], "Volumen": [], "Laenge": [], "Hoehe": []}
    if "selected_sheet_values" not in st.session_state:
        st.session_state["selected_sheet_values"] = None
    if "uploaded_file_values" not in st.session_state:
        st.session_state["uploaded_file_values"] = None

    
    # Datei-Upload
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"], key="values_file_uploader")
    if uploaded_file:
        st.session_state["uploaded_file_values"] = uploaded_file
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_names = excel_file.sheet_names
            
            # Arbeitsblatt-Auswahl
            selected_sheet = st.selectbox("Arbeitsblatt wählen", sheet_names, key="values_sheet_select")
            st.session_state["selected_sheet_values"] = selected_sheet
            
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
            st.session_state["df_values"] = df
            st.session_state["all_columns_values"] = list(df.columns)


            # 🔁 Preset definieren (optional auch am Anfang des Scripts auslagern)
            PRESET_HIERARCHY = {
                "Flaeche": ["Fläche BQ", "Flaeche", "Fläche Total", "Fläche Solibri"],
                "Volumen": ["Volumen BQ", "Volumen Total", "Volumen Solibri"],
                "Laenge": ["Länge BQ", "Laenge", "Länge Solibri"],
                "Dicke": ["Dicke BQ", "Stärke", "Dicke Solibri"],
                "Hoehe": ["Höhe BQ", "Hoehe", "Höhe Solibri"]
            }
            
            # 🔁 Hierarchiewerte automatisch setzen, wenn leer
            if all(not val for val in st.session_state["hierarchies_values"].values()):
                for measure, possible_cols in PRESET_HIERARCHY.items():
                    matched_cols = [col for col in possible_cols if col in df.columns]
                    if matched_cols:
                        # Nur erste passende Spalte wählen – oder ganzes matched_cols nehmen
                        st.session_state["hierarchies_values"][measure] = [matched_cols[0]]

            
            # Hierarchie der Hauptmengenspalten festlegen
            dynamic_loading = st.checkbox("Dynamisches Laden der Spalten aktivieren", value=True, key="values_dynamic_loading")
            st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
            for measure in st.session_state["hierarchies_values"]:
                if dynamic_loading:
                    used_in_other = []
                    for m, sel in st.session_state["hierarchies_values"].items():
                        if m != measure:
                            used_in_other.extend(sel)
                    available_options = [col for col in st.session_state["all_columns_values"] if col not in used_in_other]
                else:
                    available_options = st.session_state["all_columns_values"]
                current_selection = st.multiselect(f"Spalten für {measure}", 
                                                   options=available_options,
                                                   default=st.session_state["hierarchies_values"][measure],
                                                   key=f"values_{measure}_multiselect")
                st.session_state["hierarchies_values"][measure] = current_selection
                
            if st.button("Merge und Download", key="values_merge_button"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet in sheet_names:
                        sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                        if sheet == selected_sheet:
                            for measure, hierarchy in st.session_state["hierarchies_values"].items():
                                if hierarchy:
                                    new_col = sheet_df[hierarchy[0]]
                                    for col in hierarchy[1:]:
                                        new_col = new_col.combine_first(sheet_df[col])
                                    new_name = {
                                        "Flaeche": "Fläche (m2)",
                                        "Laenge": "Länge (m)",
                                        "Dicke": "Dicke (m)",
                                        "Hoehe": "Höhe (m)",
                                        "Volumen": "Volumen (m3)"
                                    }.get(measure, measure)
                                    sheet_df[new_name] = new_col
                            used_columns = []
                            for hierarchy in st.session_state["hierarchies_values"].values():
                                used_columns.extend(hierarchy)
                            used_columns = list(set(used_columns))
                            sheet_df.drop(columns=[col for col in used_columns if col in sheet_df.columns], inplace=True)
                
                        # 🔻 Zeichenbereinigung vor dem Schreiben
                        if delete_enabled:
                            delete_chars = [" m2", " m3", " m", " kg"]
                            if custom_chars:
                                delete_chars += [c.strip() for c in custom_chars.split(",") if c.strip()]
                            for col in sheet_df.columns:
                                if sheet_df[col].dtype == object:
                                    for char in delete_chars:
                                        sheet_df[col] = sheet_df[col].str.replace(char, "", regex=False)
                
                        # 🔻 Schreiben in Excel-Datei
                        sheet_df.to_excel(writer, sheet_name=sheet, index=False)


                output.seek(0)
                st.download_button("Download Excel", data=output, file_name = f"{supplement_name.strip() or 'default'}_merged_excel.xlsx",
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
