import streamlit as st
import pandas as pd
import io
from excel_utils import detect_header_row, apply_preset_hierarchy, clean_columns_values, rename_columns_to_standard

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

    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"], key="values_file_uploader")
    if uploaded_file:
        st.session_state["uploaded_file_values"] = uploaded_file
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_names = excel_file.sheet_names
            selected_sheet = st.selectbox("Arbeitsblatt wählen", sheet_names, key="values_sheet_select")
            st.session_state["selected_sheet_values"] = selected_sheet

            preview_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5, engine="openpyxl")
            st.subheader("Vorschau (5 Zeilen)")
            st.dataframe(preview_df)

            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, engine="openpyxl")
            header_row = detect_header_row(df_raw)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine="openpyxl")
            st.markdown(f"**Erkannter Header:** Zeile {header_row+1}")
            st.dataframe(df.head(5))

            st.session_state["df_values"] = df
            st.session_state["all_columns_values"] = list(df.columns)
            st.session_state["hierarchies_values"] = apply_preset_hierarchy(df, st.session_state["hierarchies_values"])

            dynamic_loading = st.checkbox("Dynamisches Laden der Spalten aktivieren", value=True, key="values_dynamic_loading")
            st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
            for measure in st.session_state["hierarchies_values"]:
                used_in_other = []
                if dynamic_loading:
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
                            used_columns = set(col for lst in st.session_state["hierarchies_values"].values() for col in lst)
                            sheet_df.drop(columns=[col for col in used_columns if col in sheet_df.columns], inplace=True)

                            sheet_df = rename_columns_to_standard(sheet_df)
                            sheet_df = clean_columns_values(sheet_df, delete_enabled, custom_chars)

                        sheet_df.to_excel(writer, sheet_name=sheet, index=False)

                output.seek(0)
                st.download_button("Download Excel", data=output,
                                   file_name=f"{supplement_name.strip() or 'default'}_merged_excel.xlsx",
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
