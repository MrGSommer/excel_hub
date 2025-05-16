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

    # Session‐State initialisieren
    state = st.session_state
    if "uploaded_file_values" not in state:
        state.uploaded_file_values = None
    if "sheet_names_values" not in state:
        state.sheet_names_values = []
    if "selected_sheet_values" not in state:
        state.selected_sheet_values = None
    if "header_row_values" not in state:
        state.header_row_values = None
    if "df_values" not in state:
        state.df_values = None
    if "all_columns_values" not in state:
        state.all_columns_values = []
    if "hierarchies_values" not in state:
        state.hierarchies_values = {"Dicke": [], "Flaeche": [], "Volumen": [], "Laenge": [], "Hoehe": []}

    # Upload
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"], key="values_file_uploader")
    if uploaded_file:
        # bei neuem Upload: Dateiliste und State resetten
        if state.uploaded_file_values is not uploaded_file:
            state.uploaded_file_values = uploaded_file
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            state.sheet_names_values = excel_file.sheet_names
            state.selected_sheet_values = None
            state.df_values = None
            state.all_columns_values = []
            state.hierarchies_values = {"Dicke": [], "Flaeche": [], "Volumen": [], "Laenge": [], "Hoehe": []}

        # Arbeitsblatt wählen
        selected_sheet = st.selectbox("Arbeitsblatt wählen", state.sheet_names_values, key="values_sheet_select")
        if selected_sheet and state.selected_sheet_values != selected_sheet:
            state.selected_sheet_values = selected_sheet
            # Header erkennen + DataFrame laden
            df_raw = pd.read_excel(
                state.uploaded_file_values,
                sheet_name=selected_sheet,
                header=None,
                engine="openpyxl"
            )
            header_row = detect_header_row(df_raw)
            df = pd.read_excel(
                state.uploaded_file_values,
                sheet_name=selected_sheet,
                header=header_row,
                engine="openpyxl"
            )
            state.header_row_values = header_row
            state.df_values = df
            state.all_columns_values = list(df.columns)
            state.hierarchies_values = apply_preset_hierarchy(df, state.hierarchies_values)

        # Vorschau anzeigen
        if state.df_values is not None:
            st.subheader("Vorschau (5 Zeilen)")
            st.markdown(f"**Erkannter Header:** Zeile {state.header_row_values+1}")
            st.dataframe(state.df_values.head(5))

            # Hierarchie‐Auswahl (immer dynamisch)
            st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
            for measure in state.hierarchies_values:
                # Spalten, die in anderen Measures bereits benutzt werden
                used = [c for m, cols in state.hierarchies_values.items() if m != measure for c in cols]
                options = [col for col in state.all_columns_values if col not in used]
                default = [c for c in state.hierarchies_values[measure] if c in options]
                selection = st.multiselect(
                    f"Spalten für {measure}",
                    options=options,
                    default=default,
                    key=f"values_{measure}_multiselect"
                )
                state.hierarchies_values[measure] = selection

            # Merge und Download
            if st.button("Merge und Download", key="values_merge_button"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet in state.sheet_names_values:
                        df_sheet = pd.read_excel(
                            state.uploaded_file_values,
                            sheet_name=sheet,
                            header=state.header_row_values if sheet == state.selected_sheet_values else 0,
                            engine="openpyxl"
                        )
                        if sheet == state.selected_sheet_values:
                            # Mergen
                            for measure, hierarchy in state.hierarchies_values.items():
                                if hierarchy:
                                    # erstbeste Spalte übernehmen, dann combine_first
                                    new_col = df_sheet[hierarchy[0]]
                                    for col in hierarchy[1:]:
                                        new_col = new_col.combine_first(df_sheet[col])
                                    new_name = {
                                        "Flaeche": "Fläche (m2)",
                                        "Laenge": "Länge (m)",
                                        "Dicke": "Dicke (m)",
                                        "Hoehe": "Höhe (m)",
                                        "Volumen": "Volumen (m3)"
                                    }[measure]
                                    df_sheet[new_name] = new_col
                            # benutzte Spalten entfernen
                            used_cols = {c for cols in state.hierarchies_values.values() for c in cols}
                            df_sheet.drop(columns=[c for c in used_cols if c in df_sheet.columns], inplace=True)
                            # Standardisieren & bereinigen
                            df_sheet = rename_columns_to_standard(df_sheet)
                            df_sheet = clean_columns_values(df_sheet, delete_enabled, custom_chars)

                        df_sheet.to_excel(writer, sheet_name=sheet, index=False)

                output.seek(0)
                st.download_button(
                    "Download Excel",
                    data=output,
                    file_name=f"{supplement_name.strip() or 'default'}_merged_excel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.subheader("Merge-Vorschau")
                merged_excel = pd.ExcelFile(output, engine="openpyxl")
                for sheet in merged_excel.sheet_names:
                    df_preview = pd.read_excel(merged_excel, sheet_name=sheet, nrows=5, engine="openpyxl")
                    st.markdown(f"**Sheet: {sheet}**")
                    st.dataframe(df_preview)
