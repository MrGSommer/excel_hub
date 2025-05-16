import streamlit as st
import pandas as pd
import io
from excel_utils import (
    detect_header_row,
    apply_preset_hierarchy,
    prepend_values_cleaning,
    rename_columns_to_standard
)


def app(supplement_name, delete_enabled, custom_chars):
    # Sidebar: Dynamischer Namenszusatz
    state = st.session_state
    default_supp = supplement_name
    if state.get("selected_sheet_values"):
        default_supp = state.selected_sheet_values
    elif state.get("uploaded_file_values"):
        name = state.uploaded_file_values.name
        default_supp = name.rsplit(".", 1)[0]
    supplement = st.sidebar.text_input(
        "Datei Supplement Name", value=default_supp, key="supplement_input"
    )

    st.header("Spalten Mengen Merger")
    st.markdown("""
    **Einleitung:**  
    Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt und definieren Sie eine Hierarchie für die Mengenspalten.  
    Nach dem Merge werden die benutzten Spalten entfernt.
    """)

    # Session-State initialisieren
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
    uploaded_file = st.file_uploader(
        "Excel-Datei hochladen", type=["xlsx", "xls"], key="values_file_uploader"
    )
    if uploaded_file:
        if state.uploaded_file_values is not uploaded_file:
            state.uploaded_file_values = uploaded_file
            xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
            state.sheet_names_values = xls.sheet_names
            state.selected_sheet_values = None
            state.df_values = None
            state.all_columns_values = []
            state.hierarchies_values = {"Dicke": [], "Flaeche": [], "Volumen": [], "Laenge": [], "Hoehe": []}

        # Sheet wählen
        selected_sheet = st.selectbox(
            "Arbeitsblatt wählen", state.sheet_names_values, key="values_sheet_select"
        )
        if selected_sheet and state.selected_sheet_values != selected_sheet:
            state.selected_sheet_values = selected_sheet

            # Einlesen + Header-Erkennung
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

            # Vor-Cleaning: Umbenennen + Bereinigen
            df = prepend_values_cleaning(df, delete_enabled, custom_chars)

            state.df_values = df
            state.all_columns_values = list(df.columns)
            state.hierarchies_values = apply_preset_hierarchy(df, state.hierarchies_values)

        # Vorschau
        if state.df_values is not None:
            st.subheader("Vorschau (5 Zeilen)")
            st.markdown(f"**Erkannter Header:** Zeile {state.header_row_values+1}")
            st.dataframe(state.df_values.head(5))

            # Hierarchie-Auswahl
            st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
            for measure in state.hierarchies_values:
                used = [c for m, cols in state.hierarchies_values.items() if m != measure for c in cols]
                options = [c for c in state.all_columns_values if c not in used]
                default = [c for c in state.hierarchies_values[measure] if c in options]
                sel = st.multiselect(
                    f"Spalten für {measure}",
                    options=options,
                    default=default,
                    key=f"values_{measure}_multiselect"
                )
                state.hierarchies_values[measure] = sel

            # Merge & Download
            if st.button("Merge und Download", key="values_merge_button"):
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    for sheet in state.sheet_names_values:
                        df_sheet = pd.read_excel(
                            state.uploaded_file_values,
                            sheet_name=sheet,
                            header=state.header_row_values if sheet==state.selected_sheet_values else 0,
                            engine="openpyxl"
                        )
                        if sheet == state.selected_sheet_values:
                            df_sheet = prepend_values_cleaning(df_sheet, delete_enabled, custom_chars)
                            for measure, hierarchy in state.hierarchies_values.items():
                                if hierarchy:
                                    col0 = df_sheet[hierarchy[0]]
                                    for c in hierarchy[1:]:
                                        col0 = col0.combine_first(df_sheet[c])
                                    new_name = {
                                        "Flaeche":"Fläche (m2)",
                                        "Laenge":"Länge (m)",
                                        "Dicke":"Dicke (m)",
                                        "Hoehe":"Höhe (m)",
                                        "Volumen":"Volumen (m3)"
                                    }[measure]
                                    df_sheet[new_name] = col0
                            used_cols = {c for cols in state.hierarchies_values.values() for c in cols}
                            df_sheet.drop(columns=[c for c in used_cols if c in df_sheet.columns], inplace=True)

                        df_sheet.to_excel(writer, sheet_name=sheet, index=False)

                out.seek(0)
                st.download_button(
                    "Download Excel",
                    data=out,
                    file_name=f"{supplement.strip()}_merged_excel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.subheader("Merge-Vorschau")
                merged_xl = pd.ExcelFile(out, engine="openpyxl")
                for sh in merged_xl.sheet_names:
                    df_prev = pd.read_excel(merged_xl, sheet_name=sh, nrows=5, engine="openpyxl")
                    st.markdown(f"**Sheet: {sh}**")
                    st.dataframe(df_prev)
