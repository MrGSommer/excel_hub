import streamlit as st
import pandas as pd
import io
from excel_utils import (
    detect_header_row,
    apply_preset_hierarchy,
    prepend_values_cleaning,
    rename_columns_to_standard,
    convert_size_to_m,
    convert_quantity_columns
)

def app(supplement_name, delete_enabled, custom_chars):
    # Datei-Supplement aus main.py übernehmen, sonst Sheet- oder Dateiname
    state = st.session_state
    supplement = supplement_name or (
        state.get("selected_sheet_values")
        or (state.uploaded_file_values.name.rsplit(".", 1)[0]
            if state.get("uploaded_file_values") else "")
    )

    st.header("Spalten Mengen Merger")
    st.markdown("""
    **Einleitung:**  
    Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt.  
    Anschliessend sehen Sie erst das Original, dann die bereinigte Version, bevor Sie die Hierarchie festlegen.
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
    if not uploaded_file:
        return

    # Bei neuem Upload State zurücksetzen
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
    if not selected_sheet or selected_sheet == state.selected_sheet_values:
        return
    state.selected_sheet_values = selected_sheet

    # 1) Original einlesen
    df_raw = pd.read_excel(
        state.uploaded_file_values,
        sheet_name=selected_sheet,
        header=None,
        engine="openpyxl"
    )
    header_row = detect_header_row(df_raw)
    df_original = pd.read_excel(
        state.uploaded_file_values,
        sheet_name=selected_sheet,
        header=header_row,
        engine="openpyxl"
    )
    state.header_row_values = header_row

    st.subheader("Originale Daten (5 Zeilen)")
    st.markdown(f"**Erkannter Header:** Zeile {header_row+1}")
    st.dataframe(df_original.head(5))

    # 2) Grund-Bereinigung
    df_clean = prepend_values_cleaning(df_original, delete_enabled, custom_chars)
    st.subheader("Bereinigte Daten (5 Zeilen)")
    st.dataframe(df_clean.head(5))

    state.df_values = df_clean
    state.all_columns_values = list(df_clean.columns)
    state.hierarchies_values = apply_preset_hierarchy(df_clean, state.hierarchies_values)

    # 3) Hierarchie-Auswahl
    st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
    master_cols = [
        "Teilprojekt", "Gebäude", "Baufeld", "Geschoss",
        "Unter Terrain", "eBKP-H", "Umbaustatus", "Material", "ING", "Anzahl", "GUID"
    ]
    for measure in state.hierarchies_values:
        used = [
            c for m, cols in state.hierarchies_values.items()
            if m != measure for c in cols
        ]
        options = [
            c for c in state.all_columns_values
            if c not in used and c not in master_cols
        ]
        default = [c for c in state.hierarchies_values[measure] if c in options]
        sel = st.multiselect(
            f"Spalten für {measure}",
            options=options,
            default=default,
            key=f"values_{measure}_multiselect"
        )
        state.hierarchies_values[measure] = sel

    # 4) Merge & Download
    if st.button("Merge und Download", key="values_merge_button"):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for sheet in state.sheet_names_values:
                df_sheet = pd.read_excel(
                    state.uploaded_file_values,
                    sheet_name=sheet,
                    header=state.header_row_values if sheet == state.selected_sheet_values else 0,
                    engine="openpyxl"
                )
                if sheet == state.selected_sheet_values:
                    # Grund-Bereinigung auf das Sheet
                    df_sheet = prepend_values_cleaning(df_sheet, delete_enabled, custom_chars)


                    # 4.x) Quell-Spalten säubern & konvertieren, bevor das Merging startet
                    for measure, hierarchy in state.hierarchies_values.items():
                        for src in hierarchy:
                            if src in df_sheet.columns:
                                # optionale Custom-Chars löschen
                                if delete_enabled and custom_chars:
                                    for ch in custom_chars.split(","):
                                        df_sheet[src] = (
                                            df_sheet[src]
                                            .astype(str)
                                            .str.replace(ch, "", regex=False)
                                        )
                                # Einheitserkennung & -konvertierung (0 mm ⇒ pd.NA)
                                df_sheet[src] = df_sheet[src].apply(convert_size_to_m)
                    
                    # 4.1) Erzeugen und konvertieren der gemergten Spalten
                    for measure, hierarchy in state.hierarchies_values.items():
                        if hierarchy:
                            col0 = df_sheet[hierarchy[0]]
                            for c in hierarchy[1:]:
                                col0 = col0.combine_first(df_sheet[c])
                            new_name = {
                                "Flaeche": "Fläche (m2)",
                                "Laenge":  "Länge (m)",
                                "Dicke":   "Dicke (m)",
                                "Hoehe":   "Höhe (m)",
                                "Volumen": "Volumen (m3)"
                            }[measure]
                            # direkt mit convert_size_to_m sauber machen
                            df_sheet[new_name] = col0.apply(convert_size_to_m)


                    # 4.2) Reorder: neue Spalten am kleinsten Index der Quellen einfügen
                    cols_before = list(df_sheet.columns)
                    src_cols = [
                        src for hierarchy in state.hierarchies_values.values()
                        for src in hierarchy if src in cols_before
                    ]
                    indices = [cols_before.index(c) for c in src_cols] if src_cols else []
                    if indices:
                        insert_at = min(indices)
                        merged_names = [
                            name for name in
                            ("Fläche (m2)", "Länge (m)", "Dicke (m)", "Höhe (m)", "Volumen (m3)")
                            if name in df_sheet.columns
                        ]
                        # Drop aller Quell-Spalten
                        df_sheet.drop(columns=src_cols, inplace=True)
                        # Neuordnung
                        cols_after = list(df_sheet.columns)
                        for name in merged_names:
                            cols_after.remove(name)
                        for i, name in enumerate(merged_names):
                            cols_after.insert(insert_at + i, name)
                        df_sheet = df_sheet[cols_after]
                    else:
                        # Drop, wenn keine Quellen
                        df_sheet.drop(
                            columns=[c for cols in state.hierarchies_values.values() for c in cols
                                     if c in df_sheet.columns],
                            inplace=True
                        )
                        
                df_sheet_export = convert_quantity_columns(df_sheet.copy())
                df_sheet_export.to_excel(writer, sheet_name=sheet, index=False)

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
