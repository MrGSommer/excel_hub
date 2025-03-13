import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Operation Tools", layout="wide")
st.title("Excel Operation Tools 🚀")
st.markdown("Willkommen! Wählen Sie einen Tab für verschiedene Excel-Operationen.")

tabs = st.tabs(["Excel Merger", "Weitere Tools"])

with tabs[0]:
    st.header("Excel Merger Tool ✨")
    st.markdown(
        """
        **Einleitung:**  
        Laden Sie eine Excel-Datei hoch, wählen Sie ein Arbeitsblatt und prüfen Sie den automatisch erkannten Header (Zeile mit dem Zellwert „Teilprojekt“).  
        Legen Sie anschließend per Mehrfachauswahl die Hierarchie für die Hauptmengenspalten (Dicke, Flaeche, Volumen, Laenge, Hoehe) fest.  
        Nach dem Merge werden die benutzten Spalten entfernt.  
        **Schritte:**  
        1. Excel-Datei hochladen  
        2. Arbeitsblatt auswählen und Vorschau prüfen (5 Zeilen)  
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
        st.session_state["uploaded_file"] = uploaded_file
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_names = excel_file.sheet_names

            # Schritt 2: Arbeitsblatt auswählen
            selected_sheet = st.selectbox("Arbeitsblatt auswählen", sheet_names, key="sheet_select")
            st.session_state["selected_sheet"] = selected_sheet

            # Vorschau des Arbeitsblatts (erste 5 Zeilen)
            preview_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5, engine="openpyxl")
            st.subheader("Vorschau des Arbeitsblatts (5 Zeilen)")
            st.dataframe(preview_df)

            # Schritt 3: Header-Erkennung anhand "Teilprojekt"
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
            st.markdown(f"**Erkannter Header:** Zeile {header_row + 1}")
            st.dataframe(df.head(5))
            st.session_state["df"] = df
            st.session_state["all_columns"] = list(df.columns)
            
            # Toggle: Dynamisches Laden der verfügbaren Spaltennamen
            st.markdown("### Optionen zur Spaltenauswahl")
            dynamic_loading = st.checkbox("Dynamisches Laden der verfügbaren Spaltennamen aktivieren", value=True, key="dynamic_loading")
            with st.expander("ℹ️ Info zur Spaltenauswahl"):
                st.write("Wenn aktiviert, werden bereits in anderen Hierarchien ausgewählte Spalten nicht mehr angezeigt. "
                         "Ist diese Option deaktiviert, stehen stets alle Spalten zur Auswahl.")
            
            # Schritt 4: Hierarchie der Hauptmengenspalten festlegen
            st.markdown("### Hierarchie der Hauptmengenspalten festlegen")
            for measure in st.session_state["hierarchies"]:
                # Wenn dynamisch aktiv, filtere bereits benutzte Spalten in anderen Measures
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
            
            # Schritt 5: Merge durchführen, Download-Button und Vorschau anzeigen
            if st.button("Merge und Excel herunterladen"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet in sheet_names:
                        sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                        if sheet == selected_sheet:
                            for measure in st.session_state["hierarchies"]:
                                hierarchy = st.session_state["hierarchies"][measure]
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
                            # Entferne alle in der Hierarchie verwendeten Spalten
                            used_columns = []
                            for measure in st.session_state["hierarchies"]:
                                used_columns.extend(st.session_state["hierarchies"][measure])
                            used_columns = list(set(used_columns))
                            sheet_df.drop(columns=[col for col in used_columns if col in sheet_df.columns], inplace=True)
                        sheet_df.to_excel(writer, sheet_name=sheet, index=False)
                output.seek(0)
                
                # Download-Button oberhalb der Vorschau
                st.download_button("Download Excel", data=output, file_name="merged_excel.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.subheader("Vorschau der gemergten Datei (jeweils 5 Zeilen)")
                output.seek(0)
                merged_excel = pd.ExcelFile(output, engine="openpyxl")
                for sheet in merged_excel.sheet_names:
                    df_preview = pd.read_excel(merged_excel, sheet_name=sheet, nrows=5, engine="openpyxl")
                    st.markdown(f"**Sheet: {sheet}**")
                    st.dataframe(df_preview)
        except Exception as e:
            st.error(f"Fehler: {e}")
