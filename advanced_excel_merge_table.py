import streamlit as st
import pandas as pd
import io
from collections import Counter
import openpyxl

def clean_value(value, delete_enabled, custom_chars):
    if isinstance(value, str):
        unwanted = [" m2", " m3", " m", "Nicht klassifiziert", "---"]
        if delete_enabled and custom_chars.strip():
            custom_list = [x.strip() for x in custom_chars.split(",") if x.strip()]
            unwanted.extend(custom_list)
        for u in unwanted:
            value = value.replace(u, "")
    return value

def app():
    st.header("Advanced Merger - Merge to Table")
    st.markdown("Fasst mehrere Excel-Dateien zu einer Tabelle zusammen. Spalten werden nach Häufigkeit sortiert.")
    
    supplement_name = st.text_input("File Supplement Name", value="default")
    delete_enabled = st.checkbox("Zeichen in Zellen entfernen", key="table_delete")
    custom_chars = st.text_input("Zusätzliche zu löschende Zeichen (kommagetrennt)", value="", key="table_custom")
    
    uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx", "xls"], key="merge_table", accept_multiple_files=True)
    if not uploaded_files:
        return
    
    progress_bar = st.progress(0)
    merged_data = []
    column_frequency = Counter()
    total = len(uploaded_files)
    
    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            sheet = wb.active
            headers = [cell.value for cell in sheet[1]]
            if not headers:
                st.warning(f"{uploaded_file.name} hat keine Header, wird übersprungen.")
                continue
            column_frequency.update(headers)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                row_data = {col: clean_value(val, delete_enabled, custom_chars) for col, val in zip(headers, row)}
                merged_data.append(row_data)
            progress_bar.progress(idx/total)
        except Exception as e:
            st.error(f"Fehler bei {uploaded_file.name}: {e}")
            continue
    
    if not column_frequency:
        st.error("Keine gültigen Daten gefunden.")
        return
    
    sorted_columns = [col for col, _ in column_frequency.most_common()]
    df_merged = pd.DataFrame(merged_data, columns=sorted_columns)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_merged.to_excel(writer, index=False, sheet_name=supplement_name)
    output.seek(0)
    
    st.success("Merge to Table abgeschlossen.")
    st.download_button(
        "Download Merged Table Excel", 
        data=output, 
        file_name=f"{supplement_name}_merged_output_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
