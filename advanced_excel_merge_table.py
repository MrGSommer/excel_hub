import streamlit as st
import pandas as pd
import io
import openpyxl
from collections import Counter
from excel_utils import clean_columns_values, rename_columns_to_standard

def clean_value(value, delete_enabled, custom_chars):
    if isinstance(value, str):
        unwanted = [" m2", " m3", " m", "Nicht klassifiziert", "---"]
        if delete_enabled and custom_chars.strip():
            custom_list = [x.strip() for x in custom_chars.split(",") if x.strip()]
            unwanted.extend(custom_list)
        for u in unwanted:
            value = value.replace(u, "")
    return value

def app(supplement_name, delete_enabled, custom_chars):
    st.header("Merge to Table")
    st.markdown("Fasst mehrere Excel-Dateien zu einer Tabelle zusammen. Spalten werden nach Häufigkeit sortiert.")

    uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx", "xls"], key="table_files", accept_multiple_files=True)
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
            progress_bar.progress(idx / total)
        except Exception as e:
            st.error(f"Fehler bei {uploaded_file.name}: {e}")
            continue

    if not column_frequency:
        st.error("Keine gültigen Daten gefunden.")
        return

    sorted_columns = [col for col, _ in column_frequency.most_common()]
    df_merged = pd.DataFrame(merged_data, columns=sorted_columns)
    df_merged = rename_columns_to_standard(df_merged)
    df_merged = clean_columns_values(df_merged, delete_enabled, custom_chars)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_merged.to_excel(writer, index=False, sheet_name=supplement_name)
    output.seek(0)

    st.success("Merge to Table abgeschlossen.")
    st.download_button(
        "Download Merged Table Excel",
        data=output,
        file_name=f"{supplement_name}_merged_output_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="table_download_button"
    )

if __name__ == "__main__":
    app(supplement_name="default", delete_enabled=False, custom_chars="")
