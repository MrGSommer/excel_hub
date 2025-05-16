import streamlit as st
import pandas as pd
import io
import openpyxl
from collections import Counter
from excel_utils import detect_header_row, clean_columns_values, rename_columns_to_standard


def app_flow(supplement_name, delete_enabled, custom_chars):
    st.header("Flow: Mengen Spalten Merger & Master Table")
    st.markdown(
        """
        1. Mehrere Excel-Dateien hochladen
        2. Hierarchie für Mengenspalten festlegen
        3. Werte bereinigen (Einheiten & Nullen)
        4. Spalten mergen
        5. Alle Zeilen zu einer Master-Tabelle zusammenführen
        """
    )

    uploaded_files = st.file_uploader(
        "Excel-Dateien hochladen", type=["xlsx", "xls"], accept_multiple_files=True, key="flow_files"
    )
    if not uploaded_files:
        return

    # Alle Spalten sammeln
    all_columns = []
    file_sheets = {}
    for f in uploaded_files:
        wb = pd.ExcelFile(f, engine="openpyxl")
        sheet = wb.sheet_names[0]
        df_raw = pd.read_excel(f, sheet_name=sheet, header=None, engine="openpyxl")
        header_row = detect_header_row(df_raw)
        df = pd.read_excel(f, sheet_name=sheet, header=header_row, engine="openpyxl")
        file_sheets[f.name] = (f, sheet)
        all_columns.extend(df.columns.tolist())
    all_columns = list(dict.fromkeys(all_columns))

    # Hierarchie-Auswahl
    measures = ["Flaeche", "Laenge", "Dicke", "Hoehe", "Volumen"]
    hierarchies = {}
    st.markdown("### Hierarchie der Mengenspalten festlegen")
    for m in measures:
        hier = st.multiselect(
            f"Spalten für {m}", options=all_columns, key=f"flow_{m}"
        )
        hierarchies[m] = hier

    if st.button("Flow Merge & Download", key="flow_merge_button"):
        merged_data = []
        for name, (f, sheet) in file_sheets.items():
            wb = openpyxl.load_workbook(f, data_only=True)
            ws = wb[sheet]
            headers = [c.value for c in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                row_dict = dict(zip(headers, row))
                # Bereinigung: Einheiten & Nullen
                for k, v in row_dict.items():
                    row_dict[k] = clean_value(v, delete_enabled, custom_chars)
                # Spalten mergen
                for m, cols in hierarchies.items():
                    if not cols:
                        continue
                    val = None
                    for c in cols:
                        v = row_dict.get(c)
                        if v not in (None, "", 0):
                            val = v
                            break
                    new_name = {
                        "Flaeche": "Fläche (m2)",
                        "Laenge": "Länge (m)",
                        "Dicke": "Dicke (m)",
                        "Hoehe": "Hoehe (m)",
                        "Volumen": "Volumen (m3)"
                    }[m]
                    row_dict[new_name] = val
                # Ursprungs-Spalten entfernen
                used = [c for cols in hierarchies.values() for c in cols]
                for u in used:
                    row_dict.pop(u, None)
                merged_data.append(row_dict)

        df_master = pd.DataFrame(merged_data)
        df_master = rename_columns_to_standard(df_master)
        df_master = clean_columns_values(df_master, delete_enabled, custom_chars)

        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_master.to_excel(writer, index=False, sheet_name="Master")
        output.seek(0)
        st.download_button(
            "Download Master Excel",
            data=output,
            file_name=f"{supplement_name}_flow_master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def clean_value(value, delete_enabled, custom_chars):
    if isinstance(value, str):
        # Zeichen entfernen
        unwanted = [" m2", " m3", " m", "Nicht klassifiziert", "---"]
        if delete_enabled and custom_chars.strip():
            unwanted.extend([x.strip() for x in custom_chars.split(',') if x.strip()])
        for u in unwanted:
            value = value.replace(u, "")
    # Nullwerte als None
    try:
        num = float(value)
        if num == 0.0:
            return None
        return num
    except:
        return value
