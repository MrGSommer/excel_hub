import streamlit as st
import pandas as pd
import numpy as np
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning
)
import io


def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    """
    Schnelles Vergleichen zweier Excel-Dateien auf Basis GUID.
    Ausgabe der neuen Datei mit farblicher Hervorhebung per XlsxWriter-Conditional-Formatting.
    Zeigt Fortschritt während der Verarbeitung.
    """
    st.title("Excel Vergleichstool 📝")
    col1, col2 = st.columns(2)
    with col1:
        old_file = st.file_uploader("Alte Version (Excel)", type=["xls", "xlsx"], key="old_comp")
    with col2:
        new_file = st.file_uploader("Neue Version (Excel)", type=["xls", "xlsx"], key="new_comp")
    if not old_file or not new_file:
        st.info("Bitte beide Dateien hochladen, um den Vergleich zu starten.")
        return

    # Einlesen
    xls_old = pd.ExcelFile(old_file, engine="openpyxl")
    xls_new = pd.ExcelFile(new_file, engine="openpyxl")
    common = list(set(xls_old.sheet_names) & set(xls_new.sheet_names))
    if not common:
        st.error("Keine gemeinsamen Arbeitsblaetter gefunden.")
        return
    sheet = st.selectbox("Arbeitsblatt waehlen", common)

    @st.cache_data
    def load_and_clean(file, name):
        raw = pd.read_excel(file, sheet_name=name, header=None, engine="openpyxl")
        hdr = detect_header_row(raw)
        df = pd.read_excel(file, sheet_name=name, header=hdr, engine="openpyxl")
        return prepend_values_cleaning(df, delete_enabled, custom_chars)

    df_old = load_and_clean(old_file, sheet)
    df_new = load_and_clean(new_file, sheet)
    if "GUID" not in df_old.columns or "GUID" not in df_new.columns:
        st.error("Spalte 'GUID' nicht in beiden Tabellen gefunden.")
        return

    master_cols = [
        "Teilprojekt", "Gebäude", "Baufeld", "Geschoss",
        "eBKP-H", "Umbaustatus", "Unter Terrain", "Beschreibung",
        "Material", "Typ", "Name", "Ergänzung"
    ]
    measure_cols = ["Dicke (m)", "Fläche (m2)", "Volumen (m3)", "Länge (m)", "Höhe (m)"]
    compare = [c for c in master_cols + measure_cols if c in df_old.columns and c in df_new.columns]
    if not compare:
        st.error("Keine gemeinsamen Spalten zum Vergleichen gefunden.")
        return

    # Fortschrittsanzeige
    total_steps = len(compare) + 2
    progress = st.progress(0)
    step = 0

    # DataFrame für Vergleich vorbereiten via Merge auf GUID
    df_cmp = df_new.merge(df_old[['GUID'] + compare], on='GUID', how='left', suffixes=('', '_old'))
    step += 1
    progress.progress(step/total_steps)

    # Diff-Matrix vektorisieren
    nrows = len(df_cmp)
    diff_matrix = np.zeros((nrows, len(compare)), dtype=bool)
    for j, col in enumerate(compare):
        diff_matrix[:, j] = df_cmp[col] != df_cmp[f"{col}_old"]
        step += 1
        progress.progress(step/total_steps)
    row_mask = diff_matrix.any(axis=1)

    # Excel-Ausgabe mit XlsxWriter und Conditional Formatting
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_new.to_excel(writer, sheet_name=sheet, index=False)
        workbook  = writer.book
        worksheet = writer.sheets[sheet]

        # Formate
        grey_format = workbook.add_format({'bg_color': '#DDDDDD'})
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

        # Spaltenindex-Mapping
        col_idx = {col: i for i, col in enumerate(df_new.columns)}

        # Graue Zeilen setzen
        for r, changed in enumerate(row_mask, start=1):
            if changed:
                worksheet.set_row(r, None, grey_format)
        # Gelbe Zellen setzen
        for j, col in enumerate(compare):
            colnum = col_idx[col]
            for r, changed in enumerate(diff_matrix[:, j], start=1):
                if changed:
                    worksheet.write(r, colnum, df_new.iat[r-1, colnum], yellow_format)

    step += 1
    progress.progress(1.0)

    buffer.seek(0)
    filename = f"vergleich_{supplement_name or sheet}.xlsx"
    st.download_button(
        "Formatiertes Excel herunterladen",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("Download bereitgestellt.")
