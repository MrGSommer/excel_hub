import streamlit as st
import pandas as pd
import numpy as np
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning
)
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io


def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    """
    Streamlit-App zum Vergleichen zweier Excel-Dateien anhand der GUID.
    Gibt die neue Datei farblich markiert zurück: geänderte Zeilen grau, geänderte Zellen gelb.
    GUID dient als Primary Key.
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

    # Merge DataFrames
    df = df_old.merge(
        df_new, on="GUID", how="outer", suffixes=("_old", "_new"), indicator=False
    )
    # Spalten zum Vergleichen bestimmen
    compare = [col for col in master_cols + measure_cols
               if f"{col}_old" in df.columns and f"{col}_new" in df.columns]
    # Erzeuge Bool-Array für Diffs
    diffs = np.vstack([ (df[f"{c}_old"] != df[f"{c}_new"]).fillna(False).to_numpy() for c in compare ]).T
    row_changed = diffs.any(axis=1)

    # Schreibe neue DataFrame in Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_new.to_excel(writer, sheet_name=sheet, index=False)
    buffer.seek(0)

    # Lese Workbook und Style
    wb = load_workbook(buffer)
    ws = wb[sheet]
    grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Mappi Spaltennamen zu Excel-Spaltenindices
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {col: idx+1 for idx, col in enumerate(header)}

    # Style anwenden
    for row_i, changed in enumerate(row_changed, start=2):
        if not changed:
            continue
        # ganze Zeile grau füllen
        for cell in ws[row_i]:
            cell.fill = grey_fill
        # geänderte Zellen gelb füllen
        for j, col in enumerate(compare):
            if diffs[row_i-2, j]:
                excel_col = col_idx.get(col)
                if excel_col:
                    ws.cell(row=row_i, column=excel_col).fill = yellow_fill

    # Speichere und gebe Download
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    filename = f"vergleich_{supplement_name or sheet}.xlsx"
    st.download_button(
        "Formatiertes Excel herunterladen",
        data=out,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("Download bereitgestellt.")
