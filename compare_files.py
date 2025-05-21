import streamlit as st
import pandas as pd
import numpy as np
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning
)
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
import io


def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    """
    Streamlit-App zum Vergleichen zweier Excel-Dateien anhand der GUID.
    Vergleicht definierte Master- und Measure-Spalten nur, wenn in beiden Dateien vorhanden.
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
    measure_cols = ["Dicke (m)", "Fläche (m2)", "Volumen (m3)", "Länge (m)", "Höhe (m)" ]

    df = df_old.merge(
        df_new, on="GUID", how="outer", suffixes=("_old", "_new"), indicator=True
    )
    # Spalten zum Vergleichen bestimmen
    compare = [col for col in master_cols + measure_cols
               if f"{col}_old" in df.columns and f"{col}_new" in df.columns]

    # Diff-Indikator
    diffs = [df[f"{c}_old"] != df[f"{c}_new"] for c in compare]
    df['__changed'] = np.logical_or.reduce(diffs) if diffs else False

    # Grid-Optionen
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column()

    # Zeilen grau einfärben bei Änderungen
    get_row_style = (
        "function(params) {"
        "return params.data.__changed ? {backgroundColor: '#D3D3D3'} : {};"
        "}"
    )
    gb.configure_grid_options(getRowStyle=get_row_style)

    # Zellen gelb einfärben für verglichene Spalten
    for col in compare:
        js = (
            "function(params) {"
            f"return params.data['{col}_old'] !== params.data['{col}_new']"
            " ? {backgroundColor: 'yellow'} : {};"
            "}"
        )
        gb.configure_column(f"{col}_new", cellStyleJs=js)

    grid_opts = gb.build()

    st.subheader("Vergleichsergebnis")
    AgGrid(
        df,
        gridOptions=grid_opts,
        data_return_mode=DataReturnMode.FILTERED,
        update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=False,
        height=500
    )

    # Download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
    buffer.seek(0)
    filename = f"vergleich_{supplement_name or sheet}.xlsx"
    st.download_button(
        "Markierte Excel herunterladen",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
