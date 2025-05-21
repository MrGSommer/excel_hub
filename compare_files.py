import streamlit as st
import pandas as pd
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning,
    convert_size_to_m
)
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
import io

st.set_page_config(page_title="Excel Vergleichstool", layout="wide")
st.title("Excel Vergleichstool 📝")

# Sidebar: Einstellungen
st.sidebar.header("Bereinigungseinstellungen (optional)")
delete_custom = st.sidebar.checkbox("Zusätzliche Zeichen entfernen", value=False)
custom_chars = st.sidebar.text_input("Zusätzliche Zeichen (kommagetrennt)", value="")

# Datei-Upload
st.header("Dateien zum Vergleich hochladen")
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("Alte Version (Excel)", type=["xls", "xlsx"], key="old")
with col2:
    new_file = st.file_uploader("Neue Version (Excel)", type=["xls", "xlsx"], key="new")

if not old_file or not new_file:
    st.info("Bitte beide Dateien hochladen, um den Vergleich zu starten.")
    st.stop()

# Sheet-Auswahl
old_xls = pd.ExcelFile(old_file, engine="openpyxl")
new_xls = pd.ExcelFile(new_file, engine="openpyxl")
sheets_common = list(set(old_xls.sheet_names) & set(new_xls.sheet_names))
if not sheets_common:
    st.error("Keine gemeinsamen Arbeitsblätter gefunden.")
    st.stop()
sheet = st.selectbox("Arbeitsblatt wählen", sheets_common)

# Einlesen mit Header-Erkennung
@st.cache_data
def load_and_clean(uploaded, sheet_name):
    df_raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row = detect_header_row(df_raw)
    df = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_row, engine="openpyxl")
    # Grund-Bereinigung
    df = prepend_values_cleaning(df, delete_enabled=delete_custom, custom_chars=custom_chars)
    return df

df_old = load_and_clean(old_file, sheet)
df_new = load_and_clean(new_file, sheet)

# Sicherstellen, dass GUID-Spalte besteht
if "GUID" not in df_old.columns or "GUID" not in df_new.columns:
    st.error("Spalte 'GUID' nicht in beiden Tabellen gefunden.")
    st.stop()

# Merge old/new
df = df_old.merge(df_new, on="GUID", how="outer", suffixes=("_old", "_new"), indicator=True)

def diff_mask(col):
    return df[f"{col}_old"] != df[f"{col}_new"]

# Zeilen, in denen sich GUID-Mutterspalten oder Mengenspalten ändern
master_cols = [c for c in df_old.columns if c != "GUID"]
changed_row = df.apply(lambda r: any(r[f"{c}_old"] != r[f"{c}_new"] for c in master_cols), axis=1)

# AgGrid Styling
gb = GridOptionsBuilder.from_dataframe(df)

def style_cells():
    # Zeilenshading
    gb.configure_default_column(cellStyle=lambda params: {'backgroundColor': '#D3D3D3'} if changed_row[params['rowIndex']] else {})
    # Zellen-Markierung für geänderte Werte
    for col in master_cols:
        col_new = f"{col}_new"
        gb.configure_column(
            col_new,
            cellStyle=lambda params, col=col: {'backgroundColor': 'yellow'} 
                if diff_mask(col).iloc[params['rowIndex']] else {}
        )

style_cells()

grid_options = gb.build()
st.subheader("Vergleichsergebnis")
AgGrid(
    df,
    gridOptions=grid_options,
    data_return_mode=DataReturnMode.FILTERED,
    update_mode=GridUpdateMode.NO_UPDATE,
    enable_enterprise_modules=False,
    height=500,
)

# Download der markierten Excel-Datei
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name=sheet, index=False)
out.seek(0)

st.download_button(
    "Markierte Excel herunterladen",
    data=out,
    file_name=f"vergleich_{sheet}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
