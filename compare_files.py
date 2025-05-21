import streamlit as st
import pandas as pd
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning
)
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
import io


def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    """
    Streamlit-App zum Vergleichen zweier Excel-Dateien anhand der GUID.

    Args:
        supplement_name (str): Suffix fuer Dateinamen bei Download.
        delete_enabled (bool): Ob zusaetzliche Zeichen entfernt werden.
        custom_chars (str): Kommagetrennte Liste zusaetzlicher Zeichen.
    """
    # Session-State initialisieren
    state = st.session_state
    if "uploaded_old" not in state:
        state.uploaded_old = None
    if "uploaded_new" not in state:
        state.uploaded_new = None

    # Title und Einleitung
    st.title("Excel Vergleichstool 📝")
    st.header("Dateien zum Vergleich hochladen")

    # Datei-Upload
    col1, col2 = st.columns(2)
    with col1:
        old_file = st.file_uploader(
            "Alte Version (Excel)", type=["xls", "xlsx"], key="old_comp"
        )
    with col2:
        new_file = st.file_uploader(
            "Neue Version (Excel)", type=["xls", "xlsx"], key="new_comp"
        )

    if not old_file or not new_file:
        st.info("Bitte beide Dateien hochladen, um den Vergleich zu starten.")
        return

    # Gemeinsame Sheets ermitteln
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

    # Daten laden und bereinigen
    df_old = load_and_clean(old_file, sheet)
    df_new = load_and_clean(new_file, sheet)

    # GUID-Spalte pruefen
    if "GUID" not in df_old.columns or "GUID" not in df_new.columns:
        st.error("Spalte 'GUID' nicht in beiden Tabellen gefunden.")
        return

    # Merge und Diff
    df = df_old.merge(
        df_new, on="GUID", how="outer", suffixes=("_old", "_new"), indicator=True
    )
    cols = [c for c in df_old.columns if c != "GUID"]
    changed_row = df.apply(
        lambda r: any(r[f"{c}_old"] != r[f"{c}_new"] for c in cols),
        axis=1
    )

    def diff_mask(col):
        return df[f"{col}_old"] != df[f"{col}_new"]

    # Grid-Optionen
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column()
    gb.configure_default_column(
        cellStyle=lambda params: {'backgroundColor': '#D3D3D3'} if changed_row[params['rowIndex']] else {}
    )
    for col in cols:
        gb.configure_column(
            f"{col}_new",
            cellStyle=lambda params, col=col: {'backgroundColor': 'yellow'} if diff_mask(col).iloc[params['rowIndex']] else {}
        )
    grid_opts = gb.build()

    # Ergebnis anzeigen
    st.subheader("Vergleichsergebnis")
    AgGrid(
        df,
        gridOptions=grid_opts,
        data_return_mode=DataReturnMode.FILTERED,
        update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=False,
        height=500
    )

    # Download vorbereiten
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
