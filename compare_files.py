import streamlit as st
import pandas as pd
import numpy as np
from excel_utils import (
    detect_header_row,
    prepend_values_cleaning
)
import io
import logging
import traceback

# Logging-Konfiguration für Konsole
logging.basicConfig(level=logging.ERROR)
logger = logging.getLogger(__name__)

def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    """
    Vergleicht zwei Excel-Dateien anhand GUID, behandelt alle Spalten als Text (lower).
    Gibt neue Datei mit farblicher Hervorhebung zurück: graue Zeilen, gelbe Zellen.
    Zeigt Fortschritt und aktuelle Zeileneinträge.
    Fehler werden geloggt und dem Nutzer angezeigt.
    """
    st.title("Excel Vergleichstool 📝")
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

    try:
        # Einlesen und Bereinigen
        xls_old = pd.ExcelFile(old_file, engine="openpyxl")
        xls_new = pd.ExcelFile(new_file, engine="openpyxl")
        common = list(set(xls_old.sheet_names) & set(xls_new.sheet_names))
        if not common:
            st.error("Keine gemeinsamen Arbeitsblätter gefunden.")
            return
        sheet = st.selectbox("Arbeitsblatt wählen", common)

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

        # Definierte Spalten
        master_cols = [
            "Teilprojekt", "Gebäude", "Baufeld", "Geschoss",
            "eBKP-H", "Umbaustatus", "Unter Terrain", "Beschreibung",
            "Material", "Typ", "Name", "Ergänzung"
        ]
        measure_cols = [
            "Dicke (m)", "Fläche (m2)", "Volumen (m3)",
            "Länge (m)", "Höhe (m)"
        ]
        compare = [c for c in master_cols + measure_cols if c in df_old.columns and c in df_new.columns]
        if not compare:
            st.error("Keine gemeinsamen Spalten zum Vergleichen gefunden.")
            return

        # Merge DataFrames auf GUID
        df_cmp = df_new.merge(
            df_old[['GUID'] + compare],
            on='GUID', how='left', suffixes=('', '_old')
        )

        nrows = len(df_cmp)
        diffs = np.zeros((nrows, len(compare)), dtype=bool)
        status = st.empty()

        # Vergleiche alle Spalten als Text (lower)
        for j, col in enumerate(compare):
            status.text(f"Vergleiche Spalte {j+1}/{len(compare)}: {col}")
            old_series = df_cmp[f"{col}_old"].fillna('').astype(str).str.lower()
            new_series = df_cmp[col].fillna('').astype(str).str.lower()
            diffs[:, j] = new_series != old_series
        row_mask = diffs.any(axis=1)

        # Excel-Ausgabe mit XlsxWriter (nan_inf_to_errors aktiviert)
        buffer = io.BytesIO()
        writer = pd.ExcelWriter(
            buffer,
            engine='xlsxwriter',
            engine_kwargs={'options': {'nan_inf_to_errors': True}}
        )
        with writer:
            df_new.to_excel(writer, sheet_name=sheet, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet]

            grey = workbook.add_format({'bg_color': '#DDDDDD'})
            yellow = workbook.add_format({'bg_color': '#FFFF00'})
            col_idx = {col: idx for idx, col in enumerate(df_new.columns)}

            # Zeilen und Zellen formatieren
            for i, changed in enumerate(row_mask, start=1):
                status.text(f"Verarbeite Zeile {i}/{nrows}")
                if changed:
                    worksheet.set_row(i, None, grey)
                    for j, col in enumerate(compare):
                        if diffs[i-1, j]:
                            colnum = col_idx[col]
                            val = df_new.iat[i-1, colnum]
                            if not isinstance(val, (int, float, str, bool)):
                                val = str(val)
                            worksheet.write(i, colnum, val, yellow)
            status.text("Fertig mit Formatierung.")

        buffer.seek(0)
        filename = f"vergleich_{supplement_name or sheet}.xlsx"
        st.download_button(
            "Formatiertes Excel herunterladen",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.stop()
        st.success("Download bereitgestellt.")

    except Exception as e:
        # Log to console
        logger.error("Fehler in compare_tool: %s", traceback.format_exc())
        # Show to user
        st.error(f"Ein unerwarteter Fehler ist aufgetreten: {e}")
        st.exception(e)
        return
