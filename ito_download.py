import streamlit as st
import zipfile
import io
import os

def app():
    st.header("Solibri ITO-Vorlagen herunterladen")
    st.markdown("""
    Hier finden Sie vordefinierte **Quantity Take-Offs (ITO-Dateien)** für den Einsatz in **Solibri**.
    Diese Vorlagen können direkt importiert und bei Bedarf mit den Tools dieser Plattform kombiniert werden.
    """)

    ito_files = {
        "Mehrschichtig": [
            "ito_templates/ebkph_takeoff.ito",
            "ito_templates/ebkph_sub.ito"
        ],
        "SIA 416": [
            "ito_templates/flaeche_volumen.ito"
        ],
        "Bauteilkategorien (Elementtypen)": [
            "ito_templates/elementtypen.ito",
            "ito_templates/elementtypen_bkp.ito"
        ]
    }

    selected = st.selectbox("Vorlage auswählen", list(ito_files.keys()))
    file_paths = ito_files[selected]

    if len(file_paths) == 1:
        path = file_paths[0]
        try:
            with open(path, "rb") as f:
                st.download_button(
                    label=f"📥 {os.path.basename(path)} herunterladen",
                    data=f,
                    file_name=os.path.basename(path),
                    mime="application/octet-stream"
                )
        except FileNotFoundError:
            st.error(f"Datei nicht gefunden: {path}")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for path in file_paths:
                try:
                    with open(path, "rb") as f:
                        zip_file.writestr(os.path.basename(path), f.read())
                except FileNotFoundError:
                    st.warning(f"Fehlende Datei: {path}")
        zip_buffer.seek(0)
        st.download_button(
            label=f"📦 Alle ITOs für '{selected}' als ZIP herunterladen",
            data=zip_buffer,
            file_name=f"{selected.replace(' ', '_')}.zip",
            mime="application/zip"
        )
