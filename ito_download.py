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

    # Definieren der ITO-Dateien nach Vorlagen
    ito_files = {
        "Mehrschichtig": [
            "ito_templates/MehrschichtigInklEinschichtig.ito"
        ],
        "SIA 416": [
            "ito_templates/SIA416.ito"
        ],
        "Bauteilkategorien (Elementtypen)": [
            "ito_templates/ARC Covering.ito",
            "ito_templates/ARC Geländer.ito",
            "ito_templates/ARC Fenster.ito",
            "ito_templates/ARC Stützen.ito",
            "ito_templates/ARC Treppen.ito",
            "ito_templates/ARC Türen.ito",
            "ito_templates/ARC Wände.ito",
            "ito_templates/ARC Curtain Wall.ito",
            "ito_templates/ARC Decken.ito"
        ],
        "Master Auswertung": [
            "ito_template/ARC Master.ito"
        ]
    }

    # Vorlage auswählen
    selected = st.selectbox("Vorlage auswählen", list(ito_files.keys()))
    file_paths = ito_files[selected]

    # Bestimmen des Tools basierend auf der Auswahl
    if selected == "Mehrschichtig":
        tool = "Mehrschichtig Bereinigen"
    elif selected == "Bauteilkategorien (Elementtypen)":
        tool = "Master Table"
    elif selected == "Master Auswertung":
        tool = "Spalten Mengen Merger"
    elif selected == "SIA 416":
        tool = None
        st.warning("""
            Hinweis: Für die SIA 416-Vorlage müssen Sie kein Tool verwenden, ausser Sie fügen gleiche Mengentypen später ein. 
            Dann **Spalten Mengen Merger** verwenden.
        """)
    else:
        tool = None
        st.info("Kein Tool empfohlen, da die ITO gleich so ausgegeben werden kann.")

    # Zeigen des empfohlenen Tools, wenn verfügbar
    if tool:
        st.info(f"Empfohlenes Tool für {selected}: **{tool}**")

    # Wenn mehrere ITOs vorhanden sind, als ZIP herunterladen
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
