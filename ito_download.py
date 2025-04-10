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
            "ito_templates/ARC Fenster.ito",
            "ito_templates/ARC Stützen.ito"
        ],
        "SIA 416": [
            "ito_templates/ARC Treppen.ito"
        ],
        "Bauteilkategorien (Elementtypen)": [
            "ito_templates/ARC Covering.ito",
            "ito_templates/ARC Decken.ito"
        ]
    }

    selected = st.selectbox("Vorlage auswählen", list(ito_files.keys()))
    file_paths = ito_files[selected]

    for path in file_paths:
        st.write(f"Versuche, die Datei zu öffnen: {path}")
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
        except Exception as e:
            st.error(f"Fehler beim Öffnen der Datei {path}: {e}")

