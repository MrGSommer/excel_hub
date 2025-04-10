import streamlit as st

def app():
    st.header("Solibri ITO-Vorlagen herunterladen")
    st.markdown("""
    Hier finden Sie vordefinierte **Quantity Take-Offs (ITO-Dateien)** für den Einsatz in **Solibri**.
    Diese Vorlagen können direkt importiert und bei Bedarf mit den Tools dieser Plattform kombiniert werden.
    """)

    ito_files = {
        "eBKP-H Mengenauswertung": "ito_templates/ebkph_takeoff.ito",
        "Flächen- und Volumen-ITO": "ito_templates/flaeche_volumen.ito",
        "Bauteilkategorien (Elementtypen)": "ito_templates/elementtypen.ito"
    }

    selected = st.selectbox("Vorlage auswählen", list(ito_files.keys()))
    path = ito_files[selected]

    try:
        with open(path, "rb") as f:
            st.download_button(
                label="📥 ITO-Datei herunterladen",
                data=f,
                file_name=path.split("/")[-1],
                mime="application/octet-stream"
            )
    except FileNotFoundError:
        st.error("Datei nicht gefunden. Bitte sicherstellen, dass sie im Ordner `ito_templates/` abgelegt ist.")
