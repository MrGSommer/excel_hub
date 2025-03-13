import streamlit as st

def app():
    st.header("Excel-Anforderungen für eine erfolgreiche Verarbeitung")
    st.markdown("Um sicherzustellen, dass Ihre Excel-Dateien korrekt bearbeitet werden, beachten Sie bitte die folgenden Strukturvorgaben für die einzelnen Tools:")

    with st.expander("1. Spalten Mengen Merger"):
        st.subheader("Ziel")
        st.markdown("Zusammenführen ausgewählter Mengenspalten (z.B. Dicke, Flaeche, Volumen, Laenge, Hoehe) zu einem kombinierten Wert.")
        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Header-Erkennung:**  
              Der Header muss den Zellwert **Teilprojekt** enthalten.
            - **Mengenspalten:**  
              - **Dicke:** numerisch (Einheit: m)  
              - **Flaeche:** numerisch (Einheit: m²)  
              - **Volumen:** numerisch (Einheit: m³)  
              - **Laenge:** numerisch (Einheit: m)  
              - **Hoehe:** numerisch (Einheit: m)
            - **Weitere Spalten:**  
              Zusätzliche Daten können vorhanden sein, werden aber ggf. nach dem Merge entfernt.
            - **Beispiel:**
              
              | Teilprojekt | Dicke | Flaeche | Volumen | Laenge | Hoehe | Anderes |
              |-------------|-------|---------|---------|--------|-------|---------|
              | TP1         | 0.2   | 15.0    | 3.0     | 5.0    | 2.5   | ...     |
            """
        )

    with st.expander("2. Mehrschichtig Bereinigen"):
        st.subheader("Ziel")
        st.markdown("Bereinigung von mehrschichtigen Excel-Daten mit Hierarchie (Mutter- und Subzeilen).")
        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Masterspalten:**  
              *Teilprojekt*, *Geschoss*, *Unter Terrain*  
              - **Mutterzeile:** Enthält gültige Werte.  
              - **Subzeile:** Diese Spalten sind leer und werden per Pre-Mapping ergänzt.
            - **Wichtige Zusatzspalten:**  
              - **eBKP-H:** Zur Klassifizierung der Hauptelemente.  
              - **eBKP-H Sub:** Zur Identifikation von Subzeilen.
            - **Beispiel:**
              
              | Teilprojekt | Geschoss | Unter Terrain | eBKP-H         | eBKP-H Sub | GUID    |
              |-------------|----------|---------------|----------------|------------|---------|
              | TP1         | G1       | Wert1         | Klassifiziert  |            | 123-abc |
              |             |          |               |                | Unterwert1 | 123-abc |
            """
        )

    with st.expander("3. Master Table (Advanced Merger - Master Table)"):
        st.subheader("Ziel")
        st.markdown("Zusammenführen mehrerer Arbeitsblätter einer Excel-Datei zu einer einzigen Mastertabelle.")
        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Header-Erkennung:**  
              Jedes Arbeitsblatt muss einen Header haben (z.B. anhand der Zeile mit den meisten nicht-leeren Zellen).
            - **Datenfelder:**  
              Alle Spalten der ausgewählten Blätter werden übernommen.
            - **Zusätzliche Spalte:**  
              Eine Spalte **SheetName** wird hinzugefügt, um die Herkunft jeder Zeile anzugeben.
            - **Beispiel:**
              
              | SheetName | Spalte1 | Spalte2 | ... |
              |-----------|---------|---------|-----|
              | Blatt1    | Wert1   | Wert2   | ... |
              | Blatt2    | WertA   | WertB   | ... |
            """
        )

    with st.expander("4. Merge to Table (Advanced Merger - Merge to Table)"):
        st.subheader("Ziel")
        st.markdown("Mehrere Excel-Dateien werden zu einer einzigen Tabelle zusammengeführt. Die Spalten werden basierend auf ihrer Häufigkeit sortiert.")
        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Header:**  
              Jede Excel-Datei muss eine Header-Zeile besitzen (normalerweise die erste Zeile).
            - **Datenzeilen:**  
              Die Zeilen folgen direkt nach dem Header.
            - **Konsistenz:**  
              Die Spaltennamen sollten weitgehend übereinstimmen, um eine korrekte Zusammenführung zu ermöglichen.
            - **Beispiel:**
              
              | Spalte1 | Spalte2 | Spalte3 |
              |---------|---------|---------|
              | Wert1   | Wert2   | Wert3   |
            """
        )

    with st.expander("5. Merge to Sheets (Advanced Merger - Merge to Sheets)"):
        st.subheader("Ziel")
        st.markdown("Mehrere Excel-Dateien werden in eine neue Arbeitsmappe übernommen, wobei jede Datei als eigenes Blatt erscheint.")
        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Einzeldateien:**  
              Jede Datei muss mindestens ein Arbeitsblatt mit Daten enthalten.
            - **Struktur:**  
              Das Format und die Struktur der einzelnen Blätter bleiben unverändert.
            - **Blattname:**  
              Der Blattname wird aus dem Dateinamen abgeleitet (maximal 30 Zeichen).
            - **Beispiel:**
              
              Datei 1 wird zu Blatt "Datei1", Datei 2 zu Blatt "Datei2" etc.
            """
        )

    st.markdown("---")
    st.info("Überprüfen Sie Ihre Excel-Dateien anhand dieser Anforderungen, bevor Sie eines der Tools verwenden. Eine konsistente Datenstruktur ist entscheidend für einen erfolgreichen Merge und die anschließende Bearbeitung.")
