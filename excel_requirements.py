import streamlit as st

def app():
    st.header("Excel-Anforderungen für eine erfolgreiche Verarbeitung")
    st.markdown("Bevor Sie eines der Tools verwenden, stellen Sie bitte sicher, dass Ihre Excel-Dateien folgende Strukturvorgaben erfüllen. Nur so kann eine korrekte und automatisierte Verarbeitung garantiert werden.")

    with st.expander("1. Spalten Mengen Merger"):
        st.subheader("Ziel")
        st.markdown("Zusammenführen mehrerer Mengenspalten (z. B. Dicke, Flaeche, Volumen, Laenge, Hoehe) in eine standardisierte Spalte mit definiertem Namen.")

        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Header-Erkennung:**  
              Der Header muss eine Zelle mit dem Inhalt **Teilprojekt** enthalten.  
              Der Header wird automatisch erkannt (erste Zeile mit dem Begriff "Teilprojekt").

            - **Mengenspalten (automatische Erkennung & Umbenennung):**  
              Folgende Spalten werden automatisch erkannt und in einheitliche Namen umbenannt:
              - **Fläche (m2):** `Fläche`, `Flaeche`, `Fläche BQ`, `Fläche Total`, `Fläche Solibri`
              - **Volumen (m3):** `Volumen`, `Volumen BQ`, `Volumen Total`, `Volumen Solibri`
              - **Länge (m):** `Länge`, `Laenge`, `Länge BQ`, `Länge Solibri`
              - **Dicke (m):** `Dicke`, `Stärke`, `Dicke BQ`, `Dicke Solibri`
              - **Höhe (m):** `Höhe`, `Hoehe`, `Höhe BQ`, `Höhe Solibri`

            - **Automatisierte Verarbeitung:**  
              - Alle genannten Spalten werden in der definierten Priorität zu einer Zielsummenspalte gemerged.
              - Nicht benötigte Spalten werden entfernt.
              - Werte werden als Dezimalzahl (float) interpretiert.
              - Einheiten wie `" m2"`, `" m"`, `" m3"` etc. werden automatisch entfernt.
              - Weitere Zeichen können optional über die Sidebar-Einstellungen entfernt werden.

            - **Beispiel:**

              | Teilprojekt | Fläche | Fläche BQ | Volumen | Laenge | Anderes |
              |-------------|--------|------------|---------|--------|---------|
              | TP1         | 12.0   |            | 3.0     | 2.5    | Info    |
            """
        )

    with st.expander("2. Mehrschichtig Bereinigen"):
        st.subheader("Ziel")
        st.markdown("Automatische Bereinigung von mehrschichtigen Excel-Daten mit logischer Trennung von Mutter- und Subzeilen.")

        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - **Erforderliche Spalten:**  
              - `Teilprojekt`  
              - `Geschoss`  
              - `Unter Terrain`  
              Diese werden als Masterspalten verwendet, um die Mutterzeilen zu erkennen.

            - **Zusätzliche Klassifizierung:**  
              - `eBKP-H`: Klassifizierung der Mutterzeile  
              - `eBKP-H Sub`: Klassifizierung der Subzeilen

            - **Funktion:**  
              - Mutterzeilen ohne gültige Klassifizierung (z. B. `"Nicht klassifiziert"`) werden entfernt, wenn sie keine zugehörige gültige Subzeile haben.
              - Subzeilen mit gültigem `eBKP-H Sub` werden extrahiert und in neue Einträge mit übertragenen Masterdaten überführt.
              - Einheitliche Mengenspalten werden auch hier bereinigt, umbenannt und numerisch interpretiert.

            - **Beispiel:**

              | Teilprojekt | Geschoss | Unter Terrain | eBKP-H         | eBKP-H Sub       | GUID     |
              |-------------|----------|----------------|----------------|------------------|----------|
              | TP1         | EG       | nein           | Nicht klassifiziert | Wärmedämmung  | A1B2C3   |
              |             |          |                |                | Wärmeschutzplatte | A1B2C3   |
            """
        )

    with st.expander("3. Master Table"):
        st.subheader("Ziel")
        st.markdown("Zusammenführen mehrerer Arbeitsblätter einer Excel-Datei zu einer einheitlichen Tabelle.")

        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - Jedes Arbeitsblatt muss eine erkennbare Headerzeile enthalten.  
              Die erste Zeile mit den meisten nicht-leeren Zellen wird als Header interpretiert.

            - Alle Blätter werden automatisch standardisiert:  
              - Spaltennamen werden vereinheitlicht.  
              - Einheiten in Zellen werden entfernt.  
              - Inhalte werden gereinigt und numerisch konvertiert.

            - Eine zusätzliche Spalte **SheetName** gibt den Ursprung des Datensatzes an.

            - **Beispiel:**

              | SheetName | Dicke | Fläche (m2) | GUID     |
              |-----------|--------|--------------|----------|
              | Blatt1    | 0.3    | 12.5         | A1234    |
            """
        )

    with st.expander("4. Merge to Table"):
        st.subheader("Ziel")
        st.markdown("Zusammenführen mehrerer Excel-Dateien zu einer einheitlichen Tabelle. Die Spalten werden nach Häufigkeit sortiert.")

        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - Jede Datei muss:
              - Eine Headerzeile besitzen (erste Zeile).  
              - Darunter mindestens eine Datenzeile enthalten.

            - Alle Inhalte werden automatisch bereinigt:  
              - Standardisierte Spaltennamen  
              - Entfernte Einheiten  
              - Float-Konvertierung  
              - Optional: weitere Zeichenbereinigung

            - **Beispiel:**

              | Volumen | Dicke | Fläche | Zusatzinfo |
              |---------|--------|--------|-------------|
              | 4.3     | 0.25   | 22.4   | Bemerkung   |
            """
        )

    with st.expander("5. Merge to Sheets"):
        st.subheader("Ziel")
        st.markdown("Mehrere Excel-Dateien werden in ein neues Dokument übernommen – je Datei ein neues Arbeitsblatt.")

        st.subheader("Erforderliche Struktur")
        st.markdown(
            """
            - Jedes File muss mindestens ein Blatt mit Daten beinhalten.
            - Der erste Tab jeder Datei wird automatisch übernommen.
            - Blattnamen werden aus dem Dateinamen generiert (gekürzt auf 30 Zeichen).

            - Inhalte werden automatisiert bereinigt:
              - Zeichen wie `" m2"`, `" m3"` etc. entfernt
              - Float-Parsing aktiv
              - Weitere optional definierbare Zeichenbereinigung

            - **Beispiel:**  
              Datei "Projekt_A.xlsx" → Tab "Projekt_A"
            """
        )

    st.markdown("---")
    st.info("Eine saubere und konsistente Excel-Struktur ist essenziell für die erfolgreiche Ausführung aller Tools. Prüfen Sie Ihre Dateien sorgfältig vor dem Upload.")
