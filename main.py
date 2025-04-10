import streamlit as st
from excel_requirements import app as excel_requirements
from spalten_values_merger import app as values_merger
from mehrschichtig_bereinigen import app as mehrsch_bereinigen
from advanced_excel_merge_master import app as advanced_merge_master
from advanced_excel_merge_table import app as advanced_merge_table
from advanced_excel_merge_sheets import app as advanced_merge_sheets
from tool_advisor import app as tool_advisor

st.set_page_config(page_title="Excel Operation Tools", layout="wide")
st.title("Excel Operation Tools 🚀")
st.markdown("Willkommen! Wählen Sie einen Tab für verschiedene Excel-Operationen.")

# Globale Einstellungen in der Sidebar mit einem expander für optionale Infos
st.sidebar.header("Globale Einstellungen für Advanced Excel Merger")
with st.sidebar.expander("Globale Verarbeitungseinstellungen"):
    st.markdown(
        """
        **Diese Einstellungen gelten für alle Tools:**

        - 🔤 **File Supplement Name:**  
          Namenszusatz für die generierten Dateien (z. B. `default_merged_output.xlsx`).
        
        - ✂️ **Zeichen in Zellen entfernen:**  
          Aktiviert das Entfernen von Einheiten wie `" m2"`, `" m3"`, `" kg"` etc.
        
        - ✏️ **Zusätzliche Zeichen (kommagetrennt):**  
          Optional weitere Zeichen, die entfernt werden sollen (z. B. `"cm", "CHF"`).

        Diese Einstellungen sorgen für eine konsistente Datenbereinigung über alle Module hinweg.
        """
    )

supplement_name = st.sidebar.text_input("File Supplement Name", value="default", key="global_supplement")
delete_enabled = st.sidebar.checkbox("Zeichen in Zellen entfernen", key="global_delete")
custom_chars = st.sidebar.text_input("Zusätzliche zu löschende Zeichen (kommagetrennt)", value="", key="global_custom")

tabs = st.tabs([
    "Tool-Beratung",
    "Excel-Anforderungen",
    "Spalten Mengen Merger", 
    "Mehrschichtig Bereinigen", 
    "Master Table", 
    "Merge to Table", 
    "Merge to Sheets"
])

with tabs[0]:
    tool_advisor()
    
with tabs[1]:
    excel_requirements()

with tabs[2]:
    values_merger(supplement_name, delete_enabled, custom_chars)

with tabs[3]:
    mehrsch_bereinigen(supplement_name, delete_enabled, custom_chars)

with tabs[4]:
    advanced_merge_master(supplement_name, delete_enabled, custom_chars)

with tabs[5]:
    advanced_merge_table(supplement_name, delete_enabled, custom_chars)

with tabs[6]:
    advanced_merge_sheets(supplement_name, delete_enabled, custom_chars)


