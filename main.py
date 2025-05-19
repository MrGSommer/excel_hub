import streamlit as st
from excel_requirements import app as excel_requirements
from spalten_values_merger import app as values_merger
from mehrschichtig_bereinigen import app as mehrsch_bereinigen
from advanced_excel_merge_master import app as advanced_merge_master
from advanced_excel_merge_table import app as advanced_merge_table
from advanced_excel_merge_sheets import app as advanced_merge_sheets
from merging_flow import app as merging_flow_columns_to_table
from app_advisor import app_advisor
from ito_download import app as download_templates

st.set_page_config(page_title="Excel Operation Tools", layout="wide")
st.title("Excel Operation Tools 🚀")
st.markdown("Willkommen! Wählen Sie einen Tab für verschiedene Excel-Operationen oder den Download von Solibri ITOs.")

# Globale Einstellungen für alle Tools
st.sidebar.header("Globale Einstellungen für alle Tools")
with st.sidebar.expander("Verarbeitungseinstellungen"):
    st.markdown(
        """
        **File Supplement Name:**  
        Leer lassen ⇒ jede App setzt Default (Dateiname oder Sheet)

        **Globale Bereinigung:**  
        Entfernt standardmässig Mengeneinheiten (z. B. m2, m3, kg).

        **Zusätzliche Zeichen (kommagetrennt):**  
        z. B. cm, CHF  
        ⇒ Nur löschen, wenn Häkchen aktiviert
        """
    )

# Supplement-Name (leer ⇒ Sub-Apps setzen selbst Default)
supplement_name = st.sidebar.text_input(
    "File Supplement Name",
    value="",
    key="global_supplement"
)

# Zusätzliche Zeichen
custom_chars = st.sidebar.text_input(
    "Zusätzliche Zeichen (kommagetrennt)",
    value="",
    key="global_custom"
)
delete_custom = st.sidebar.checkbox(
    "Zusätzliche Zeichen löschen",
    value=False,
    key="global_custom_delete"
)

# Flags für Sub-Apps
delete_enabled = True
if not delete_custom:
    custom_chars = ""




tabs = st.tabs([
    "Tool-Beratung",
    "Excel-Anforderungen",
    "Spalten Mengen Merger", 
    "Mehrschichtig Bereinigen", 
    "Master Table", 
    "Merge to Table", 
    "Merge to Sheets",
    "Flow Merging",
    "Download Templates"
])

with tabs[0]:
    app_advisor()

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

with tabs[7]:
    merging_flow_columns_to_table(supplement_name, delete_enabled, custom_chars)

with tabs[8]:
    download_templates()
