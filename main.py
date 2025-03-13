import streamlit as st
from spalten_values_merger import app as values_merger
from mehrschichtig_bereinigen import app as mehrsch_bereinigen
from advanced_excel_merge_master import app as advanced_merge_master
from advanced_excel_merge_table import app as advanced_merge_table
from advanced_excel_merge_sheets import app as advanced_merge_sheets

st.set_page_config(page_title="Excel Operation Tools", layout="wide")
st.title("Excel Operation Tools 🚀")
st.markdown("Willkommen! Wählen Sie einen Tab für verschiedene Excel-Operationen.")

tabs = st.tabs([
    "Spalten Mengen Merger", 
    "Mehrschichtig Bereinigen", 
    "Master Table", 
    "Merge to Table", 
    "Merge to Sheets"
])

with tabs[0]:
    values_merger()

with tabs[1]:
    mehrsch_bereinigen()

with tabs[2]:
    advanced_merge_master()

with tabs[3]:
    advanced_merge_table()

with tabs[4]:
    advanced_merge_sheets()
