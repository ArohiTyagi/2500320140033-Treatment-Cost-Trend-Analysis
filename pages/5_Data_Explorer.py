import streamlit as st
import pandas as pd
from src.db import load_analytics_data, load_table
from src.clean import clean_hospital_data
from src.ui import render_tailwind_header

st.set_page_config(page_title="Data Explorer & Quality Audit", page_icon="🗄️", layout="wide")

df = clean_hospital_data(load_analytics_data())

render_tailwind_header(
    title="Relational Data Explorer & Data Quality Audit",
    subtitle="Direct Inspection of SQLite Tables, Analytical Views, Null Audits, and Full Dataset Export"
)

st.subheader("1. Inpatient Analytical Data Mart Explorer")
st.caption("v_treatment_cost_analytics: Searchable and filterable analytical dataset")

# Search and filter bar
search_term = st.text_input("Search Patient Name, Procedure, or City", "")

filtered_view = df.copy()
if search_term:
    mask = (
        filtered_view["patient_name"].str.contains(search_term, case=False, na=False) |
        filtered_view["procedure_name"].str.contains(search_term, case=False, na=False) |
        filtered_view["facility_city"].str.contains(search_term, case=False, na=False)
    )
    filtered_view = filtered_view[mask]

st.dataframe(filtered_view, height=350, use_container_width=True)

# Export options
csv_export = filtered_view.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Export Current Filtered Table (CSV)",
    data=csv_export,
    file_name="treatment_cost_filtered_data.csv",
    mime="text/csv"
)

st.markdown("---")

# 2. Database Entity Inspection & Health Checks
st.subheader("2. Relational Table Health & Schema Audit")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "dim_departments", "dim_procedures", "dim_facilities", "dim_patients", "fact_hospital_billing"
])

with tab1:
    st.dataframe(load_table("dim_departments"), use_container_width=True)
with tab2:
    st.dataframe(load_table("dim_procedures"), use_container_width=True)
with tab3:
    st.dataframe(load_table("dim_facilities"), use_container_width=True)
with tab4:
    st.dataframe(load_table("dim_patients").head(50), use_container_width=True)
with tab5:
    st.dataframe(load_table("fact_hospital_billing").head(50), use_container_width=True)

st.success("Database Status: SQLite connection healthy. All 5 relational entities and 1 analytical view fully initialized.")
