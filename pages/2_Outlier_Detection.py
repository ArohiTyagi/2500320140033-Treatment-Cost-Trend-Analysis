import streamlit as st
import pandas as pd
import plotly.express as px
from src.db import load_analytics_data
from src.clean import clean_hospital_data
from src.analysis import detect_outliers_iqr
from src.ui import render_tailwind_header, render_insight_box

st.set_page_config(page_title="Billing Outlier Detection", page_icon="🔍", layout="wide")

df = clean_hospital_data(load_analytics_data())

render_tailwind_header(
    title="Statistical Billing Outlier & Anomaly Detection",
    subtitle="Automated Inpatient Invoice Auditing via Tukey's IQR Fences and Standardized Z-Scores"
)

# Sidebar Controls
st.sidebar.header("Audit Sensitivity Controls")
iqr_factor = st.sidebar.slider("Tukey IQR Multiplier", min_value=1.0, max_value=3.0, value=1.5, step=0.1)
z_thresh = st.sidebar.slider("Minimum Z-Score Filter", min_value=1.5, max_value=4.0, value=2.5, step=0.1)

# 1. Outlier Detection Engine
outliers_df = detect_outliers_iqr(df, factor=iqr_factor)
high_z_outliers = outliers_df[outliers_df["z_score"] >= z_thresh] if not outliers_df.empty else pd.DataFrame()

# 2. Scatter Plot Visualizing Outliers
st.subheader("1. Inpatient Cost vs Length of Stay (Highlighting Statistical Anomalies)")
st.caption("OQ-06 & BR-05: Invoices exceeding statistical variance boundaries flagged in Crimson")

fig_scatter = px.scatter(
    df,
    x="los_days",
    y="total_billed_cost",
    color="is_outlier",
    color_discrete_map={0: "#1B365D", 1: "#A51C30"},
    hover_data=["billing_id", "dept_name", "procedure_name", "patient_name"],
    labels={"los_days": "Length of Stay (Days)", "total_billed_cost": "Total Billed Cost (₹)", "is_outlier": "Outlier Flag"}
)
fig_scatter.update_traces(marker=dict(size=9, opacity=0.8))
fig_scatter.update_layout(height=420, margin=dict(l=20, r=20, t=20, b=20), plot_bgcolor="#FFFFFF")
st.plotly_chart(fig_scatter, use_container_width=True)

# 3. Flagged Invoice Audit Table
st.subheader("2. Flagged High-Cost Claim Audit Ledger")
st.caption("Detailed breakdown of flagged hospitalizations for clinical auditor review")

if not high_z_outliers.empty:
    display_cols = [
        "billing_id", "patient_name", "patient_age", "dept_name", "procedure_name",
        "los_days", "admission_type", "total_billed_cost", "upper_fence", "z_score", "insurance_type"
    ]
    st.dataframe(
        high_z_outliers[display_cols].rename(columns={
            "billing_id": "Invoice ID",
            "patient_name": "Patient",
            "patient_age": "Age",
            "dept_name": "Specialty",
            "procedure_name": "Procedure",
            "los_days": "LOS",
            "admission_type": "Acuity",
            "total_billed_cost": "Billed (₹)",
            "upper_fence": "Tukey Fence (₹)",
            "z_score": "Z-Score",
            "insurance_type": "Insurer"
        }),
        height=320,
        use_container_width=True
    )
    
    # Export CSV
    csv_bytes = high_z_outliers.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Flagged Outliers Audit Ledger (CSV)",
        data=csv_bytes,
        file_name="flagged_hospital_billing_outliers.csv",
        mime="text/csv"
    )
else:
    st.success("No invoices exceed the configured Z-score and IQR sensitivity threshold.")

render_insight_box(
    title="Audit Recommendation",
    text=f"Detected {len(high_z_outliers)} hospitalization invoices with extreme cost variance (Z > {z_thresh}). Recommended Action: Audit pharmacy itemization and surgical consumable unbundling before submitting claims to Third-Party Administrators (TPAs).",
    style="danger"
)
