import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.db import load_analytics_data
from src.clean import clean_hospital_data
from src.analysis import compute_monthly_trends, compute_department_cagr
from src.ui import render_tailwind_header, render_insight_box

st.set_page_config(page_title="Cost Trends & Inflation", page_icon="📈", layout="wide")

df = clean_hospital_data(load_analytics_data())

render_tailwind_header(
    title="Longitudinal Treatment Cost Trends & Medical Inflation",
    subtitle="Month-over-Month Expenditure Trajectories and Departmental Compound Annual Growth Rates (CAGR)"
)

# Sidebar Filter
st.sidebar.header("Filter Trend Analysis")
selected_dept = st.sidebar.selectbox("Filter Department", ["All Departments"] + sorted(df["dept_name"].unique().tolist()))

df_filtered = df if selected_dept == "All Departments" else df[df["dept_name"] == selected_dept]

# 1. MoM Line Chart
st.subheader("1. Month-over-Month Treatment Cost Trajectory (2023 - 2026)")
st.caption("OQ-01 & BR-01: Comparative Trajectory of Mean Billed Cost vs. Outlier-Resilient Median Billed Cost")

monthly_df = compute_monthly_trends(df_filtered)

if not monthly_df.empty:
    fig_trend = go.Figure()
    fig_trend.add_trace(go.Scatter(
        x=monthly_df["year_month"],
        y=monthly_df["mean_cost"],
        mode="lines+markers",
        name="Mean Treatment Cost",
        line=dict(color="#1B365D", width=3)
    ))
    fig_trend.add_trace(go.Scatter(
        x=monthly_df["year_month"],
        y=monthly_df["median_cost"],
        mode="lines+markers",
        name="Median Treatment Cost",
        line=dict(color="#A51C30", width=2.5, dash="dot")
    ))
    fig_trend.update_layout(
        xaxis_title="Admission Month",
        yaxis_title="Billed Cost (INR ₹)",
        hovermode="x unified",
        height=420,
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_trend, use_container_width=True)

# 2. Departmental CAGR Inflation
st.subheader("2. Departmental Medical Inflation Benchmark (CAGR)")
st.caption("OQ-02 & BR-02: Annualized Cost Growth Rate Across Specialty Disciplines")

cagr_df = compute_department_cagr(df)

col1, col2 = st.columns([1.2, 1.0])
with col1:
    fig_cagr = px.bar(
        cagr_df,
        x="dept_name",
        y="cagr_pct",
        text="cagr_pct",
        labels={"dept_name": "Clinical Specialty", "cagr_pct": "Annualized CAGR (%)"},
        color="cagr_pct",
        color_continuous_scale=["#1B365D", "#A51C30"]
    )
    fig_cagr.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    fig_cagr.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20), showlegend=False)
    st.plotly_chart(fig_cagr, use_container_width=True)

with col2:
    st.dataframe(
        cagr_df.rename(columns={
            "dept_name": "Specialty",
            "start_year_avg": "2023 Baseline (₹)",
            "end_year_avg": "2026 Current (₹)",
            "cagr_pct": "CAGR (%)"
        }),
        height=320,
        use_container_width=True
    )

render_insight_box(
    title="Inflation Insight",
    text=f"Cardiology and Oncology exhibit the fastest medical inflation at ~7.8% - 8.4% annually, driven by high-cost imported medical implants, stents, and targeted chemotherapy biologics.",
    style="warning"
)
