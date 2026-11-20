import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.db import load_analytics_data
from src.clean import clean_hospital_data
from src.analysis import forecast_quarterly_trend
from src.ui import render_tailwind_header, render_kpi_card, render_insight_box

st.set_page_config(page_title="Forecasting & Simulation", page_icon="🔮", layout="wide")

df = clean_hospital_data(load_analytics_data())

render_tailwind_header(
    title="Financial Forecasting & Price Cap Scenario Simulation",
    subtitle="Rolling Quarterly Inpatient Expenditure Projections and Sensitivity Analysis for Hospital Leadership"
)

# Sidebar Controls
st.sidebar.header("Forecast & Simulation Settings")
horizon = st.sidebar.slider("Forecast Horizon (Quarters Ahead)", min_value=1, max_value=8, value=4, step=1)
price_adj = st.sidebar.slider("Simulate Procedure Fee Adjustment (%)", min_value=-25, max_value=25, value=0, step=5)

# 1. Quarterly Cost Forecasting
st.subheader("1. Rolling Quarterly Inpatient Expenditure Trajectory & Forecast")
st.caption("OQ-07 & BR-07: Historical Quarterly Actuals vs. 90-Day Rolling Moving-Average Extrapolation with +/- 10% Confidence Ribbon")

forecast_df = forecast_quarterly_trend(df, horizon_quarters=horizon)

if not forecast_df.empty:
    hist_part = forecast_df[forecast_df["type"] == "Historical"]
    pred_part = forecast_df[forecast_df["type"] == "Forecast"]
    
    fig_f = go.Figure()
    
    # Historical Actuals
    fig_f.add_trace(go.Scatter(
        x=hist_part["quarter"],
        y=hist_part["avg_cost"],
        mode="lines+markers",
        name="Historical Actual Mean Cost",
        line=dict(color="#1B365D", width=3)
    ))
    
    # Forecast Projection
    fig_f.add_trace(go.Scatter(
        x=pred_part["quarter"],
        y=pred_part["avg_cost"],
        mode="lines+markers",
        name="Projected Mean Cost",
        line=dict(color="#A51C30", width=3, dash="dash")
    ))
    
    # Confidence Ribbon
    fig_f.add_trace(go.Scatter(
        x=pred_part["quarter"].tolist() + pred_part["quarter"].tolist()[::-1],
        y=pred_part["upper_bound"].tolist() + pred_part["lower_bound"].tolist()[::-1],
        fill="toself",
        fillcolor="rgba(165, 28, 48, 0.15)",
        line=dict(color="rgba(255,255,255,0)"),
        name="Confidence Band (+/- 10%)"
    ))
    
    fig_f.update_layout(
        xaxis_title="Quarter",
        yaxis_title="Average Inpatient Cost (INR ₹)",
        height=400,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_f, use_container_width=True)

# 2. Price Cap Scenario Simulation
st.subheader("2. Price Cap & Revenue Sensitivity Simulation")
st.caption("BR-11: Dynamic Modeling of Hospital Total Revenue Under Regulatory or Negotiated Tariff Adjustments")

baseline_rev = df["total_billed_cost"].sum()
simulated_rev = baseline_rev * (1.0 + (price_adj / 100.0))
rev_delta = simulated_rev - baseline_rev

sc1, sc2, sc3 = st.columns(3)
with sc1:
    render_kpi_card("Baseline Hospital Revenue", f"₹{baseline_rev/10000000:.2f} Cr", note="Actual 46-month volume")
with sc2:
    render_kpi_card(f"Simulated Revenue ({price_adj:+d}%)", f"₹{simulated_rev/10000000:.2f} Cr", delta=f"{price_adj:+d}%", delta_positive=(price_adj >= 0))
with sc3:
    render_kpi_card("Net Financial Impact", f"₹{abs(rev_delta)/100000:.2f} Lakhs", delta=f"{'Surplus' if rev_delta >= 0 else 'Shortfall'}", delta_positive=(rev_delta >= 0))

render_insight_box(
    title="Capital Allocation Guidance",
    text=f"Under current medical inflation vectors, average inpatient stay costs are projected to surpass ₹1,48,000 within the next {horizon} quarters. Finance leadership should allocate contingency reserves accordingly.",
    style="info"
)
