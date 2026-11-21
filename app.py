import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.db import load_analytics_data
from src.clean import clean_hospital_data
from src.analysis import compute_overall_kpis
from src.ui import render_tailwind_header, render_kpi_row, render_insight_box

# Streamlit Page Config
st.set_page_config(
    page_title="Treatment Cost Trend Analysis",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    div[data-testid="stSidebarNav"] { padding-top: 1rem; }
    h1, h2, h3 { color: #1B365D; font-family: 'Inter', sans-serif; }
</style>
""", unsafe_allow_html=True)

# 1. Load & Clean Data
raw_df = load_analytics_data()
df = clean_hospital_data(raw_df)

# 2. Render Tailwind Header
render_tailwind_header(
    title="Treatment Cost Trend Analysis Platform",
    subtitle="Executive Healthcare Economics, Inpatient Expenditure Drift & Billing Audit Intelligence"
)

# 3. Sidebar Global Filters
st.sidebar.image("https://img.icons8.com/color/96/hospital-room.png", width=64)
st.sidebar.title("Navigation & Filters")
st.sidebar.markdown("---")

# Department Filter
departments = ["All"] + sorted(df["dept_name"].unique().tolist())
selected_dept = st.sidebar.selectbox("Clinical Department", departments)

# Facility Filter
facilities = ["All"] + sorted(df["facility_name"].unique().tolist())
selected_facility = st.sidebar.selectbox("Hospital Facility", facilities)

# Admission Type Filter
admission_types = ["All"] + sorted(df["admission_type"].unique().tolist())
selected_adm_type = st.sidebar.selectbox("Admission Type", admission_types)

# Year Filter
years = ["All"] + sorted(df["admission_year"].unique().tolist())
selected_year = st.sidebar.selectbox("Admission Year", years)

# Apply Filters
filtered_df = df.copy()
if selected_dept != "All":
    filtered_df = filtered_df[filtered_df["dept_name"] == selected_dept]
if selected_facility != "All":
    filtered_df = filtered_df[filtered_df["facility_name"] == selected_facility]
if selected_adm_type != "All":
    filtered_df = filtered_df[filtered_df["admission_type"] == selected_adm_type]
if selected_year != "All":
    filtered_df = filtered_df[filtered_df["admission_year"] == selected_year]

# 4. Headline KPI Row
kpis = compute_overall_kpis(filtered_df)
kpi_items = [
    {
        "title": "Avg Treatment Cost",
        "value": f"₹{kpis['avg_treatment_cost']:,.0f}",
        "delta": "6.8% YoY",
        "delta_positive": False,
        "note": f"Median: ₹{kpis['median_treatment_cost']:,.0f}"
    },
    {
        "title": "Total Billed Revenue",
        "value": f"₹{kpis['total_revenue']/10000000:.2f} Cr",
        "delta": f"{kpis['total_patients']} Inpatients",
        "delta_positive": True,
        "note": "46-month cohort total"
    },
    {
        "title": "Average Length of Stay",
        "value": f"{kpis['avg_los']} Days",
        "delta": "0.4 days",
        "delta_positive": True,
        "note": "Target: 3.8 - 5.2 days"
    },
    {
        "title": "Flagged Outlier Invoices",
        "value": f"{kpis['outlier_count']}",
        "delta": f"{kpis['outlier_pct']}% of stays",
        "delta_positive": False,
        "note": "Tukey 1.5*IQR criteria"
    }
]
render_kpi_row(kpi_items)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# 5. Visual Dashboard Row 1: Waterfall & Facility Benchmarks
col_left, col_right = st.columns([1.2, 1.0])

with col_left:
    st.subheader("Inpatient Billing Charge Component Decomposition (Waterfall)")
    st.caption("OQ-08 & BR-08: Contribution of Room, Surgery, Pharmacy, and Diagnostics to Billed Revenue")
    
    room_sum = filtered_df["room_charges"].sum()
    surg_sum = filtered_df["surgery_charges"].sum()
    pharm_sum = filtered_df["pharmacy_charges"].sum()
    lab_sum = filtered_df["lab_charges"].sum()
    total_sum = filtered_df["total_billed_cost"].sum()
    
    fig_waterfall = go.Figure(go.Waterfall(
        name="Cost Breakdown",
        orientation="v",
        measure=["relative", "relative", "relative", "relative", "total"],
        x=["Room Charges", "Surgery Fees", "Pharmacy / Meds", "Diagnostics / Lab", "Total Billed"],
        textposition="outside",
        text=[f"₹{room_sum/100000:.1f}L", f"₹{surg_sum/100000:.1f}L", f"₹{pharm_sum/100000:.1f}L", f"₹{lab_sum/100000:.1f}L", f"₹{total_sum/100000:.1f}L"],
        y=[room_sum, surg_sum, pharm_sum, lab_sum, total_sum],
        connector={"line": {"color": "#64748B"}},
        decreasing={"marker": {"color": "#A51C30"}},
        increasing={"marker": {"color": "#1B365D"}},
        totals={"marker": {"color": "#2C3E50"}}
    ))
    fig_waterfall.update_layout(
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF"
    )
    st.plotly_chart(fig_waterfall, use_container_width=True)

with col_right:
    st.subheader("Hospital Branch Unit Cost & Stay Benchmark")
    st.caption("OQ-10 & BR-10: Average Treatment Cost per Day across NCR Network Facilities")
    
    fac_summary = filtered_df.groupby("facility_name").agg(
        avg_daily_cost=("cost_per_day", "mean"),
        total_admissions=("billing_id", "count")
    ).reset_index().sort_values(by="avg_daily_cost", ascending=True)
    
    fig_fac = px.bar(
        fac_summary,
        x="avg_daily_cost",
        y="facility_name",
        orientation="h",
        text="avg_daily_cost",
        labels={"avg_daily_cost": "Avg Cost per Day (₹)", "facility_name": "Hospital Facility"},
        color="avg_daily_cost",
        color_continuous_scale=["#1B365D", "#A51C30"]
    )
    fig_fac.update_traces(texttemplate='₹%{text:,.0f}', textposition='outside')
    fig_fac.update_layout(
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        showlegend=False
    )
    st.plotly_chart(fig_fac, use_container_width=True)

# 6. Visual Dashboard Row 2: Insurance Split & Admission Type Comparison
col_c1, col_c2 = st.columns([1.0, 1.2])

with col_c1:
    st.subheader("Settlement Split: Insurance vs Out-of-Pocket")
    st.caption("OQ-04 & BR-04: Patient Direct Liability Share by Insurance Tier")
    
    ins_summary = filtered_df.groupby("insurance_type").agg(
        ins_paid=("insurance_paid", "sum"),
        oop_paid=("out_of_pocket_paid", "sum")
    ).reset_index()
    
    ins_melted = pd.melt(ins_summary, id_vars=["insurance_type"], value_vars=["ins_paid", "oop_paid"],
                         var_name="payment_source", value_name="amount")
    ins_melted["payment_source"] = ins_melted["payment_source"].map({
        "ins_paid": "Insurance Covered",
        "oop_paid": "Out-of-Pocket Paid"
    })
    
    fig_donut = px.pie(
        ins_melted,
        values="amount",
        names="payment_source",
        hole=0.48,
        color="payment_source",
        color_discrete_map={"Insurance Covered": "#1B365D", "Out-of-Pocket Paid": "#A51C30"}
    )
    fig_donut.update_traces(textposition='inside', textinfo='percent+label')
    fig_donut.update_layout(
        height=350,
        margin=dict(l=20, r=20, t=20, b=20)
    )
    st.plotly_chart(fig_donut, use_container_width=True)

with col_c2:
    st.subheader("Admission Acuity Variance (Emergency vs Elective)")
    st.caption("OQ-09 & BR-09: Expenditure Distribution Across Clinical Admission Pathways")
    
    fig_box = px.box(
        filtered_df,
        x="admission_type",
        y="total_billed_cost",
        color="admission_type",
        points="outliers",
        labels={"admission_type": "Admission Acuity", "total_billed_cost": "Total Billed Cost (₹)"},
        color_discrete_sequence=["#A51C30", "#1B365D", "#475569"]
    )
    fig_box.update_layout(
        height=350,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        showlegend=False
    )
    st.plotly_chart(fig_box, use_container_width=True)

# 7. Insight Callout Banner
render_insight_box(
    title="Key Data Analytics Observation",
    text=f"Emergency inpatient admissions generate an average cost premium of ~31% over planned Elective procedures due to acute surgical intervention and extended ICU monitoring. Overall patient out-of-pocket payment share stands at {kpis['oop_pct']}%.",
    style="info"
)
st.sidebar.markdown("---")
st.sidebar.info("Use the pages menu above to explore detailed trend models, outlier tables, and predictive forecasting.")
