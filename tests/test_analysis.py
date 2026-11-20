import pytest
import pandas as pd
import numpy as np
from src.analysis import (
    compute_overall_kpis,
    compute_monthly_trends,
    compute_department_cagr,
    detect_outliers_iqr,
    compute_los_correlation,
    forecast_quarterly_trend
)
from src.clean import clean_hospital_data

@pytest.fixture
def sample_test_df():
    """Generates a small deterministic dataset for algorithm verification."""
    data = [
        {"billing_id": 1, "dept_name": "Cardiology", "admission_date": "2023-01-10", "los_days": 5, "room_charges": 25000, "surgery_charges": 100000, "pharmacy_charges": 20000, "lab_charges": 15000, "total_billed_cost": 160000, "insurance_paid": 120000, "out_of_pocket_paid": 40000, "is_outlier": 0, "insurance_type": "Corporate"},
        {"billing_id": 2, "dept_name": "Cardiology", "admission_date": "2024-01-15", "los_days": 6, "room_charges": 30000, "surgery_charges": 110000, "pharmacy_charges": 22000, "lab_charges": 16000, "total_billed_cost": 178000, "insurance_paid": 130000, "out_of_pocket_paid": 48000, "is_outlier": 0, "insurance_type": "Corporate"},
        {"billing_id": 3, "dept_name": "Orthopedics", "admission_date": "2023-02-12", "los_days": 4, "room_charges": 20000, "surgery_charges": 80000, "pharmacy_charges": 15000, "lab_charges": 10000, "total_billed_cost": 125000, "insurance_paid": 95000, "out_of_pocket_paid": 30000, "is_outlier": 0, "insurance_type": "Private"},
        {"billing_id": 4, "dept_name": "Orthopedics", "admission_date": "2024-02-18", "los_days": 5, "room_charges": 25000, "surgery_charges": 88000, "pharmacy_charges": 17000, "lab_charges": 12000, "total_billed_cost": 142000, "insurance_paid": 100000, "out_of_pocket_paid": 42000, "is_outlier": 0, "insurance_type": "Private"},
        {"billing_id": 5, "dept_name": "Cardiology", "admission_date": "2024-03-20", "los_days": 18, "room_charges": 120000, "surgery_charges": 350000, "pharmacy_charges": 80000, "lab_charges": 50000, "total_billed_cost": 600000, "insurance_paid": 450000, "out_of_pocket_paid": 150000, "is_outlier": 1, "insurance_type": "Government"},
    ]
    df = pd.DataFrame(data)
    df["admission_date"] = pd.to_datetime(df["admission_date"])
    return df

def test_kpi_computation(sample_test_df):
    kpis = compute_overall_kpis(sample_test_df)
    assert kpis["total_patients"] == 5
    assert kpis["outlier_count"] == 1
    assert kpis["outlier_pct"] == 20.0
    assert kpis["avg_treatment_cost"] == 241000.0

def test_monthly_trends(sample_test_df):
    trends = compute_monthly_trends(sample_test_df)
    assert not trends.empty
    assert "mean_cost" in trends.columns
    assert "mom_growth_pct" in trends.columns

def test_department_cagr(sample_test_df):
    cagr_df = compute_department_cagr(sample_test_df)
    assert not cagr_df.empty
    assert "dept_name" in cagr_df.columns
    assert "cagr_pct" in cagr_df.columns

def test_outlier_detection(sample_test_df):
    outliers = detect_outliers_iqr(sample_test_df, factor=1.5)
    # The extreme row with total_billed_cost=600000 must be identified
    assert not outliers.empty
    assert 600000 in outliers["total_billed_cost"].values

def test_los_correlation(sample_test_df):
    corr = compute_los_correlation(sample_test_df)
    assert "pearson_r" in corr
    # High LOS should strongly positively correlate with cost
    assert corr["pearson_r"] > 0.8

def test_forecast_model(sample_test_df):
    forecast = forecast_quarterly_trend(sample_test_df, horizon_quarters=2)
    assert not forecast.empty
    assert "Forecast Q+1" in forecast["quarter"].values
    assert "upper_bound" in forecast.columns

def test_data_cleaning():
    dirty_data = pd.DataFrame([
        {"billing_id": 1, "los_days": 0, "total_billed_cost": 1000, "insurance_paid": 800, "out_of_pocket_paid": 200, "insurance_type": None}
    ])
    cleaned = clean_hospital_data(dirty_data)
    assert cleaned.loc[0, "los_days"] == 1
    assert cleaned.loc[0, "insurance_type"] == "Unspecified"
    assert cleaned.loc[0, "out_of_pocket_pct"] == 20.0
