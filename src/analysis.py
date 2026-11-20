import pandas as pd
import numpy as np
from scipy import stats

def compute_overall_kpis(df: pd.DataFrame) -> dict:
    """Computes headline KPIs for the executive dashboard."""
    if df.empty:
        return {
            "total_patients": 0,
            "total_revenue": 0.0,
            "avg_treatment_cost": 0.0,
            "median_treatment_cost": 0.0,
            "avg_los": 0.0,
            "outlier_count": 0,
            "outlier_pct": 0.0,
            "oop_pct": 0.0
        }
    
    total_rev = float(df["total_billed_cost"].sum())
    total_oop = float(df["out_of_pocket_paid"].sum())
    
    return {
        "total_patients": int(len(df)),
        "total_revenue": round(total_rev, 2),
        "avg_treatment_cost": round(float(df["total_billed_cost"].mean()), 2),
        "median_treatment_cost": round(float(df["total_billed_cost"].median()), 2),
        "avg_los": round(float(df["los_days"].mean()), 1),
        "outlier_count": int(df["is_outlier"].sum()),
        "outlier_pct": round(float(df["is_outlier"].mean() * 100), 1),
        "oop_pct": round((total_oop / total_rev * 100) if total_rev > 0 else 0.0, 1)
    }

def compute_monthly_trends(df: pd.DataFrame) -> pd.DataFrame:
    """Computes month-over-month average and median costs with growth percentage."""
    if df.empty:
        return pd.DataFrame()
    
    df_copy = df.copy()
    if not np.issubdtype(df_copy["admission_date"].dtype, np.datetime64):
        df_copy["admission_date"] = pd.to_datetime(df_copy["admission_date"])
        
    df_copy["year_month"] = df_copy["admission_date"].dt.to_period("M").astype(str)
    
    monthly = df_copy.groupby("year_month").agg(
        mean_cost=("total_billed_cost", "mean"),
        median_cost=("total_billed_cost", "median"),
        total_admissions=("billing_id", "count"),
        total_revenue=("total_billed_cost", "sum")
    ).reset_index()
    
    monthly["mean_cost"] = monthly["mean_cost"].round(2)
    monthly["median_cost"] = monthly["median_cost"].round(2)
    monthly["mom_growth_pct"] = monthly["mean_cost"].pct_change().fillna(0.0) * 100
    monthly["mom_growth_pct"] = monthly["mom_growth_pct"].round(2)
    
    return monthly

def compute_department_cagr(df: pd.DataFrame) -> pd.DataFrame:
    """Computes the Compound Annual Growth Rate (CAGR) of mean treatment cost by department."""
    if df.empty:
        return pd.DataFrame()
        
    df_copy = df.copy()
    if not np.issubdtype(df_copy["admission_date"].dtype, np.datetime64):
        df_copy["admission_date"] = pd.to_datetime(df_copy["admission_date"])
        
    df_copy["year"] = df_copy["admission_date"].dt.year
    min_year = df_copy["year"].min()
    max_year = df_copy["year"].max()
    num_years = max(1, max_year - min_year)
    
    yearly_dept = df_copy.groupby(["dept_name", "year"])["total_billed_cost"].mean().unstack()
    
    results = []
    for dept, row in yearly_dept.iterrows():
        start_val = row.get(min_year, np.nan)
        end_val = row.get(max_year, np.nan)
        if pd.notna(start_val) and pd.notna(end_val) and start_val > 0:
            cagr = ((end_val / start_val) ** (1.0 / num_years) - 1.0) * 100
        else:
            cagr = 0.0
        results.append({
            "dept_name": dept,
            "start_year_avg": round(start_val, 2) if pd.notna(start_val) else 0.0,
            "end_year_avg": round(end_val, 2) if pd.notna(end_val) else 0.0,
            "cagr_pct": round(cagr, 2)
        })
        
    return pd.DataFrame(results).sort_values(by="cagr_pct", ascending=False)

def detect_outliers_iqr(df: pd.DataFrame, factor=1.5) -> pd.DataFrame:
    """Detects cost outliers using Tukey's Fences IQR criteria per department or globally."""
    if df.empty:
        return pd.DataFrame()
        
    outliers_list = []
    # If grouped by department has too few rows, compute on the cohort
    use_grouping = (len(df) >= 20) and ("dept_name" in df.columns)
    groups = df.groupby("dept_name") if use_grouping else [("All", df)]
    
    for name, group in groups:
        q1 = group["total_billed_cost"].quantile(0.25)
        q3 = group["total_billed_cost"].quantile(0.75)
        iqr = q3 - q1
        upper_fence = q3 + (factor * iqr)
        lower_fence = max(0, q1 - (factor * iqr))
        
        dept_outliers = group[(group["total_billed_cost"] > upper_fence) | (group["total_billed_cost"] < lower_fence)].copy()
        if not dept_outliers.empty:
            dept_outliers["q1"] = round(q1, 2)
            dept_outliers["q3"] = round(q3, 2)
            dept_outliers["iqr"] = round(iqr, 2)
            dept_outliers["upper_fence"] = round(upper_fence, 2)
            
            std = group["total_billed_cost"].std()
            mean = group["total_billed_cost"].mean()
            dept_outliers["z_score"] = np.round((dept_outliers["total_billed_cost"] - mean) / (std if std > 0 else 1.0), 2)
            outliers_list.append(dept_outliers)
        
    if outliers_list:
        return pd.concat(outliers_list, ignore_index=True)
    return pd.DataFrame()

def compute_los_correlation(df: pd.DataFrame) -> dict:
    """Computes Pearson and Spearman correlation between Length of Stay (LOS) and total cost."""
    if len(df) < 2:
        return {"pearson_r": 0.0, "p_value": 1.0, "spearman_r": 0.0}
        
    p_r, p_val = stats.pearsonr(df["los_days"], df["total_billed_cost"])
    s_r, _ = stats.spearmanr(df["los_days"], df["total_billed_cost"])
    return {
        "pearson_r": round(float(p_r), 3),
        "p_value": float(p_val),
        "spearman_r": round(float(s_r), 3)
    }

def forecast_quarterly_trend(df: pd.DataFrame, horizon_quarters=4) -> pd.DataFrame:
    """Forecasts next quarterly treatment costs using a rolling moving-average model."""
    if df.empty:
        return pd.DataFrame()
        
    df_copy = df.copy()
    if not np.issubdtype(df_copy["admission_date"].dtype, np.datetime64):
        df_copy["admission_date"] = pd.to_datetime(df_copy["admission_date"])
        
    df_copy["quarter"] = df_copy["admission_date"].dt.to_period("Q").astype(str)
    
    q_data = df_copy.groupby("quarter")["total_billed_cost"].mean().reset_index()
    q_data.columns = ["quarter", "avg_cost"]
    
    # Calculate simple linear trend
    x = np.arange(len(q_data))
    y = q_data["avg_cost"].values
    
    if len(x) > 1:
        slope, intercept = np.polyfit(x, y, 1)
    else:
        slope, intercept = 0, y[0] if len(y) > 0 else 0
        
    last_idx = len(x)
    forecast_rows = []
    for i in range(1, horizon_quarters + 1):
        future_idx = last_idx + i - 1
        predicted_val = (slope * future_idx) + intercept
        q_label = f"Forecast Q+{i}"
        forecast_rows.append({
            "quarter": q_label,
            "avg_cost": round(predicted_val, 2),
            "lower_bound": round(predicted_val * 0.90, 2),
            "upper_bound": round(predicted_val * 1.10, 2),
            "type": "Forecast"
        })
        
    hist_df = q_data.copy()
    hist_df["lower_bound"] = hist_df["avg_cost"]
    hist_df["upper_bound"] = hist_df["avg_cost"]
    hist_df["type"] = "Historical"
    
    return pd.concat([hist_df, pd.DataFrame(forecast_rows)], ignore_index=True)
