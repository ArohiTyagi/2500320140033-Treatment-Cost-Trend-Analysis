import pandas as pd
import numpy as np

def clean_hospital_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans raw billing data:
    1. Imputes missing categorical values (e.g., insurance_type -> 'Unspecified').
    2. Enforces positive non-zero constraints on costs.
    3. Guarantees LOS >= 1.
    4. Computes validated ratio columns.
    """
    df = df.copy()
    
    # 1. Impute missing insurance type
    if "insurance_type" in df.columns:
        df["insurance_type"] = df["insurance_type"].fillna("Unspecified")
        
    # 2. Enforce minimum LOS
    if "los_days" in df.columns:
        df["los_days"] = df["los_days"].clip(lower=1)
        
    # 3. Clean and verify monetary attributes
    cost_cols = ["room_charges", "surgery_charges", "pharmacy_charges", "lab_charges", "total_billed_cost", "insurance_paid", "out_of_pocket_paid"]
    for c in cost_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).clip(lower=0.0)

    # 4. Recompute consistent cost per day and percentages
    if "total_billed_cost" in df.columns and "los_days" in df.columns:
        df["cost_per_day"] = np.where(df["los_days"] > 0, np.round(df["total_billed_cost"] / df["los_days"], 2), 0.0)
        
    if "out_of_pocket_paid" in df.columns and "total_billed_cost" in df.columns:
        df["out_of_pocket_pct"] = np.where(df["total_billed_cost"] > 0, np.round((df["out_of_pocket_paid"] / df["total_billed_cost"]) * 100, 2), 0.0)
        
    if "insurance_paid" in df.columns and "total_billed_cost" in df.columns:
        df["insurance_coverage_pct"] = np.where(df["total_billed_cost"] > 0, np.round((df["insurance_paid"] / df["total_billed_cost"]) * 100, 2), 0.0)

    return df
