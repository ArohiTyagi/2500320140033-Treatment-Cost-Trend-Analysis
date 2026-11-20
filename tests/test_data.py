import pytest
import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "app.db"

def test_database_file_exists():
    assert DB_PATH.exists(), f"Database file does not exist at: {DB_PATH}"

def test_tables_exist():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [row[0] for row in cur.fetchall()]
    conn.close()
    
    expected = ["dim_departments", "dim_procedures", "dim_facilities", "dim_patients", "fact_hospital_billing"]
    for t in expected:
        assert t in tables, f"Expected table '{t}' not found in database."

def test_fact_billing_record_count():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM fact_hospital_billing;")
    count = cur.fetchone()[0]
    conn.close()
    assert count >= 150, f"Expected at least 150 fact rows, but found {count}."

def test_primary_key_uniqueness():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(DISTINCT billing_id), COUNT(*) FROM fact_hospital_billing;")
    distinct_cnt, total_cnt = cur.fetchone()
    conn.close()
    assert distinct_cnt == total_cnt, "Duplicate primary keys found in fact_hospital_billing!"

def test_foreign_key_referential_integrity():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_key_check;")
    violations = cur.fetchall()
    conn.close()
    assert len(violations) == 0, f"Foreign key integrity violations detected: {violations}"

def test_cost_additive_balance():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT room_charges, surgery_charges, pharmacy_charges, lab_charges, total_billed_cost FROM fact_hospital_billing", conn)
    conn.close()
    
    computed_total = df["room_charges"] + df["surgery_charges"] + df["pharmacy_charges"] + df["lab_charges"]
    discrepancies = (computed_total.round(2) != df["total_billed_cost"].round(2)).sum()
    assert discrepancies == 0, f"Found {discrepancies} rows where component charges do not sum to total_billed_cost!"

def test_analytical_view_exists_and_populated():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM v_treatment_cost_analytics;")
    view_count = cur.fetchone()[0]
    conn.close()
    assert view_count > 0, "View v_treatment_cost_analytics is empty or uninitialized."
