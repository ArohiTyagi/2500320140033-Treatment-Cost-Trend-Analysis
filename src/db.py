import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "app.db"

def get_db_connection():
    """Returns a connection to the SQLite database with foreign keys enabled."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

@st.cache_data(ttl=3600)
def load_analytics_data():
    """Loads the pre-joined analytical view from SQLite into a Pandas DataFrame."""
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database file not found at: {DB_PATH}")
    conn = get_db_connection()
    query = "SELECT * FROM v_treatment_cost_analytics"
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    # Ensure correct datetime parsing
    df["admission_date"] = pd.to_datetime(df["admission_date"])
    df["discharge_date"] = pd.to_datetime(df["discharge_date"])
    return df

def load_table(table_name):
    """Loads an atomic table from SQLite for audit or inspection."""
    conn = get_db_connection()
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df
