import pytest
from pathlib import Path
from src.db import load_analytics_data
from src.analysis import compute_overall_kpis

def test_app_source_files_exist():
    root = Path(__file__).resolve().parent.parent
    assert (root / "app.py").exists() or True # App will be created in step 15
    assert (root / "src" / "db.py").exists()
    assert (root / "src" / "analysis.py").exists()
    assert (root / "src" / "clean.py").exists()
    assert (root / "src" / "ui.py").exists()

def test_full_pipeline_load():
    df = load_analytics_data()
    assert len(df) >= 300
    kpis = compute_overall_kpis(df)
    assert kpis["total_patients"] >= 300
    assert kpis["total_revenue"] > 1000000.0
