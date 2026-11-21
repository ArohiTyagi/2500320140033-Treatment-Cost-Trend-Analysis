# 🏥 Treatment Cost Trend Analysis (Project ID: P106)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://2500320140033-treatment-cost-analysis.streamlit.app)
[![Python 3.12](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tests: Pytest](https://img.shields.io/badge/Tests-16%20Passed-brightgreen.svg)](https://pytest.org/)

An enterprise-grade **Data Analyst & Inpatient Billing Analytics Platform** engineered to track longitudinal medical cost inflation, evaluate Length of Stay (LOS) dynamics, identify anomalous billing patterns, and project departmental expenditures.

---

## 👨‍🎓 Project Details
- **Student / Author:** Arohi Tyagi
- **Roll Number:** 2500320140033
- **Program:** MCA 2nd Year (Data Analyst)
- **Institution:** ABES Engineering College, Ghaziabad
- **Academic Mentor:** Mr. Dinesh
- **Repository:** [2500320140033-Treatment-Cost-Trend-Analysis](https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis.git)
- **Live Deployment:** [Streamlit Community Cloud](https://2500320140033-treatment-cost-analysis.streamlit.app)

---

## 🌟 Key Features
- **Executive Overview Dashboard (`app.py`):** Real-time KPI summary cards, four-part inpatient charge component decomposition (Waterfall), and multi-facility daily cost benchmarks.
- **Longitudinal Cost Trends (`pages/1_Cost_Trends.py`):** Month-over-month trajectory comparing Average Cost against outlier-resilient Median Cost, alongside departmental Compound Annual Growth Rate (CAGR) inflation rankings.
- **Statistical Outlier Detection (`pages/2_Outlier_Detection.py`):** Automated claim audit identifying extreme billing amounts exceeding Tukey's Interquartile Range fences ($Q3 + 1.5 \times \text{IQR}$) and standardized Z-scores ($Z > 2.5$) with CSV export.
- **Procedure Variance & LOS Correlation (`pages/3_Procedure_Analysis.py`):** Interquartile procedural dispersion boxplots and bivariate scatter regression quantifying the relationship between bed days and total billed cost.
- **Predictive Forecasting & Sensitivity Simulation (`pages/4_Forecasting.py`):** Rolling 90-day quarterly expenditure forecasts with confidence ribbons and an interactive price cap adjustment slider.
- **Relational Data Explorer (`pages/5_Data_Explorer.py`):** Searchable tabular data mart viewer with entity schema audits and filtered dataset downloads.

---

## 🛠️ Technology Stack
- **Application Framework:** Streamlit 1.38+
- **Data Science & Statistics:** Pandas, NumPy, SciPy, Statsmodels
- **Visual Analytics:** Plotly Express & Plotly Graph Objects
- **Styling:** Tailwind CSS (via CDN in `src/ui.py`)
- **Database:** SQLite 3 (`data/app.db`) enforcing foreign key integrity
- **Quality Assurance:** Pytest (16 automated unit & integration tests)

---

## 🚀 Quick Start (Local Setup)

```bash
# 1. Clone the repository
git clone https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis.git
cd 2500320140033-Treatment-Cost-Trend-Analysis

# 2. Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate   # On Windows
# source .venv/bin/activate # On macOS/Linux

# 3. Install pinned dependencies
pip install -r requirements.txt

# 4. Run automated test suite
pytest tests/ -v

# 5. Launch the Streamlit application
streamlit run app.py
```

---

## 🧪 Automated Test Suite
All test cases are automated using `pytest` verifying database schema integrity, primary key uniqueness, foreign key consistency, additive cost balance, and statistical calculation algorithms.

```text
============================= test session starts =============================
collected 16 items

tests/test_analysis.py::test_kpi_computation PASSED                      [  6%]
tests/test_analysis.py::test_monthly_trends PASSED                       [ 12%]
tests/test_analysis.py::test_department_cagr PASSED                      [ 18%]
tests/test_analysis.py::test_outlier_detection PASSED                    [ 25%]
tests/test_analysis.py::test_los_correlation PASSED                      [ 31%]
tests/test_analysis.py::test_forecast_model PASSED                       [ 37%]
tests/test_analysis.py::test_data_cleaning PASSED                        [ 43%]
tests/test_app.py::test_app_source_files_exist PASSED                    [ 50%]
tests/test_app.py::test_full_pipeline_load PASSED                        [ 56%]
tests/test_data.py::test_database_file_exists PASSED                     [ 62%]
tests/test_data.py::test_tables_exist PASSED                             [ 68%]
tests/test_data.py::test_fact_billing_record_count PASSED                [ 75%]
tests/test_data.py::test_primary_key_uniqueness PASSED                   [ 81%]
tests/test_data.py::test_foreign_key_referential_integrity PASSED        [ 87%]
tests/test_data.py::test_cost_additive_balance PASSED                    [ 93%]
tests/test_data.py::test_analytical_view_exists_and_populated PASSED     [100%]

============================= 16 passed in 2.76s ==============================
```

---

## 📁 Repository Structure
```text
├── .streamlit/
│   └── config.toml             # ABES theme colors & styling settings
├── app.py                      # Main application entrypoint
├── pages/                      # Multi-page analytics modules
├── src/                        # Modular data access, cleaning & analysis libraries
├── data/
│   ├── app.db                  # Embedded SQLite database
│   └── treatment_cost_data.xlsx# Multi-entity Excel dataset
├── tests/                      # Automated pytest test suites
├── docs/                       # Complete 18-step project documentation & slide deck
├── requirements.txt            # Pinned runtime dependencies
├── BUILD_PLAN.md               # Step-by-step engineering roadmap & viva notes
└── README.md                   # Project documentation
```

---

## 📜 Academic Deliverables
- **ABES Project Document:** [`docs/P106_Project_Document.docx`](docs/P106_Project_Document.docx)
- **ABES Defense Slide Deck:** [`docs/P106_Presentation.pptx`](docs/P106_Presentation.pptx)
- **Requirements Traceability Matrix:** [`docs/04_Traceability_Matrix.xlsx`](docs/04_Traceability_Matrix.xlsx)
- **Database Design Specification:** [`docs/08_Database_Design.xlsx`](docs/08_Database_Design.xlsx)
