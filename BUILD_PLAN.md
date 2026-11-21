# BUILD PLAN: Treatment Cost Trend Analysis

## 1. Project Summary
The **Treatment Cost Trend Analysis** platform is an enterprise-grade Data Analytics & Inpatient Billing Analytics dashboard. Engineered for Arohi Tyagi (Roll No: 2500320140033) in the MCA 2nd Year Data Analytics program at ABES Engineering College under mentor Prof. Kalpana Dwivedi. The solution addresses medical inflation unpredictability, unexplained departmental cost drift, and billing anomalies across hospital networks. The system connects an embedded SQLite database (`data/app.db`) to a modular Python analytical engine and renders interactive Plotly visualizations within a Tailwind-styled Streamlit multi-page interface. The platform tracks month-over-month expenditure trends, calculates departmental CAGR, correlates Length of Stay (LOS) with patient charges, detects statistical outliers via Tukey's IQR fences and Z-scores, and forecasts rolling 90-day expenditures. Rigorous quality assurance is maintained with automated pytest suites achieving a 100% pass rate. The project culminates in institutional ABES Word documentation and a defense slide deck.

### Document Index:
- [Step 1: Project Understanding](docs/01_Project_Understanding.docx)
- [Step 2: HMW Statements](docs/02_HMW_Statements.docx)
- [Step 3: Needs Analysis](docs/03_Needs_Analysis.docx)
- [Step 4: Traceability Matrix](docs/04_Traceability_Matrix.xlsx)
- [Step 5: Business Requirements Document](docs/05_BRD.docx)
- [Step 7: Technical Requirements Document](docs/07_TRD.docx)
- [Step 8: Database Design](docs/08_Database_Design.xlsx)
- [Step 9: ER Diagram & Architecture](docs/09_ER_Diagram.docx)
- [Step 10: Synthetic Dataset](data/treatment_cost_data.xlsx)
- [Step 11: Agile User Stories](docs/11_User_Stories.docx)
- [Step 12: UI Framework Specification](docs/12_UI_Framework.docx)
- [Step 13: Test Cases & Execution Report](docs/13_Test_Cases.docx)

---

## 2. Directory Structure
```text
033TrendCostAnalysis/
├── .streamlit/
│   └── config.toml
├── app.py
├── pages/
│   ├── 1_Overview.py
│   ├── 2_Cost_Trends.py
│   ├── 3_Outlier_Detection.py
│   ├── 4_Procedure_Analysis.py
│   ├── 5_Forecasting.py
│   └── 6_Data_Explorer.py
├── src/
│   ├── __init__.py
│   ├── db.py
│   ├── clean.py
│   ├── analysis.py
│   ├── ui.py
│   └── generate_data.py
├── data/
│   ├── app.db
│   ├── treatment_cost_data.xlsx
│   └── *.csv
├── tests/
│   ├── test_data.py
│   ├── test_analysis.py
│   └── test_app.py
├── docs/
├── requirements.txt
└── README.md
```

---

## 3. Granular Build Tasks

### BUILD-01: Environment Configuration & Dependencies
- **Goal:** Pin dependencies and configure Streamlit theme settings.
- **Files:** `requirements.txt`, `.streamlit/config.toml`.
- **Linked:** NFR-01, FEAT-14.
- **Verification:** Run `pip install -r requirements.txt`.

### BUILD-02: SQLite Schema & Data Access Layer
- **Goal:** Build SQLite connection, cached query loader, and relational view extraction.
- **Files:** `src/db.py`, `data/app.db`.
- **Linked:** BR-01, US-12, FEAT-13.
- **Verification:** Run `pytest tests/test_data.py`.

### BUILD-03: Data Cleaning & Preprocessing Pipeline
- **Goal:** Implement data cleaning, missing value imputation, and ratio validation.
- **Files:** `src/clean.py`.
- **Linked:** BR-01, BR-04, US-04.
- **Verification:** `pytest tests/test_analysis.py::test_data_cleaning`.

### BUILD-04: Analytical Algorithms & Statistical Calculations
- **Goal:** Implement KPI computation, MoM trends, CAGR, Tukey IQR outliers, and forecasting.
- **Files:** `src/analysis.py`.
- **Linked:** BR-01, BR-02, BR-03, BR-05, BR-07, US-01 through US-07.
- **Verification:** `pytest tests/test_analysis.py`.

### BUILD-05: Reusable Tailwind UI Design Components
- **Goal:** Inject Tailwind CSS CDN and render responsive executive headers, KPI cards, and insight banners.
- **Files:** `src/ui.py`.
- **Linked:** UI-01 to UI-06.
- **Verification:** Import and render components in Streamlit iframe without CSS leaks.

### BUILD-06: Executive Overview Dashboard (UI-01)
- **Goal:** Build main overview page with KPI cards, charge component waterfall, and facility benchmarks.
- **Files:** `app.py`, `pages/1_Overview.py`.
- **Linked:** US-01, US-08, US-10.
- **Verification:** Page loads and updates dynamically upon filter changes.

### BUILD-07: Longitudinal Cost Trends & Departmental Inflation (UI-02)
- **Goal:** Build time-series line charts and CAGR inflation ranking bar charts.
- **Files:** `pages/2_Cost_Trends.py`.
- **Linked:** US-01, US-02.
- **Verification:** MoM trajectory and CAGR table render with 0 errors.

### BUILD-08: Outlier Detection & Billing Audit Screen (UI-03)
- **Goal:** Deliver interactive Z-score slider, Tukey IQR boxplots, and downloadable audit table.
- **Files:** `pages/3_Outlier_Detection.py`.
- **Linked:** US-06, US-09.
- **Verification:** Highlighting and CSV export functional.

### BUILD-09: Procedure Variance & LOS Correlation Screen (UI-04)
- **Goal:** Deliver procedure IQR boxes, LOS vs Cost scatter plots, and Pearson r metrics.
- **Files:** `pages/4_Procedure_Analysis.py`.
- **Linked:** US-03, US-05.
- **Verification:** Scatter regression slope renders.

### BUILD-10: Forecasting, Scenario Simulation & Explorer (UI-05, UI-06)
- **Goal:** Provide rolling 90-day forecast with confidence bands, fee cap sliders, and raw tabular data explorer.
- **Files:** `pages/5_Forecasting.py`, `pages/6_Data_Explorer.py`.
- **Linked:** US-07, US-11, US-12.
- **Verification:** `pytest tests/` runs with 100% pass rate.

---

## 4. Definition of Done (DoD)
1. All 7 Must-Have Business Requirements (BR-01 through BR-08) fully functioning in the dashboard.
2. 100% of automated unit and functional tests passing via pytest.
3. Zero critical runtime errors or memory leaks during navigation.
4. Complete documentation set saved in `docs/` according to ABES rubrics.
5. Repository prepared for zero-downtime deployment on Streamlit Community Cloud.

---

## 5. Execution Commands
- **Run Application Locally:** `streamlit run app.py`
- **Run Test Suite:** `pytest tests/ -v`
- **Regenerate Synthetic Data:** `python src/generate_data.py`

---

## 6. Viva Defense Preparation (Top 10 Questions & Answers)
1. **Q: What is the core business objective of this application?**  
   *A:* To track inpatient treatment cost trends, calculate departmental medical inflation, correlate Length of Stay with hospital costs, and flag statistical billing outliers before claim submission.
2. **Q: Why combine Mean and Median in cost trend analysis?**  
   *A:* Healthcare expenditures are heavily right-skewed by intensive ICU admissions and surgical complications. Mean reflects total financial liability, while Median provides an outlier-resilient central benchmark.
3. **Q: What method is used for statistical outlier detection?**  
   *A:* Tukey's Fences IQR criteria ($Q3 + 1.5 \times \text{IQR}$) complemented by standardized Z-scores ($Z > 3.0$).
4. **Q: How does SQLite maintain referential integrity in your architecture?**  
   *A:* Through explicit foreign keys enforced at runtime via `PRAGMA foreign_keys = ON;` and validation tests in pytest.
5. **Q: What is the benefit of the pre-joined view `v_treatment_cost_analytics`?**  
   *A:* It flattens the dimensional snowflake model into a single performant data mart, eliminating repetitive multi-table SQL joins in Streamlit.
6. **Q: How does caching improve Streamlit performance?**  
   *A:* `@st.cache_data(ttl=3600)` retains the loaded dataframe in server memory, achieving sub-second UI reactivity across filters.
7. **Q: How is CAGR calculated?**  
   *A:* $\text{CAGR} = (\text{Cost}_{\text{End}} / \text{Cost}_{\text{Start}})^{1/n} - 1$, measuring annualized departmental inflation.
8. **Q: How are out-of-pocket patient costs calculated?**  
   *A:* Total Billed Cost minus Insurance Paid amount, ensuring 100% settlement parity.
9. **Q: What testing strategy is employed?**  
   *A:* Automated unit and integration testing via pytest verifying schema integrity, mathematical formulas, and full pipeline execution.
10. **Q: How is the app deployed?**  
    *A:* Deployed from a public GitHub repository directly onto Streamlit Community Cloud using pinned `requirements.txt`.
