import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_step08_db():
    wb = openpyxl.Workbook()
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    alt_fill = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    def format_sheet(ws, title, headers, rows):
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells("A1:G1")
        t_cell = ws["A1"]
        t_cell.value = title
        t_cell.font = Font(name="Calibri", size=13, bold=True, color="1B365D")
        t_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        ws.row_dimensions[3].height = 22
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for r_idx, r_data in enumerate(rows, 4):
            ws.row_dimensions[r_idx].height = 20
            is_even = (r_idx % 2 == 0)
            for c_idx, val in enumerate(r_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = regular_font
                cell.border = thin_border
                if is_even:
                    cell.fill = alt_fill
                if c_idx in [1, 3, 4, 5]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row < 3:
                    continue
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # Sheet 1: Overview
    ws_ov = wb.active
    ws_ov.title = "1_Overview"
    ws_ov.views.sheetView[0].showGridLines = True
    ws_ov["A1"] = "DATABASE DESIGN SPECIFICATION — TREATMENT COST TREND ANALYSIS"
    ws_ov["A1"].font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    ws_ov["A2"] = "Relational SQLite schema design (Star/Snowflake hybrid) with foreign keys and view definitions."
    ws_ov["A2"].font = Font(name="Calibri", size=11, italic=True, color="5A6978")

    ov_rows = [
        ("Architecture Pattern", "Star Schema Dimensional Model (Fact + 4 Conformed Dimensions)"),
        ("Database Engine", "SQLite 3.x (Serverless, portable embedded file: data/app.db)"),
        ("Referential Integrity", "Enforced via 'PRAGMA foreign_keys = ON;' on all database connections"),
        ("Main Fact Table", "fact_hospital_billing (Target row grain: individual inpatient discharge bill)"),
        ("Planned Fact Rows", "250 to 500 records (MVP baseline) up to 3,000 for longitudinal stress testing"),
        ("Dimension Tables", "dim_departments, dim_procedures, dim_facilities, dim_patients"),
        ("Primary Analytical View", "v_treatment_cost_analytics (Denormalized 32-column view for Streamlit)"),
        ("Operational Coverage", "Supports all 10 Operational Questions (OQ-01 through OQ-10)")
    ]
    for idx, (k, v) in enumerate(ov_rows, 4):
        ws_ov.cell(row=idx, column=1, value=k).font = Font(name="Calibri", size=10, bold=True)
        ws_ov.cell(row=idx, column=2, value=v).font = regular_font
        ws_ov.row_dimensions[idx].height = 20
    ws_ov.column_dimensions["A"].width = 28
    ws_ov.column_dimensions["B"].width = 75

    # Sheet 2: Tables_Summary
    ws_ts = wb.create_sheet(title="2_Tables_Summary")
    ts_headers = ["Table Name", "Table Grain", "Primary Key (PK)", "Foreign Keys (FK)", "Column Count", "Planned Rows"]
    ts_data = [
        ("dim_departments", "1 row per medical department", "dept_id", "None", 4, 6),
        ("dim_procedures", "1 row per clinical procedure", "procedure_id", "dept_id -> dim_departments", 6, 18),
        ("dim_facilities", "1 row per hospital branch/facility", "facility_id", "None", 5, 5),
        ("dim_patients", "1 row per registered patient", "patient_id", "None", 6, 250),
        ("fact_hospital_billing", "1 row per inpatient admission & billing event", "billing_id", "patient_id, procedure_id, facility_id", 15, 300)
    ]
    format_sheet(ws_ts, "TABLES SUMMARY & GRAIN MATRIX", ts_headers, ts_data)

    # Sheet 3: Dict_fact_billing
    ws_fb = wb.create_sheet(title="3_Dict_Fact_Billing")
    fb_headers = ["Column Name", "Data Type", "Key Type", "Nullable", "Allowed Values / Range", "Description", "OQ / BR Supported"]
    fb_data = [
        ("billing_id", "INTEGER", "PK", "No", "10001 - 99999", "Unique sequential billing invoice identifier", "Core PK"),
        ("patient_id", "INTEGER", "FK", "No", "References dim_patients(patient_id)", "Foreign key to patient profile", "OQ-03, BR-03"),
        ("procedure_id", "INTEGER", "FK", "No", "References dim_procedures(procedure_id)", "Foreign key to treatment procedure", "OQ-02, OQ-05"),
        ("facility_id", "INTEGER", "FK", "No", "References dim_facilities(facility_id)", "Foreign key to hospital branch", "OQ-10, BR-10"),
        ("admission_date", "DATE", "None", "No", "2023-01-01 to 2026-12-31", "Date patient was admitted to inpatient bed", "OQ-01, BR-01"),
        ("discharge_date", "DATE", "None", "No", ">= admission_date", "Date patient was discharged", "OQ-01, BR-01"),
        ("los_days", "INTEGER", "None", "No", "1 to 45", "Length of stay in days: max(1, discharge - admission)", "OQ-03, BR-03"),
        ("admission_type", "VARCHAR", "None", "No", "Emergency, Elective, Urgent", "Acuity level at hospital admission", "OQ-09, BR-09"),
        ("room_charges", "DECIMAL(10,2)", "None", "No", ">= 0.00", "Inpatient room and nursing care fee", "OQ-08, BR-08"),
        ("surgery_charges", "DECIMAL(10,2)", "None", "No", ">= 0.00", "Surgical theatre, surgeon, and anesthesia charges", "OQ-08, BR-08"),
        ("pharmacy_charges", "DECIMAL(10,2)", "None", "No", ">= 0.00", "Medication and consumable expenditure", "OQ-08, BR-08"),
        ("lab_charges", "DECIMAL(10,2)", "None", "No", ">= 0.00", "Diagnostic radiology and pathology fees", "OQ-08, BR-08"),
        ("total_billed_cost", "DECIMAL(10,2)", "None", "No", "= room + surgery + pharmacy + lab", "Total gross hospitalization billed amount", "All OQs, BR-01"),
        ("insurance_paid", "DECIMAL(10,2)", "None", "No", "0.00 to total_billed_cost", "Reimbursement covered by insurance provider", "OQ-04, BR-04"),
        ("out_of_pocket_paid", "DECIMAL(10,2)", "None", "No", "= total_billed_cost - insurance_paid", "Patient co-pay / out-of-pocket settlement", "OQ-04, BR-04"),
        ("insurance_type", "VARCHAR", "None", "No", "Corporate, Private, Government, Uninsured", "Insurance scheme category", "OQ-04, BR-04"),
        ("is_outlier", "BOOLEAN", "None", "No", "0 (Normal) or 1 (Outlier)", "Calculated flag for charges exceeding Tukey's IQR fence", "OQ-06, BR-05")
    ]
    format_sheet(ws_fb, "DATA DICTIONARY — fact_hospital_billing", fb_headers, fb_data)

    # Sheet 4: Dict_Dimensions
    ws_dd = wb.create_sheet(title="4_Dict_Dimensions")
    dd_headers = ["Table Name", "Column Name", "Data Type", "Key", "Nullable", "Description", "OQ / BR Supported"]
    dd_data = [
        ("dim_departments", "dept_id", "INTEGER", "PK", "No", "Unique department primary identifier", "OQ-01, OQ-02"),
        ("dim_departments", "dept_name", "VARCHAR(50)", "None", "No", "Cardiology, Orthopedics, Neurology, Oncology, etc.", "OQ-01, OQ-02"),
        ("dim_departments", "hod_name", "VARCHAR(50)", "None", "No", "Head of Clinical Department", "Governance"),
        ("dim_departments", "budget_cap_inr", "DECIMAL(12,2)", "None", "No", "Annual departmental target expenditure cap", "BR-01, BR-02"),
        ("dim_procedures", "procedure_id", "INTEGER", "PK", "No", "Unique procedure primary identifier", "OQ-05, BR-06"),
        ("dim_procedures", "dept_id", "INTEGER", "FK", "No", "Foreign key referencing dim_departments(dept_id)", "OQ-02, OQ-05"),
        ("dim_procedures", "procedure_name", "VARCHAR(100)", "None", "No", "Angioplasty, Knee Replacement, MRI Brain, etc.", "OQ-05, BR-06"),
        ("dim_procedures", "base_cost_inr", "DECIMAL(10,2)", "None", "No", "Baseline standard protocol cost", "OQ-05, BR-06"),
        ("dim_procedures", "complexity_level", "VARCHAR(20)", "None", "No", "Low, Medium, High", "OQ-05, BR-06"),
        ("dim_facilities", "facility_id", "INTEGER", "PK", "No", "Unique hospital location identifier", "OQ-10, BR-10"),
        ("dim_facilities", "facility_name", "VARCHAR(50)", "None", "No", "Max Healthcare Noida, Apollo Delhi, Fortis Gurugram", "OQ-10, BR-10"),
        ("dim_facilities", "city", "VARCHAR(50)", "None", "No", "Noida, Delhi, Gurugram, Ghaziabad", "OQ-10, BR-10"),
        ("dim_facilities", "bed_capacity", "INTEGER", "None", "No", "Total licensed inpatient beds", "OQ-10, BR-10"),
        ("dim_patients", "patient_id", "INTEGER", "PK", "No", "Unique synthetic patient profile identifier", "OQ-03, BR-03"),
        ("dim_patients", "patient_name", "VARCHAR(50)", "None", "No", "Fictional Indian full name", "Demographics"),
        ("dim_patients", "age", "INTEGER", "None", "No", "1 to 90 years", "OQ-03, BR-03"),
        ("dim_patients", "gender", "VARCHAR(10)", "None", "No", "Male, Female, Other", "OQ-03, BR-03"),
        ("dim_patients", "city", "VARCHAR(50)", "None", "No", "Patient residence city", "Demographics")
    ]
    format_sheet(ws_dd, "DATA DICTIONARY — DIMENSION TABLES", dd_headers, dd_data)

    # Sheet 5: Relationships
    ws_rel = wb.create_sheet(title="5_Relationships")
    rel_headers = ["Parent Table", "Child Table", "Join Key", "Cardinality", "Integrity Enforcement Rule"]
    rel_data = [
        ("dim_departments", "dim_procedures", "dept_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT ON UPDATE CASCADE"),
        ("dim_departments", "fact_hospital_billing", "Via dim_procedures.dept_id", "1 : N (Transitive One-to-Many)", "Referential integrity via procedure_id"),
        ("dim_procedures", "fact_hospital_billing", "procedure_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT ON UPDATE CASCADE"),
        ("dim_facilities", "fact_hospital_billing", "facility_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT ON UPDATE CASCADE"),
        ("dim_patients", "fact_hospital_billing", "patient_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT ON UPDATE CASCADE")
    ]
    format_sheet(ws_rel, "RELATIONSHIPS & REFERENTIAL INTEGRITY", rel_headers, rel_data)

    # Sheet 6: KPIs_Measures
    ws_kpi = wb.create_sheet(title="6_KPIs_Measures")
    kpi_headers = ["KPI / Measure Name", "Mathematical / SQL Formula", "Linked OQ", "Business Purpose"]
    kpi_data = [
        ("Average Treatment Cost", "AVG(total_billed_cost)", "OQ-01", "Monitors overall cost drift per patient"),
        ("Median Treatment Cost", "MEDIAN(total_billed_cost)", "OQ-01", "Robust central tendency ignoring skew"),
        ("Annualized CAGR", "POWER((Cost_t / Cost_0), (1 / years)) - 1", "OQ-02", "Tracks medical inflation rate by department"),
        ("Average Length of Stay (ALOS)", "AVG(los_days)", "OQ-03", "Assesses inpatient bed utilization efficiency"),
        ("Out-of-Pocket Ratio", "SUM(out_of_pocket_paid) / SUM(total_billed_cost) * 100", "OQ-04", "Measures patient direct financial burden"),
        ("Interquartile Range (IQR)", "Percentile(total_cost, 0.75) - Percentile(total_cost, 0.25)", "OQ-05", "Measures procedural cost dispersion"),
        ("Outlier Ratio", "COUNT(CASE WHEN is_outlier = 1 THEN 1 END) / COUNT(*) * 100", "OQ-06", "Quantifies suspicious/anomalous billing rate"),
        ("Rolling 90-Day Forecast", "AVG(total_cost) OVER (ORDER BY date ROWS 2 PRECEDING)", "OQ-07", "Projects next quarter operational expenditure"),
        ("Room Share %", "SUM(room_charges) / SUM(total_billed_cost) * 100", "OQ-08", "Evaluates facility accommodation overhead"),
        ("Emergency Cost Premium", "AVG(Cost_Emergency) - AVG(Cost_Elective)", "OQ-09", "Isolates acute care cost differential"),
        ("Facility Daily Rate", "SUM(total_billed_cost) / SUM(los_days)", "OQ-10", "Benchmarks hospital unit productivity")
    ]
    format_sheet(ws_kpi, "KEY PERFORMANCE INDICATORS & FORMULAS", kpi_headers, kpi_data)

    # Sheet 7: SQLite_DDL
    ws_ddl = wb.create_sheet(title="7_SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    ws_ddl["A1"] = "SQLITE DATABASE DEFINITION (DDL)"
    ws_ddl["A1"].font = Font(name="Calibri", size=13, bold=True, color="1B365D")

    ddl_statements = [
        "PRAGMA foreign_keys = ON;",
        "",
        "-- 1. Dimension: Departments",
        "CREATE TABLE dim_departments (",
        "    dept_id INTEGER PRIMARY KEY,",
        "    dept_name TEXT NOT NULL UNIQUE,",
        "    hod_name TEXT NOT NULL,",
        "    budget_cap_inr REAL NOT NULL CHECK(budget_cap_inr > 0)",
        ");",
        "",
        "-- 2. Dimension: Procedures",
        "CREATE TABLE dim_procedures (",
        "    procedure_id INTEGER PRIMARY KEY,",
        "    dept_id INTEGER NOT NULL,",
        "    procedure_name TEXT NOT NULL,",
        "    base_cost_inr REAL NOT NULL CHECK(base_cost_inr > 0),",
        "    complexity_level TEXT NOT NULL CHECK(complexity_level IN ('Low', 'Medium', 'High')),",
        "    FOREIGN KEY (dept_id) REFERENCES dim_departments(dept_id) ON DELETE RESTRICT",
        ");",
        "",
        "-- 3. Dimension: Facilities",
        "CREATE TABLE dim_facilities (",
        "    facility_id INTEGER PRIMARY KEY,",
        "    facility_name TEXT NOT NULL UNIQUE,",
        "    city TEXT NOT NULL,",
        "    bed_capacity INTEGER NOT NULL CHECK(bed_capacity > 0)",
        ");",
        "",
        "-- 4. Dimension: Patients",
        "CREATE TABLE dim_patients (",
        "    patient_id INTEGER PRIMARY KEY,",
        "    patient_name TEXT NOT NULL,",
        "    age INTEGER NOT NULL CHECK(age BETWEEN 1 AND 120),",
        "    gender TEXT NOT NULL CHECK(gender IN ('Male', 'Female', 'Other')),",
        "    city TEXT NOT NULL",
        ");",
        "",
        "-- 5. Fact Table: Hospital Billing",
        "CREATE TABLE fact_hospital_billing (",
        "    billing_id INTEGER PRIMARY KEY,",
        "    patient_id INTEGER NOT NULL,",
        "    procedure_id INTEGER NOT NULL,",
        "    facility_id INTEGER NOT NULL,",
        "    admission_date TEXT NOT NULL,",
        "    discharge_date TEXT NOT NULL,",
        "    los_days INTEGER NOT NULL CHECK(los_days >= 1),",
        "    admission_type TEXT NOT NULL CHECK(admission_type IN ('Emergency', 'Elective', 'Urgent')),",
        "    room_charges REAL NOT NULL CHECK(room_charges >= 0),",
        "    surgery_charges REAL NOT NULL CHECK(surgery_charges >= 0),",
        "    pharmacy_charges REAL NOT NULL CHECK(pharmacy_charges >= 0),",
        "    lab_charges REAL NOT NULL CHECK(lab_charges >= 0),",
        "    total_billed_cost REAL NOT NULL CHECK(total_billed_cost > 0),",
        "    insurance_paid REAL NOT NULL CHECK(insurance_paid >= 0),",
        "    out_of_pocket_paid REAL NOT NULL CHECK(out_of_pocket_paid >= 0),",
        "    insurance_type TEXT NOT NULL CHECK(insurance_type IN ('Corporate', 'Private', 'Government', 'Uninsured')),",
        "    is_outlier INTEGER NOT NULL DEFAULT 0 CHECK(is_outlier IN (0, 1)),",
        "    FOREIGN KEY (patient_id) REFERENCES dim_patients(patient_id),",
        "    FOREIGN KEY (procedure_id) REFERENCES dim_procedures(procedure_id),",
        "    FOREIGN KEY (facility_id) REFERENCES dim_facilities(facility_id)",
        ");",
        "",
        "-- 6. Analytical View (Joined Data Mart)",
        "CREATE VIEW v_treatment_cost_analytics AS",
        "SELECT",
        "    b.billing_id,",
        "    b.admission_date,",
        "    b.discharge_date,",
        "    strftime('%Y', b.admission_date) AS admission_year,",
        "    strftime('%Y-%m', b.admission_date) AS admission_month,",
        "    b.los_days,",
        "    b.admission_type,",
        "    p.patient_id,",
        "    p.patient_name,",
        "    p.age AS patient_age,",
        "    p.gender AS patient_gender,",
        "    d.dept_id,",
        "    d.dept_name,",
        "    pr.procedure_id,",
        "    pr.procedure_name,",
        "    pr.complexity_level,",
        "    f.facility_id,",
        "    f.facility_name,",
        "    f.city AS facility_city,",
        "    b.room_charges,",
        "    b.surgery_charges,",
        "    b.pharmacy_charges,",
        "    b.lab_charges,",
        "    b.total_billed_cost,",
        "    b.insurance_paid,",
        "    b.out_of_pocket_paid,",
        "    ROUND((b.out_of_pocket_paid / b.total_billed_cost) * 100, 2) AS out_of_pocket_pct,",
        "    ROUND((b.insurance_paid / b.total_billed_cost) * 100, 2) AS insurance_coverage_pct,",
        "    ROUND(b.total_billed_cost / b.los_days, 2) AS cost_per_day,",
        "    b.insurance_type,",
        "    b.is_outlier",
        "FROM fact_hospital_billing b",
        "JOIN dim_patients p ON b.patient_id = p.patient_id",
        "JOIN dim_procedures pr ON b.procedure_id = pr.procedure_id",
        "JOIN dim_departments d ON pr.dept_id = d.dept_id",
        "JOIN dim_facilities f ON b.facility_id = f.facility_id;"
    ]

    for idx, line in enumerate(ddl_statements, 3):
        cell = ws_ddl.cell(row=idx, column=1, value=line)
        cell.font = Font(name="Consolas", size=9.5, color="2C3E50" if not line.startswith("--") else "7F7F7F")
        ws_ddl.row_dimensions[idx].height = 16
    ws_ddl.column_dimensions["A"].width = 95

    target_path = "docs/08_Database_Design.xlsx"
    wb.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step08_db()
