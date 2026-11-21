import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, col_widths, header_bg="1B365D", alt_bg="F2F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        is_header = (i == 0)
        for j, cell in enumerate(row.cells):
            if j < len(col_widths):
                cell.width = col_widths[j]
            set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
            if is_header:
                set_cell_background(cell, header_bg)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(9.5)
                        run.font.name = "Calibri"
            else:
                if i % 2 == 1:
                    set_cell_background(cell, "FFFFFF")
                else:
                    set_cell_background(cell, alt_bg)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9)
                        run.font.name = "Calibri"

def build_step05_brd():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("BUSINESS REQUIREMENTS DOCUMENT (BRD)")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Treatment Cost Trend Analysis Platform | Project ID: 033")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(90, 105, 120)

    doc.add_paragraph()

    # Section 1: Document Control
    doc.add_heading("1. Document Control & Revision History", level=1)
    dc_table = doc.add_table(rows=5, cols=4)
    dc_widths = [Inches(1.0), Inches(1.8), Inches(1.7), Inches(2.0)]
    dc_table.rows[0].cells[0].paragraphs[0].add_run("Version")
    dc_table.rows[0].cells[1].paragraphs[0].add_run("Author / Role")
    dc_table.rows[0].cells[2].paragraphs[0].add_run("Reviewer / Mentor")
    dc_table.rows[0].cells[3].paragraphs[0].add_run("Release Status / Remarks")

    dc_data = [
        ("v1.0", "Arohi Tyagi (Lead Analyst)", "Prof. Kalpana Dwivedi", "Initial Baseline Approval"),
        ("Project ID", "033 - Trend Cost Analysis", "Department", "MCA 2nd Yr (Data Analytics), ABES"),
        ("Date", "November 2026", "Environment", "Streamlit Cloud / SQLite"),
        ("Classification", "Academic Enterprise Capstone", "Traceability Status", "100% Traceable to HMW & OQ")
    ]
    for idx, vals in enumerate(dc_data):
        row = dc_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(dc_table, dc_widths, header_bg="2C3E50")

    doc.add_paragraph()

    # Section 2: Executive Summary
    doc.add_heading("2. Executive Summary", level=1)
    doc.add_paragraph(
        "Modern health systems face compounding inpatient cost pressures driven by specialized clinical technologies, varying lengths of stay, "
        "and opaque medical billing practices. This Business Requirements Document (BRD) specifies the functional, operational, and "
        "business standards for the Treatment Cost Trend Analysis platform. By transforming raw historical hospitalization and billing "
        "transactions into actionable time-series metrics, variance indicators, and anomaly flags, the solution equips healthcare leadership "
        "with the tools required to curb expenditure drift, audit rogue charges, and safeguard patient affordability."
    )

    # Section 3: Business Problem & Opportunity
    doc.add_heading("3. Business Problem & Opportunity", level=1)
    doc.add_paragraph(
        "Current hospital billing review relies on retrospective accounting ledger summaries, creating a 30-to-60-day lag between clinical discharge "
        "and financial variance detection. Consequently, departmental budget overruns persist unmitigated, and claim dispute rates remain elevated. "
        "The proposed platform establishes an interactive analytical system capable of computing monthly cost momentum, highlighting anomalous "
        "invoices via Tukey's Interquartile Range (IQR) filters, and breaking down bills into constituent surgical, pharmacy, diagnostic, and accommodation charges."
    )

    # Section 4: Objectives & Success KPIs
    doc.add_heading("4. Business Objectives & Key Performance Indicators (KPIs)", level=1)
    kpi_table = doc.add_table(rows=6, cols=5)
    kpi_widths = [Inches(1.5), Inches(1.5), Inches(1.3), Inches(1.0), Inches(1.2)]
    kpi_table.rows[0].cells[0].paragraphs[0].add_run("KPI Name")
    kpi_table.rows[0].cells[1].paragraphs[0].add_run("Numerator")
    kpi_table.rows[0].cells[2].paragraphs[0].add_run("Denominator")
    kpi_table.rows[0].cells[3].paragraphs[0].add_run("Period")
    kpi_table.rows[0].cells[4].paragraphs[0].add_run("Target Value")

    kpi_data = [
        ("Average Treatment Cost", "Sum of Total Billed Cost", "Count of Patient Admissions", "Monthly", "< ₹1,25,000 / stay"),
        ("Outlier Detection Ratio", "Flagged High-Cost Outliers", "Total Hospital Invoices", "Monthly", "1.5% to 3.5% of cohort"),
        ("Out-of-Pocket Payment %", "Sum of Patient Out-of-Pocket Paid", "Total Billed Healthcare Cost", "Quarterly", "< 28.0% of total"),
        ("Departmental CAGR", "Current Period Dept Mean Cost", "Base Period Dept Mean Cost", "Annual", "Within 6% - 8% p.a."),
        ("Average Length of Stay (ALOS)", "Sum of Inpatient Bed Days", "Total Patient Discharges", "Monthly", "3.8 to 5.2 Days")
    ]
    for idx, vals in enumerate(kpi_data):
        row = kpi_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(kpi_table, kpi_widths)

    doc.add_paragraph()

    # Section 5: Stakeholder RACI Matrix
    doc.add_heading("5. Stakeholder RACI Matrix", level=1)
    raci_table = doc.add_table(rows=7, cols=6)
    raci_widths = [Inches(1.8), Inches(0.9), Inches(0.9), Inches(0.9), Inches(1.0), Inches(1.0)]
    raci_table.rows[0].cells[0].paragraphs[0].add_run("Project Lifecycle Phase")
    raci_table.rows[0].cells[1].paragraphs[0].add_run("SH-01 CFO")
    raci_table.rows[0].cells[2].paragraphs[0].add_run("SH-02 Clinical")
    raci_table.rows[0].cells[3].paragraphs[0].add_run("SH-03 Billing")
    raci_table.rows[0].cells[4].paragraphs[0].add_run("SH-04 Insurer")
    raci_table.rows[0].cells[5].paragraphs[0].add_run("SH-05 Analyst")

    raci_data = [
        ("Requirements & Scope Definition", "Accountable", "Consulted", "Informed", "Informed", "Responsible"),
        ("Database & Schema Modeling", "Informed", "Consulted", "Consulted", "Informed", "Responsible"),
        ("Analytical Engine & Algorithm Dev", "Consulted", "Consulted", "Accountable", "Informed", "Responsible"),
        ("Dashboard & Visualization Testing", "Accountable", "Responsible", "Responsible", "Consulted", "Responsible"),
        ("Outlier Audit & Model Verification", "Informed", "Consulted", "Accountable", "Consulted", "Responsible"),
        ("Cloud Deployment & Sign-off", "Accountable", "Informed", "Informed", "Informed", "Responsible")
    ]
    for idx, vals in enumerate(raci_data):
        row = raci_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(raci_table, raci_widths)

    doc.add_paragraph()

    # Section 6: Scope
    doc.add_heading("6. Project Scope: In-Scope vs. Out-of-Scope", level=1)
    doc.add_paragraph("In-Scope Functional Boundaries:")
    in_scope = [
        "Synthetic longitudinal inpatient billing dataset covering 2023–2026 across 6 clinical departments.",
        "Interactive time-series visualizations of monthly mean, median, and cumulative treatment expenditures.",
        "Automated statistical outlier identification using Tukey's Fences (1.5 * IQR) and standardized Z-scores.",
        "Cost component waterfall analysis dissecting room charges, surgical fees, pharmacy costs, and laboratory tests.",
        "Rolling 90-day predictive trend projection for departmental financial planning."
    ]
    for pt in in_scope:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(pt)

    doc.add_paragraph("Out-of-Scope Boundaries:")
    out_scope = [
        "Live bidirectional HL7/FHIR messaging synchronization with external proprietary EHR hospital databases.",
        "Clinical computer vision or AI automated disease diagnostics.",
        "Real-time banking payment gateway settlement for patient out-of-pocket transactions."
    ]
    for pt in out_scope:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(pt)

    doc.add_paragraph()

    # Section 7: Detailed Business Requirements
    doc.add_heading("7. Detailed Business Requirements & Acceptance Criteria", level=1)
    
    br_details = [
        ("BR-01", "MoM Cost Trend Aggregation", "Must", "HMW-01", 
         "The business needs to aggregate and track month-over-month inpatient billing totals, means, and medians so that leadership can spot systemic financial drift.",
         "Given a multi-year billing dataset, when filtered by department and year, then the platform must compute and display monthly Mean Cost, Median Cost, and MoM % Change with < 100ms calculation latency."),
        ("BR-02", "Departmental Cost Inflation", "Must", "HMW-02",
         "The business needs to compute annual compound inflation rates by clinical department so that high-inflation specialties can be identified.",
         "Given chronological admission records, when a user selects a department, then the system must display the annualized CAGR and cost trajectory with statistical trend fit line."),
        ("BR-03", "Length of Stay (LOS) vs Cost Correlation", "Must", "HMW-03",
         "The business needs to evaluate the relationship between Length of Stay (LOS) and total billed expenditure across patient demographics so that clinical bed turnover is optimized.",
         "Given patient discharge records, the platform must compute Pearson correlation between LOS days and total cost, segmented by gender and age cohorts."),
        ("BR-04", "Insurance Reimbursement vs Out-of-Pocket", "Must", "HMW-04",
         "The business needs to quantify the financial split between insurance reimbursements and patient out-of-pocket costs so that affordability and insurance adequacy can be assessed.",
         "The platform must verify that Insurance_Paid + Out_Of_Pocket_Paid == Total_Billed_Cost for 100% of records, displaying interactive donut and stacked proportion distributions."),
        ("BR-05", "Automated Billing Outlier Detection", "Must", "HMW-06",
         "The business needs to automatically identify statistical billing outliers beyond normal distribution thresholds so that suspicious claims can be audited before billing finalization.",
         "The platform must flag transactions where Total_Billed_Cost > Q3 + 1.5 * IQR or Z-score > 3.0, presenting flagged claims in a dedicated review table with export capability."),
        ("BR-06", "Inter-Procedure Cost Variance", "Must", "HMW-05",
         "The business needs to analyze inter-patient cost variability (IQR and standard deviation) for standard surgical procedures so that clinical protocol compliance can be enforced.",
         "The system must render box-whisker IQR distribution plots for each procedure with median markers, lower fence, and upper fence."),
        ("BR-07", "Rolling 90-Day Cost Projection", "Should", "HMW-07",
         "The business needs to generate short-term rolling 90-day cost projections by clinical specialty so that finance teams can anticipate quarterly working capital needs.",
         "The forecasting module must extrapolate historical quarterly trends with confidence intervals (+/- 10%) using polynomial or moving average smoothing."),
        ("BR-08", "Atomic Cost Component Waterfall", "Must", "HMW-08",
         "The business needs to decompose patient bills into distinct cost categories (room, surgery, medicine, lab) so that the principal operational cost drivers are visible.",
         "The UI must render a waterfall chart showing the additive contribution of Room, Surgery, Pharmacy, and Diagnostics summing to 100% of total revenue."),
        ("BR-09", "Admission Type Expenditure Disparity", "Should", "HMW-09",
         "The business needs to compare expenditure patterns across admission types (Emergency vs. Elective vs. Urgent) so that emergency care surge pricing can be monitored.",
         "The platform must compute group statistical variance (ANOVA) and render comparative distribution violins across admission types."),
        ("BR-10", "Multi-Facility Daily Cost Benchmark", "Should", "HMW-10",
         "The business needs to benchmark cost performance and average daily charges across multiple hospital branches so that operational cost disparities are normalized.",
         "The system must compare Average Daily Cost (Total Cost / LOS) across branch facilities with benchmark deviation indicators."),
        ("BR-11", "Scenario Fee Cap Simulation", "Could", "HMW-01",
         "The business needs to simulate the financial impact of potential price cap adjustments on procedure revenues so that executive scenario planning is enabled.",
         "The user can adjust a slider (-20% to +20%) to simulate total projected hospital revenue adjustments dynamically."),
        ("BR-12", "Real-Time EHR API Integration", "Won't", "HMW-06",
         "The business needs automated real-time bi-directional synchronization with external hospital Electronic Health Record (EHR) systems.",
         "Deferred to enterprise phase 2; not required for standalone offline analytics capstone.")
    ]

    for br_id, title, priority, source, desc, ac in br_details:
        p_h = doc.add_paragraph()
        r_id = p_h.add_run(f"{br_id}: {title} ")
        r_id.bold = True
        r_id.font.color.rgb = RGBColor(27, 54, 93)
        r_pr = p_h.add_run(f"[{priority} | Source: {source}]")
        r_pr.italic = True
        
        p_d = doc.add_paragraph()
        p_d.add_run("Requirement: ").bold = True
        p_d.add_run(desc)
        
        p_ac = doc.add_paragraph()
        p_ac.add_run("Acceptance Criteria: ").bold = True
        p_ac.add_run(ac)

    doc.add_paragraph()

    # Section 8: Business Rules
    doc.add_heading("8. Business Rules Specification", level=1)
    b_rules = [
        ("BRULE-01: Cost Balance Consistency", "For every inpatient bill, Total_Billed_Cost MUST EQUAL (Room_Charges + Surgery_Charges + Pharmacy_Charges + Lab_Charges)."),
        ("BRULE-02: Settlement Parity", "For every billed record, Total_Billed_Cost MUST EQUAL (Insurance_Paid + Out_Of_Pocket_Paid). Out-of-pocket cannot exceed total cost."),
        ("BRULE-03: Length of Stay (LOS) Integrity", "LOS_Days = max(1, Discharge_Date - Admission_Date). Same-day discharges are recorded as 1 day."),
        ("BRULE-04: Non-Negativity", "All monetary attributes (charges, insurance paid, out-of-pocket) and patient age must be strictly non-negative (> 0)."),
        ("BRULE-05: Department Procedure Association", "Every medical procedure must belong strictly to one valid clinical department master entity."),
        ("BRULE-06: Outlier Definition", "A transaction is classified as an outlier if Total_Billed_Cost > Q3 + 1.5 * IQR or Z-score > 3.0 within its procedural category.")
    ]
    for r_code, r_text in b_rules:
        p_r = doc.add_paragraph(style='List Bullet')
        r_c = p_r.add_run(f"{r_code}: ")
        r_c.bold = True
        p_r.add_run(r_text)

    doc.add_paragraph()

    # Section 9: Data Requirements
    doc.add_heading("9. Data Ingestion & Refresh Requirements", level=1)
    doc.add_paragraph(
        "The application reads from a structured SQLite database (`data/app.db`) containing relational tables: "
        "`dim_departments`, `dim_procedures`, `dim_patients`, `dim_facilities`, and `fact_hospital_billing`. "
        "Analytical data is loaded into memory via `st.cache_data` to ensure rapid query execution and seamless interactive filtering."
    )

    # Section 10: Assumptions, Constraints & Dependencies
    doc.add_heading("10. Assumptions, Constraints & Technical Dependencies", level=1)
    doc.add_paragraph(
        "Assumptions: Synthetic dataset reflects real-world Indian private and corporate healthcare billing distributions.\n"
        "Constraints: 1 GB RAM memory cap on Streamlit Community Cloud; zero-cost open-source technology stack.\n"
        "Dependencies: Python 3.11+, Streamlit, Pandas, Plotly Express, NumPy, SciPy, Pytest."
    )

    # Section 11: Risks & Mitigation
    doc.add_heading("11. Risks & Mitigation Strategies", level=1)
    risk_table = doc.add_table(rows=5, cols=4)
    risk_widths = [Inches(1.2), Inches(1.8), Inches(1.6), Inches(1.9)]
    risk_table.rows[0].cells[0].paragraphs[0].add_run("Risk ID")
    risk_table.rows[0].cells[1].paragraphs[0].add_run("Risk Description")
    risk_table.rows[0].cells[2].paragraphs[0].add_run("Impact & Likelihood")
    risk_table.rows[0].cells[3].paragraphs[0].add_run("Mitigation Strategy")

    risk_data = [
        ("RISK-01", "Synthetic data lacks realistic statistical covariance.", "High / Medium", "Incorporate parameterized distributions (Log-normal costs, Poisson LOS) with medical inflation drift."),
        ("RISK-02", "Streamlit Cloud timeout during large data aggregation.", "Medium / Low", "Implement `@st.cache_data` caching and optimize SQLite indexes on date and department columns."),
        ("RISK-03", "UI rendering inconsistencies across mobile/desktop browsers.", "Low / Medium", "Implement responsive Tailwind card containers and standard Plotly autosize layouts."),
        ("RISK-04", "Academic documentation divergence from ABES rubric.", "High / Low", "Strictly enforce ABES formatting standards, Georgia/Calibri typography, and dxa table widths.")
    ]
    for idx, vals in enumerate(risk_data):
        row = risk_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(risk_table, risk_widths)

    doc.add_paragraph()

    # Section 12: Glossary
    doc.add_heading("12. Glossary of Terms", level=1)
    glossary = [
        ("ALOS", "Average Length of Stay: the average number of days an inpatient remains hospitalized."),
        ("CAGR", "Compound Annual Growth Rate: annualized rate of healthcare expenditure inflation."),
        ("IQR", "Interquartile Range: difference between 75th percentile (Q3) and 25th percentile (Q1)."),
        ("Out-of-Pocket", "The share of healthcare costs paid directly by the patient without insurance coverage."),
        ("TPA", "Third-Party Administrator: licensed organization that processes insurance claims.")
    ]
    for term, defn in glossary:
        p_g = doc.add_paragraph(style='List Bullet')
        p_g.add_run(f"{term}: ").bold = True
        p_g.add_run(defn)

    doc.add_paragraph()

    # Section 13: Sign-Off
    doc.add_heading("13. Stakeholder Approval & Sign-off", level=1)
    sign_table = doc.add_table(rows=4, cols=4)
    sign_widths = [Inches(1.8), Inches(1.8), Inches(1.4), Inches(1.5)]
    sign_table.rows[0].cells[0].paragraphs[0].add_run("Stakeholder Name")
    sign_table.rows[0].cells[1].paragraphs[0].add_run("Designation / Role")
    sign_table.rows[0].cells[2].paragraphs[0].add_run("Approval Status")
    sign_table.rows[0].cells[3].paragraphs[0].add_run("Signature Date")

    sign_data = [
        ("Arohi Tyagi", "Lead Student Analyst (Roll: 2500320140033)", "Approved", "20/11/2026"),
        ("Prof. Kalpana Dwivedi", "Academic Mentor & Evaluator, ABES", "Approved", "20/11/2026"),
        ("Hospital Advisory Panel", "Academic Healthcare Subject Proxy", "Approved", "20/11/2026")
    ]
    for idx, vals in enumerate(sign_data):
        row = sign_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(sign_table, sign_widths)

    target_path = "docs/05_BRD.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step05_brd()
