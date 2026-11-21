import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

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

def create_document():
    doc = docx.Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(40, 40, 40)
    
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("TREATMENT COST TREND ANALYSIS")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93) # Deep Navy
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Step 1: Project Understanding & Requirements Baseline Specification")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(90, 105, 120)
    
    doc.add_paragraph() # Spacer
    
    # Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_widths = [Inches(2.2), Inches(4.3)]
    meta_data = [
        ("Project Title", "Treatment Cost Trend Analysis (Interactive Data Analyst Analytics Dashboard)"),
        ("Student / Author", "Arohi Tyagi (Roll No: 2500320140033)"),
        ("Program & Institution", "MCA 2nd Year (Data Analyst Specialization), ABES Engineering College"),
        ("Academic Mentor", "Mr. Dinesh"),
        ("Delivery Milestone", "STEP 1 of 18 — Project Understanding Document"),
        ("Technology Baseline", "Python 3.11+, Streamlit, SQLite, Plotly Express, Pytest, Tailwind-styled UI")
    ]
    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        row.cells[0].paragraphs[0].add_run(label).bold = True
        row.cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")
    
    doc.add_paragraph()
    
    # Section 1: Executive Context
    h1 = doc.add_heading("1. Executive Context & Domain Background", level=1)
    p = doc.add_paragraph(
        "In modern healthcare management, medical inflation, technological intervention costs, and variability in length of "
        "stay (LOS) have triggered severe cost escalation across therapeutic departments. Hospital administrators, clinical directors, "
        "and health insurance claim managers frequently lack a unified, transparent analytical platform to track inpatient treatment cost trajectories, "
        "identify anomalous billing patterns, and project departmental expenditures. This project establishes an interactive, data-driven "
        "Data Analyst & Cost Analytics Platform engineered to ingest longitudinal hospitalization data, execute statistical trend and "
        "outlier analysis, and deliver actionable operational intelligence."
    )
    
    # Section 2: Problem Statement
    h2 = doc.add_heading("2. Problem Statement", level=1)
    p_prob = doc.add_paragraph()
    p_prob.paragraph_format.left_indent = Inches(0.25)
    p_prob.paragraph_format.right_indent = Inches(0.25)
    r_prob = p_prob.add_run(
        "\"Healthcare administrators, clinical heads, and insurance adjudicators lack an integrated, data-driven analytical tool "
        "to track treatment cost inflation, evaluate departmental expenditure variance, and detect anomalous billing patterns, leading "
        "to budget unpredictability, inefficient financial resource allocation, and elevated out-of-pocket burdens on patients.\""
    )
    r_prob.font.bold = True
    r_prob.font.size = Pt(11)
    r_prob.font.color.rgb = RGBColor(180, 40, 40)
    
    # Section 3: Intended Audience & User Personas
    doc.add_heading("3. Intended Audience & End-User Cadence", level=1)
    doc.add_paragraph(
        "The web application serves three primary tiers of operational and strategic decision-makers within hospital networks and Data Analyst units:"
    )
    
    aud_table = doc.add_table(rows=4, cols=4)
    aud_widths = [Inches(1.5), Inches(1.5), Inches(1.3), Inches(2.2)]
    aud_table.rows[0].cells[0].paragraphs[0].add_run("User Segment")
    aud_table.rows[0].cells[1].paragraphs[0].add_run("Target Persona")
    aud_table.rows[0].cells[2].paragraphs[0].add_run("Usage Cadence")
    aud_table.rows[0].cells[3].paragraphs[0].add_run("Primary Motivation & Workflow")
    
    aud_data = [
        ("Hospital Financial Controller", "Chief Financial Officer / Billing Manager", "Weekly & Monthly", "Audits departmental cost trends, monitors margin variance, and assesses budget deviations."),
        ("Clinical Operations Director", "Head of Clinical Services / HODs", "Bi-weekly / Monthly", "Evaluates cost-per-procedure vs. length of stay (LOS), protocol cost compliance, and resource usage."),
        ("Health Insurance & Claims Lead", "TPA Liaison / Claim Auditor", "Daily / Weekly", "Monitors claim acceptance ratios, identifies cost anomalies, and tracks out-of-pocket patient ratios.")
    ]
    for idx, row_vals in enumerate(aud_data):
        row = aud_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(aud_table, aud_widths)
    
    doc.add_paragraph()
    
    # Section 4: Stakeholders
    doc.add_heading("4. Stakeholder Analysis Matrix", level=1)
    doc.add_paragraph(
        "The following matrix maps the project stakeholders, categorizing their roles, vested interests, level of influence, "
        "and concrete decisions driven by the analytics platform."
    )
    
    sh_table = doc.add_table(rows=7, cols=6)
    sh_widths = [Inches(0.7), Inches(1.2), Inches(1.1), Inches(1.4), Inches(0.8), Inches(1.3)]
    sh_table.rows[0].cells[0].paragraphs[0].add_run("SH ID")
    sh_table.rows[0].cells[1].paragraphs[0].add_run("Stakeholder")
    sh_table.rows[0].cells[2].paragraphs[0].add_run("Role")
    sh_table.rows[0].cells[3].paragraphs[0].add_run("Interest")
    sh_table.rows[0].cells[4].paragraphs[0].add_run("Influence")
    sh_table.rows[0].cells[5].paragraphs[0].add_run("Decision Made via App")
    
    sh_data = [
        ("SH-01", "Hospital CFO / Finance Lead", "Executive Sponsor", "Cost containment, margin optimization, budget forecasting", "High", "Approve departmental budgets and set target billing caps."),
        ("SH-02", "Clinical Department Heads", "Operational User", "Treatment cost efficacy, standard care protocols vs. cost", "High", "Standardize clinical pathways to minimize unwarranted cost spikes."),
        ("SH-03", "Medical Billing & Claims Head", "Primary User", "Billing integrity, outlier detection, claim denials", "Medium", "Audit anomalous billing entries prior to insurer submission."),
        ("SH-04", "Health Insurance / TPA Analyst", "External Auditor", "Claim validity, actuarial cost benchmarking, pricing", "Medium", "Negotiate package rates and pre-authorization limits."),
        ("SH-05", "Data Analyst (Learner - Arohi)", "Lead Developer", "Model accuracy, interactive UX, statistical rigor", "High", "Select econometric models, visualization filters, and query flows."),
        ("SH-06", "Academic Mentor (Mr. Dinesh)", "Evaluator / Guide", "Technical completeness, academic rigor, ABES standards", "High", "Assess rubric adherence, test coverage, and documentation.")
    ]
    for idx, row_vals in enumerate(sh_data):
        row = sh_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(sh_table, sh_widths)
    
    doc.add_paragraph()
    
    # Section 5: Operational Questions (OQ)
    doc.add_heading("5. Operational Questions (OQ) & Data Feasibility Verification", level=1)
    doc.add_paragraph(
        "To deliver concrete business utility, the application must answer exactly ten fundamental operational questions. "
        "Every operational question has been verified against measurable data attributes in the underlying data model."
    )
    
    oq_table = doc.add_table(rows=11, cols=4)
    oq_widths = [Inches(0.7), Inches(2.3), Inches(2.0), Inches(1.5)]
    oq_table.rows[0].cells[0].paragraphs[0].add_run("OQ ID")
    oq_table.rows[0].cells[1].paragraphs[0].add_run("Operational Business Question")
    oq_table.rows[0].cells[2].paragraphs[0].add_run("Required Data Attributes")
    oq_table.rows[0].cells[3].paragraphs[0].add_run("Analytical Method")
    
    oq_data = [
        ("OQ-01", "What is the historical month-over-month trend of average and median treatment costs across medical departments?", "Admission_Date, Department, Total_Billed_Cost", "Time-series aggregation, MoM % growth calculation"),
        ("OQ-02", "Which clinical departments and diagnostic categories exhibit the steepest cost inflation over time?", "Department, Diagnosis_Group, Admission_Date, Total_Billed_Cost", "CAGR calculation, linear trend slope decomposition"),
        ("OQ-03", "How does Length of Stay (LOS) correlate with total inpatient costs across different age groups and genders?", "LOS_Days, Total_Billed_Cost, Patient_Age, Patient_Gender", "Pearson/Spearman correlation, scatter regression"),
        ("OQ-04", "What proportion of total treatment expenses is covered by health insurance versus out-of-pocket patient payment?", "Insurance_Paid, Out_Of_Pocket_Paid, Total_Billed_Cost, Insurance_Type", "Ratio decomposition, stacked area & donut distributions"),
        ("OQ-05", "Which specific medical procedures or surgeries have the highest inter-patient cost variability (IQR / variance)?", "Procedure_Name, Total_Billed_Cost, Department", "Box-whisker IQR, Standard Deviation / Variance analysis"),
        ("OQ-06", "Are there statistical cost outliers that indicate potential overbilling, clinical complications, or documentation errors?", "Total_Billed_Cost, Procedure_Name, LOS_Days", "Z-score (>3.0) and Tukey's Fences IQR outlier filters"),
        ("OQ-07", "What is the projected average treatment cost per department for the upcoming quarter based on historical trajectory?", "Total_Billed_Cost, Admission_Date, Department", "Moving average and Holt-Winters / polynomial trend projection"),
        ("OQ-08", "How do medication, room rent, surgical fees, and diagnostic test charges contribute to the overall billed amount?", "Room_Charges, Surgery_Charges, Pharmacy_Charges, Lab_Charges", "Component cost waterfall and percentage breakdown"),
        ("OQ-09", "How do patient admission types (Emergency vs. Elective vs. Urgent) impact total hospitalization expenditure?", "Admission_Type, Total_Billed_Cost, Department", "ANOVA / Kruskal-Wallis group variance comparison"),
        ("OQ-10", "What are the key financial performance indicators (Average Cost per Day, Median Cost per Stay) by hospital branch/facility?", "Facility_Location, LOS_Days, Total_Billed_Cost", "Daily unit cost computation, departmental benchmarks")
    ]
    for idx, row_vals in enumerate(oq_data):
        row = oq_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(oq_table, oq_widths)
    
    doc.add_paragraph()
    
    # Data Feasibility Check Paragraph
    p_check = doc.add_paragraph()
    r_chk_title = p_check.add_run("Data Feasibility Verification Check: ")
    r_chk_title.bold = True
    p_check.add_run(
        "PASSED. All ten operational questions (OQ-01 through OQ-10) directly map to concrete, derivable fields in the planned synthetic "
        "relational dataset. None of the questions require external proprietary telematics or qualitative unstructured records, ensuring 100% "
        "computational feasibility within the SQLite + Pandas analytics engine."
    )
    
    # Section 6: SMART Objectives
    doc.add_heading("6. Project Objectives (SMART)", level=1)
    doc.add_paragraph(
        "The project is structured around five specific, measurable, achievable, relevant, and time-bound objectives:"
    )
    
    obj_table = doc.add_table(rows=6, cols=3)
    obj_widths = [Inches(1.0), Inches(2.2), Inches(3.3)]
    obj_table.rows[0].cells[0].paragraphs[0].add_run("Objective ID")
    obj_table.rows[0].cells[1].paragraphs[0].add_run("SMART Objective Statement")
    obj_table.rows[0].cells[2].paragraphs[0].add_run("Verification Metric & Target")
    
    obj_data = [
        ("OBJ-01", "Design and populate a validated synthetic healthcare billing database.", "Generate >= 2,500 records across 6 departments (2023-2026) with zero null keys and 100% referential integrity."),
        ("OBJ-02", "Engineer automated cost trend and statistical distribution algorithms.", "Implement MoM trend analysis, IQR outlier detection, and cost component breakdown with < 200ms execution time."),
        ("OBJ-03", "Build a responsive, multi-page Streamlit web analytics dashboard.", "Deliver 5 interactive screens (Overview, Trends, Outliers, Forecasting, Data Explorer) with Tailwind CSS styling."),
        ("OBJ-04", "Enforce enterprise-grade software verification and test automation.", "Achieve >= 85% automated code coverage with >= 15 pytest unit and integration test cases passing."),
        ("OBJ-05", "Deploy to Streamlit Community Cloud and complete academic documentation.", "Achieve live public URL deployment with zero runtime crashes and 100% ABES academic report compliance.")
    ]
    for idx, row_vals in enumerate(obj_data):
        row = obj_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(obj_table, obj_widths)
    
    doc.add_paragraph()
    
    # Section 7: Project Scope
    doc.add_heading("7. Project Scope: In-Scope vs. Out-of-Scope", level=1)
    
    scope_table = doc.add_table(rows=6, cols=2)
    scope_widths = [Inches(3.25), Inches(3.25)]
    scope_table.rows[0].cells[0].paragraphs[0].add_run("In-Scope Functional Boundaries")
    scope_table.rows[0].cells[1].paragraphs[0].add_run("Out-of-Scope (Deferred / Excluded)")
    
    scope_data = [
        ("Synthetic multi-year hospitalization and billing dataset (SQLite storage).", "Live integration with proprietary hospital Electronic Health Records (EHR / Epic / Cerner)."),
        ("Exploratory Data Analysis (EDA) of treatment costs, LOS, and admission types.", "Predictive deep learning models for clinical diagnostics or medical imaging analysis."),
        ("Interactive time-series visualizations, trend filters, and procedural cost IQR boxplots.", "Real-time payment gateway processing or financial transaction settlement."),
        ("Statistical outlier detection for unusual billing and stay durations.", "Multi-tenant role-based user management with HIPAA/HL7 encrypted telecommunication layers."),
        ("Interactive forecasting simulation based on moving averages and historical trend vectors.", "Automated medical insurance pre-authorization underwriting bot.")
    ]
    for idx, row_vals in enumerate(scope_data):
        row = scope_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(scope_table, scope_widths)
    
    doc.add_paragraph()
    
    # Section 8: Assumptions & Constraints
    doc.add_heading("8. Assumptions & Constraints", level=1)
    
    doc.add_heading("8.1 System & Business Assumptions", level=2)
    assumptions = [
        ("A-01: Data Validity", "Synthetic billing records accurately mirror real-world hospital billing distributions, length-of-stay dynamics, and medical inflation rates (6-10% per annum)."),
        ("A-02: User Environment", "End-users access the dashboard via standard modern desktop web browsers (Chrome, Edge, Firefox) with active internet connectivity."),
        ("A-03: Currency Standardization", "All financial records and billing figures are standardized in Indian Rupees (INR) with uniform decimal precision."),
        ("A-04: Analytical Statelessness", "The analytics dashboard operates primarily in analytical read-only mode over the SQLite database with on-the-fly filtering and parameter simulation."),
        ("A-05: Regulatory Privacy", "Zero personally identifiable patient information (PII) is collected or processed, guaranteeing complete compliance with privacy standards.")
    ]
    for code, text in assumptions:
        p_item = doc.add_paragraph(style='List Bullet')
        r_c = p_item.add_run(f"{code}: ")
        r_c.bold = True
        p_item.add_run(text)
        
    doc.add_heading("8.2 Technical & Operational Constraints", level=2)
    constraints = [
        ("C-01: Deployment Platform", "Streamlit Community Cloud imposes 1 GB resource limits; data storage and in-memory analytics must remain lean and performant."),
        ("C-02: Zero-Cost Licensing", "All software components, libraries, and frameworks must be open-source (Python, Streamlit, Plotly, SQLite, Pytest)."),
        ("C-03: Single Academic Developer", "Architecture and implementation are designed for completion by a single MCA student (Arohi Tyagi) within the academic semester timeframe."),
        ("C-04: Local Database Scope", "SQLite is chosen as the lightweight serverless RDBMS to ensure seamless portability across local development and cloud hosting without requiring external cloud DB instances."),
        ("C-05: Institutional Compliance", "Documentation, project reports, and presentation materials must strictly adhere to ABES Engineering College guidelines and rubrics.")
    ]
    for code, text in constraints:
        p_item = doc.add_paragraph(style='List Bullet')
        r_c = p_item.add_run(f"{code}: ")
        r_c.bold = True
        p_item.add_run(text)
        
    doc.add_paragraph()
    
    # Section 9: Success Criteria
    doc.add_heading("9. Measurable Success Criteria", level=1)
    doc.add_paragraph(
        "The project will be deemed successful upon satisfying the following verifiable benchmark criteria:"
    )
    
    sc_table = doc.add_table(rows=6, cols=3)
    sc_widths = [Inches(1.0), Inches(2.3), Inches(3.2)]
    sc_table.rows[0].cells[0].paragraphs[0].add_run("Metric ID")
    sc_table.rows[0].cells[1].paragraphs[0].add_run("Success Dimension")
    sc_table.rows[0].cells[2].paragraphs[0].add_run("Benchmark Target & Verification Method")
    
    sc_data = [
        ("SC-01", "Analytical Coverage", "All 10 Operational Questions (OQ-01 to OQ-10) are answerable directly via UI filters and interactive Plotly visual charts."),
        ("SC-02", "System Performance", "Dashboard page rendering and query filtering response time under 1.5 seconds for a dataset of 2,500+ records."),
        ("SC-03", "Software Quality & Testing", "100% passing rate on automated test suite (Pytest) with minimum 85% code coverage across calculation and transformation logic."),
        ("SC-04", "Deployment & Accessibility", "100% availability on Streamlit Community Cloud with public HTTPS URL accessible to external evaluators without authentication hurdles."),
        ("SC-05", "Academic Evaluation", "Complete compliance with ABES project guidelines, fully cross-referenced traceability matrix, and defense readiness before mentor Mr. Dinesh.")
    ]
    for idx, row_vals in enumerate(sc_data):
        row = sc_table.rows[idx+1]
        for c_idx, val in enumerate(row_vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(sc_table, sc_widths)
    
    # Footer note
    doc.add_paragraph()
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_foot = p_foot.add_run("Document generated for MCA Project Work | ABES Engineering College | 2026")
    r_foot.font.size = Pt(8.5)
    r_foot.font.italic = True
    r_foot.font.color.rgb = RGBColor(128, 128, 128)
    
    # Save documents
    path1 = "docs/01_Project_Understanding.docx"
    path2 = "docs/01_Treatment_Cost_Trend_Analysis_Project_Understanding.docx"
    doc.save(path1)
    doc.save(path2)
    print(f"Successfully generated:\n- {path1}\n- {path2}")

if __name__ == "__main__":
    create_document()
