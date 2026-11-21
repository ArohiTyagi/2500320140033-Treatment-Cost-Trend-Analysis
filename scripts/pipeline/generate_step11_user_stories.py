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

def build_step11():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("AGILE USER STORIES CATALOG & SPRINT PLAN")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("User Story Specifications, Acceptance Criteria & Epics | Step 11")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(90, 105, 120)

    doc.add_paragraph()

    # Meta
    meta_table = doc.add_table(rows=5, cols=2)
    meta_widths = [Inches(2.2), Inches(4.3)]
    meta_data = [
        ("Project Title", "Treatment Cost Trend Analysis"),
        ("Student / Roll No", "Arohi Tyagi | Roll No: 2500320140033 | MCA 2nd Yr DA, ABES"),
        ("Academic Mentor", "Prof. Kalpana Dwivedi"),
        ("Agile Framework", "Scrum Framework (2 Sprints across 5 Epics, 13 User Stories, 42 Story Points)"),
        ("Acceptance Standard", "Gherkin Syntax (Given / When / Then) with 100% OQ Traceability")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: User Stories Table
    doc.add_heading("1. Agile User Stories Matrix", level=1)
    doc.add_paragraph(
        "Thirteen granular, actionable user stories were defined across five core epics. Each story represents a deliverable unit "
        "executable within half a development day."
    )

    us_table = doc.add_table(rows=14, cols=7)
    us_widths = [Inches(0.7), Inches(1.1), Inches(2.3), Inches(0.8), Inches(0.6), Inches(0.6), Inches(0.4)]
    us_table.rows[0].cells[0].paragraphs[0].add_run("US ID")
    us_table.rows[0].cells[1].paragraphs[0].add_run("Persona")
    us_table.rows[0].cells[2].paragraphs[0].add_run("User Story (As a... I want... So that...)")
    us_table.rows[0].cells[3].paragraphs[0].add_run("Feature ID")
    us_table.rows[0].cells[4].paragraphs[0].add_run("BR ID")
    us_table.rows[0].cells[5].paragraphs[0].add_run("Priority")
    us_table.rows[0].cells[6].paragraphs[0].add_run("Pts")

    us_data = [
        ("US-01", "Hospital CFO", "As a CFO, I want to view monthly average and median treatment cost trajectories so that I can detect budget overruns early.", "FEAT-01", "BR-01", "Must", "3"),
        ("US-02", "Clinical Head", "As a Clinical Head, I want to calculate annual cost inflation (CAGR) by department so that I can target high-inflation specialties.", "FEAT-02", "BR-02", "Must", "5"),
        ("US-03", "Medical Director", "As a Medical Director, I want to analyze how Length of Stay correlates with cost across patient demographics so that I can optimize bed utilization.", "FEAT-03", "BR-03", "Must", "3"),
        ("US-04", "TPA Claims Lead", "As a Claims Lead, I want to view the financial split between insurance reimbursements and out-of-pocket costs so that patient burden is clear.", "FEAT-04", "BR-04", "Must", "3"),
        ("US-05", "Clinical Head", "As a Clinical Head, I want to evaluate interquartile cost variance for surgical procedures so that clinical protocol compliance can be enforced.", "FEAT-05", "BR-06", "Must", "5"),
        ("US-06", "Billing Auditor", "As a Billing Auditor, I want an automated outlier detection tool to flag anomalous high-cost claims so that billing errors are resolved before submission.", "FEAT-06", "BR-05", "Must", "5"),
        ("US-07", "Hospital CFO", "As a CFO, I want to view a 90-day rolling treatment expenditure forecast so that upcoming working capital needs can be anticipated.", "FEAT-07", "BR-07", "Should", "3"),
        ("US-08", "Finance Officer", "As a Finance Officer, I want to view a waterfall decomposition of charges (room, surgery, lab, pharmacy) so that primary cost drivers are visible.", "FEAT-08", "BR-08", "Must", "3"),
        ("US-09", "Operations Lead", "As an Operations Lead, I want to compare expenditures between Emergency and Elective admissions so that acute care pricing is understood.", "FEAT-09", "BR-09", "Should", "3"),
        ("US-10", "Hospital CFO", "As a CFO, I want to benchmark daily bed charges across hospital branches so that regional cost variations can be normalized.", "FEAT-10", "BR-10", "Should", "3"),
        ("US-11", "Finance Analyst", "As a Finance Analyst, I want to simulate procedure price adjustments via interactive sliders so that scenario planning is supported.", "FEAT-11", "BR-11", "Could", "2"),
        ("US-12", "Data Auditor", "As a Data Auditor, I want to inspect raw relational billing records and export filtered tables to CSV so that external audits can be conducted.", "FEAT-13", "BR-01", "Must", "2"),
        ("US-13", "Lead Developer", "As a Developer, I want automated unit tests for all statistical calculations so that data integrity and regression safety are guaranteed.", "FEAT-14", "BR-01", "Must", "2")
    ]
    for idx, vals in enumerate(us_data):
        row = us_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(us_table, us_widths)

    doc.add_paragraph()

    # Section 2: Acceptance Criteria
    doc.add_heading("2. Acceptance Criteria (Given / When / Then)", level=1)
    
    crit_list = [
        ("US-01: Cost Trends Tracking", 
         "Given the app is launched and filtered to 'Cardiology', When the Monthly Trend page is selected, Then an interactive line chart shows mean and median monthly cost from 2023 to 2026 with a MoM % change table."),
        ("US-02: Departmental Inflation (CAGR)",
         "Given multi-year hospitalization data, When a user selects a department from the sidebar, Then the system displays the computed CAGR percentage and renders a trend fit line."),
        ("US-03: LOS vs Cost Correlation",
         "Given the Procedure Variance page is open, When the user inspects the scatter plot, Then each point represents one patient stay with regression slope, and Pearson r value displayed in a metric card."),
        ("US-04: Insurance Reimbursement vs Out-of-Pocket",
         "Given patient billing records, When viewing the financial summary, Then a Plotly donut chart displays the percentage split of Insurance Covered vs Out-of-Pocket Paid totaling 100%."),
        ("US-05: Procedure IQR Variance",
         "Given clinical procedures within a department, When the procedure box plot is rendered, Then the 25th, 50th, and 75th percentiles are visible with upper and lower Tukey whiskers."),
        ("US-06: Statistical Outlier Flagging",
         "Given unfiltered hospital invoices, When the Outlier page is opened, Then transactions with Z-score > 3.0 or exceeding Q3 + 1.5 * IQR are displayed in a highlighted table with patient ID and cost.")
    ]
    for us_t, gwt in crit_list:
        p_c = doc.add_paragraph()
        r_t = p_c.add_run(f"{us_t}\n")
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(27, 54, 93)
        p_c.add_run(f"Acceptance Criteria: {gwt}")

    doc.add_paragraph()

    # Section 3: Epics & Sprint Plan
    doc.add_heading("3. Epic Grouping & Sprint Delivery Roadmap", level=1)
    
    sp_table = doc.add_table(rows=6, cols=5)
    sp_widths = [Inches(1.0), Inches(1.2), Inches(1.5), Inches(1.4), Inches(1.4)]
    sp_table.rows[0].cells[0].paragraphs[0].add_run("Sprint")
    sp_table.rows[0].cells[1].paragraphs[0].add_run("Epic Name")
    sp_table.rows[0].cells[2].paragraphs[0].add_run("Included User Stories")
    sp_table.rows[0].cells[3].paragraphs[0].add_run("Total Story Points")
    sp_table.rows[0].cells[4].paragraphs[0].add_run("Key Deliverable Outcome")

    sp_data = [
        ("Sprint 1 (Days 1-3)", "Epic 1: Data Engine", "US-12, US-13", "4 Pts", "SQLite DB, DDL schema, reproducible generator & data tests"),
        ("Sprint 1 (Days 3-5)", "Epic 2: Core Dashboard", "US-01, US-08, US-10", "9 Pts", "app.py, overview KPIs, Tailwind UI header, component waterfall"),
        ("Sprint 2 (Days 6-8)", "Epic 3: Trend & Inflation", "US-02, US-04", "8 Pts", "MoM trendline, CAGR algorithms, insurance reimbursement donut"),
        ("Sprint 2 (Days 8-10)", "Epic 4: Variance & Outliers", "US-03, US-05, US-06", "13 Pts", "LOS scatter correlation, procedure IQR boxplots, outlier detector"),
        ("Sprint 2 (Days 11-12)", "Epic 5: Forecasting & Insights", "US-07, US-09, US-11", "8 Pts", "Rolling 90-day forecast, admission comparison, price simulation")
    ]
    for idx, vals in enumerate(sp_data):
        row = sp_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(sp_table, sp_widths)

    target_path = "docs/11_User_Stories.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step11()
