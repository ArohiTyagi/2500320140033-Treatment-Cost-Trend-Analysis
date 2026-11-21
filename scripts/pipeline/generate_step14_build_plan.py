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

def build_step14():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("SYSTEM BUILD PLAN & IMPLEMENTATION ROADMAP")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Step-by-Step Task Breakdown, DoD, & Viva Notes | Step 14")
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
        ("Execution Roadmap", "10 Modular Build Tasks (BUILD-01 to BUILD-10) with Verification Commands"),
        ("Quality Threshold", "100% Passing Automated Tests & Live Cloud URL")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Summary
    doc.add_heading("1. Executive Project Summary", level=1)
    doc.add_paragraph(
        "The Treatment Cost Trend Analysis platform is an enterprise-grade Data Analytics & Inpatient Billing Analytics dashboard. "
        "Engineered for Arohi Tyagi (Roll No: 2500320140033) in the MCA 2nd Year Data Analytics program at ABES Engineering College "
        "under mentor Prof. Kalpana Dwivedi. The solution addresses medical inflation unpredictability, unexplained departmental cost drift, "
        "and billing anomalies across hospital networks. The system connects an embedded SQLite database (data/app.db) to a modular "
        "Python analytical engine and renders interactive Plotly visualizations within a Tailwind-styled Streamlit multi-page interface. "
        "The platform tracks month-over-month expenditure trends, calculates departmental CAGR, correlates Length of Stay (LOS) with "
        "patient charges, detects statistical outliers via Tukey's IQR fences and Z-scores, and forecasts rolling 90-day expenditures."
    )

    # Section 2: Build Tasks Table
    doc.add_heading("2. Granular Build Tasks (BUILD-01 to BUILD-10)", level=1)
    
    b_table = doc.add_table(rows=11, cols=4)
    b_widths = [Inches(1.0), Inches(2.3), Inches(1.8), Inches(1.4)]
    b_table.rows[0].cells[0].paragraphs[0].add_run("Task ID")
    b_table.rows[0].cells[1].paragraphs[0].add_run("Implementation Objective")
    b_table.rows[0].cells[2].paragraphs[0].add_run("Target Code Files")
    b_table.rows[0].cells[3].paragraphs[0].add_run("Verification Gate")

    b_data = [
        ("BUILD-01", "Configure environment & pinned dependencies", "requirements.txt, .streamlit/config.toml", "pip install -r requirements.txt"),
        ("BUILD-02", "Establish SQLite database & data access layer", "src/db.py, data/app.db", "pytest tests/test_data.py"),
        ("BUILD-03", "Build data cleaning & validation pipeline", "src/clean.py", "pytest tests/test_analysis.py"),
        ("BUILD-04", "Develop statistical analysis & KPI engine", "src/analysis.py", "pytest tests/test_analysis.py"),
        ("BUILD-05", "Implement reusable Tailwind UI components", "src/ui.py", "Component iframe smoke test"),
        ("BUILD-06", "Construct Executive Overview page (UI-01)", "app.py, pages/1_Overview.py", "Streamlit UI render"),
        ("BUILD-07", "Build Cost Trends & Inflation page (UI-02)", "pages/2_Cost_Trends.py", "MoM & CAGR visual check"),
        ("BUILD-08", "Build Outlier Detection & Audit page (UI-03)", "pages/3_Outlier_Detection.py", "Z-score slider & CSV export"),
        ("BUILD-09", "Build Procedure & LOS Variance page (UI-04)", "pages/4_Procedure_Analysis.py", "Scatter & IQR boxplots"),
        ("BUILD-10", "Build Forecasting & Explorer pages (UI-05/06)", "pages/5_Forecasting.py, pages/6_Data_Explorer.py", "16/16 tests passing")
    ]
    for idx, vals in enumerate(b_data):
        row = b_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(b_table, b_widths)

    doc.add_paragraph()

    # Section 3: Definition of Done
    doc.add_heading("3. Definition of Done (DoD)", level=1)
    dods = [
        "All Must-Have Business Requirements (BR-01 through BR-08) verified and operating with zero runtime errors.",
        "100% of automated unit and functional test cases in pytest passing cleanly.",
        "Application running smoothly in local environment and deployed to Streamlit Community Cloud with public HTTPS URL.",
        "All 18 project lifecycle deliverables compiled, formatted, and archived in the docs/ directory.",
        "Defense slide deck (11 slides) prepared strictly compliant with ABES Engineering College rubrics."
    ]
    for d in dods:
        p_d = doc.add_paragraph(style='List Bullet')
        p_d.add_run(d)

    target_path = "docs/14_BUILD_PLAN.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step14()
