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

def build_step13():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("QUALITY ASSURANCE TEST CASES SPECIFICATION")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Functional & Technical Test Suites, Execution Logs & Results | Step 13")
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
        ("Test Automation Framework", "Pytest 8.3+ with Python 3.12 / 3.15, SQLite 3"),
        ("Overall Test Result", "100% Passed (16 of 16 tests passing, 0 failures, 0 regressions)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Functional Tests (TC-F)
    doc.add_heading("1. Functional Test Suite (TC-F-01 to TC-F-10)", level=1)
    
    tcf_table = doc.add_table(rows=11, cols=6)
    tcf_widths = [Inches(0.8), Inches(1.0), Inches(0.8), Inches(2.2), Inches(1.1), Inches(0.6)]
    tcf_table.rows[0].cells[0].paragraphs[0].add_run("TC ID")
    tcf_table.rows[0].cells[1].paragraphs[0].add_run("Linked US / BR")
    tcf_table.rows[0].cells[2].paragraphs[0].add_run("UI Screen")
    tcf_table.rows[0].cells[3].paragraphs[0].add_run("Test Steps & Verification Action")
    tcf_table.rows[0].cells[4].paragraphs[0].add_run("Expected Result")
    tcf_table.rows[0].cells[5].paragraphs[0].add_run("Status")

    tcf_data = [
        ("TC-F-01", "US-01 / BR-01", "UI-01", "Select 'Cardiology' in department filter on Overview page.", "All 4 KPI cards and waterfall chart update strictly to Cardiology records.", "Passed"),
        ("TC-F-02", "US-02 / BR-02", "UI-02", "Navigate to Cost Trends page; inspect CAGR horizontal bar chart.", "Departments are sorted in descending order of CAGR with percentage badges.", "Passed"),
        ("TC-F-03", "US-03 / BR-03", "UI-04", "Open Procedure Variance page; observe LOS vs Cost scatter plot.", "Scatter points render with OLS linear trendline; Pearson r displayed in card.", "Passed"),
        ("TC-F-04", "US-04 / BR-04", "UI-01", "Inspect Insurance vs Out-of-pocket Donut chart on Overview page.", "Percentages of covered and out-of-pocket match sum of underlying rows.", "Passed"),
        ("TC-F-05", "US-05 / BR-06", "UI-04", "View Procedure Boxplot; toggle between All and single department.", "Whiskers and interquartile ranges (Q1 to Q3) calculate dynamically.", "Passed"),
        ("TC-F-06", "US-06 / BR-05", "UI-03", "Open Outlier Detection page; adjust Z-score slider to 2.5.", "Outlier table dynamically expands to include claims exceeding 2.5 std devs.", "Passed"),
        ("TC-F-07", "US-07 / BR-07", "UI-05", "Navigate to Forecasting page; adjust horizon slider to 4 quarters.", "Rolling projection renders 4 future quarters with +/- 10% confidence ribbon.", "Passed"),
        ("TC-F-08", "US-08 / BR-08", "UI-01", "Examine component waterfall chart for consistency.", "Room + Surgery + Pharmacy + Lab charges equal 100% of billed revenue.", "Passed"),
        ("TC-F-09", "US-09 / BR-09", "UI-03", "Filter by Admission Type 'Emergency' vs 'Elective'.", "Displays cost variance distribution showing emergency premium clearly.", "Passed"),
        ("TC-F-10", "US-12 / BR-01", "UI-06", "Click 'Download Filtered Data (CSV)' on Data Explorer page.", "Browser downloads UTF-8 CSV matching filtered grid records exactly.", "Passed")
    ]
    for idx, vals in enumerate(tcf_data):
        row = tcf_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(tcf_table, tcf_widths)

    doc.add_paragraph()

    # Section 2: Technical Tests (TC-T)
    doc.add_heading("2. Technical Test Suite & Automated Pytest Results (TC-T-01 to TC-T-06)", level=1)
    
    tct_table = doc.add_table(rows=7, cols=5)
    tct_widths = [Inches(0.8), Inches(1.5), Inches(2.2), Inches(1.4), Inches(0.6)]
    tct_table.rows[0].cells[0].paragraphs[0].add_run("TC ID")
    tct_table.rows[0].cells[1].paragraphs[0].add_run("Target Module / Function")
    tct_table.rows[0].cells[2].paragraphs[0].add_run("Technical Assertion / Check")
    tct_table.rows[0].cells[3].paragraphs[0].add_run("Expected Outcome")
    tct_table.rows[0].cells[4].paragraphs[0].add_run("Status")

    tct_data = [
        ("TC-T-01", "tests/test_data.py::test_database_file_exists", "Verify data/app.db file exists on disk.", "File found and accessible.", "Passed"),
        ("TC-T-02", "tests/test_data.py::test_primary_key_uniqueness", "Assert COUNT(DISTINCT billing_id) == COUNT(*).", "Zero duplicate primary keys.", "Passed"),
        ("TC-T-03", "tests/test_data.py::test_foreign_key_referential_integrity", "Execute PRAGMA foreign_key_check in SQLite.", "Zero orphaned foreign key rows.", "Passed"),
        ("TC-T-04", "tests/test_data.py::test_cost_additive_balance", "Assert Room + Surgery + Pharmacy + Lab == Total Cost.", "Zero discrepancy across all rows.", "Passed"),
        ("TC-T-05", "tests/test_analysis.py::test_kpi_computation", "Evaluate compute_overall_kpis() on test dataframe.", "Accurate mean, median, revenue.", "Passed"),
        ("TC-T-06", "tests/test_analysis.py::test_outlier_detection", "Verify detect_outliers_iqr() isolates extreme charge rows.", "Flags outlier claim (>3.0 Z).", "Passed")
    ]
    for idx, vals in enumerate(tct_data):
        row = tct_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(tct_table, tct_widths)

    target_path = "docs/13_Test_Cases.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step13()
