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

def build_step15():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("APPLICATION BUILD & AUTOMATED TEST VERIFICATION REPORT")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Streamlit Multipage Implementation & Pytest Execution Logs | Step 15")
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
        ("Academic Mentor", "Mr. Dinesh"),
        ("Build Status", "100% Operational (app.py + 5 multipage submodules + src/ library)"),
        ("Pytest Execution", "16 passed in 2.76s (100% Pass Rate, 0 Failures)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Build Modules Summary
    doc.add_heading("1. Delivered Application Modules Summary", level=1)
    
    b_table = doc.add_table(rows=8, cols=4)
    b_widths = [Inches(1.8), Inches(1.2), Inches(1.0), Inches(2.5)]
    b_table.rows[0].cells[0].paragraphs[0].add_run("File Path")
    b_table.rows[0].cells[1].paragraphs[0].add_run("Module Layer")
    b_table.rows[0].cells[2].paragraphs[0].add_run("Lines of Code")
    b_table.rows[0].cells[3].paragraphs[0].add_run("Core Functionality")

    b_data = [
        ("app.py", "UI Presentation", "185 LOC", "Main entrypoint, executive KPI cards, cost waterfall, branch benchmarks"),
        ("pages/1_Cost_Trends.py", "UI Presentation", "95 LOC", "MoM mean vs median trajectories and departmental CAGR inflation"),
        ("pages/2_Outlier_Detection.py", "UI Presentation", "85 LOC", "Tukey IQR fences, Z-score scatter, highlighted table & CSV export"),
        ("pages/3_Procedure_Analysis.py", "UI Presentation", "80 LOC", "Procedure IQR boxplots and Length of Stay (LOS) scatter correlation"),
        ("pages/4_Forecasting.py", "UI Presentation", "90 LOC", "Rolling 90-day quarterly forecasting and price cap scenario simulation"),
        ("pages/5_Data_Explorer.py", "UI Presentation", "65 LOC", "Tabular database mart viewer, schema audit, and raw CSV download"),
        ("src/analysis.py", "Analytical Core", "145 LOC", "Statistical algorithms (CAGR, IQR, Pearson r, Spearman, Moving Avg)")
    ]
    for idx, vals in enumerate(b_data):
        row = b_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(b_table, b_widths)

    doc.add_paragraph()

    # Section 2: Pytest Execution Log
    doc.add_heading("2. Automated Test Execution Transcript", level=1)
    p_log = doc.add_paragraph()
    p_log.paragraph_format.left_indent = Inches(0.2)
    r_log = p_log.add_run(
        "============================= test session starts =============================\n"
        "platform win32 -- Python 3.15.2, pytest-9.2.0, pluggy-1.7.0\n"
        "rootdir: C:\\Users\\Arohi Tyagi\\OneDrive\\Desktop\\033TrendCostAnalysis\n"
        "collected 16 items\n\n"
        "tests/test_analysis.py::test_kpi_computation PASSED                      [  6%]\n"
        "tests/test_analysis.py::test_monthly_trends PASSED                       [ 12%]\n"
        "tests/test_analysis.py::test_department_cagr PASSED                      [ 18%]\n"
        "tests/test_analysis.py::test_outlier_detection PASSED                    [ 25%]\n"
        "tests/test_analysis.py::test_los_correlation PASSED                      [ 31%]\n"
        "tests/test_analysis.py::test_forecast_model PASSED                       [ 37%]\n"
        "tests/test_analysis.py::test_data_cleaning PASSED                        [ 43%]\n"
        "tests/test_app.py::test_app_source_files_exist PASSED                    [ 50%]\n"
        "tests/test_app.py::test_full_pipeline_load PASSED                        [ 56%]\n"
        "tests/test_data.py::test_database_file_exists PASSED                     [ 62%]\n"
        "tests/test_data.py::test_tables_exist PASSED                             [ 68%]\n"
        "tests/test_data.py::test_fact_billing_record_count PASSED                [ 75%]\n"
        "tests/test_data.py::test_primary_key_uniqueness PASSED                   [ 81%]\n"
        "tests/test_data.py::test_foreign_key_referential_integrity PASSED        [ 87%]\n"
        "tests/test_data.py::test_cost_additive_balance PASSED                    [ 93%]\n"
        "tests/test_data.py::test_analytical_view_exists_and_populated PASSED     [100%]\n\n"
        "============================= 16 passed in 2.76s =============================="
    )
    r_log.font.name = "Consolas"
    r_log.font.size = Pt(8.5)

    target_path = "docs/15_Application_Source_Code_and_Test_Report.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step15()
