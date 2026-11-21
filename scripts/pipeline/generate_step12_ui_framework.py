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

def build_step12():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("USER INTERFACE (UI) FRAMEWORK & WIREFRAME SPECIFICATION")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Streamlit Multipage Architecture & Tailwind CSS Design System | Step 12")
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
        ("UI Technology", "Streamlit 1.38+, Tailwind CSS via st.components.v1.html, Plotly Express"),
        ("Theme Palette", "ABES Burgundy (#A51C30), Deep Navy (#1B365D), Off-White (#F4F1F1)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Screen Inventory
    doc.add_heading("1. Application Screen Inventory Matrix", level=1)
    
    ui_table = doc.add_table(rows=7, cols=6)
    ui_widths = [Inches(0.7), Inches(1.2), Inches(1.4), Inches(1.4), Inches(1.0), Inches(0.8)]
    ui_table.rows[0].cells[0].paragraphs[0].add_run("UI ID")
    ui_table.rows[0].cells[1].paragraphs[0].add_run("Screen Name")
    ui_table.rows[0].cells[2].paragraphs[0].add_run("Core Purpose")
    ui_table.rows[0].cells[3].paragraphs[0].add_run("UI Components")
    ui_table.rows[0].cells[4].paragraphs[0].add_run("Filters")
    ui_table.rows[0].cells[5].paragraphs[0].add_run("User Story")

    ui_data = [
        ("UI-01", "Executive Overview", "High-level KPI health check and cost breakdown waterfall", "Tailwind Header, 4 KPI cards, Cost Component Waterfall, Branch Bar Chart", "Year, Dept, Facility", "US-01, US-08, US-10"),
        ("UI-02", "Cost Trends & Inflation", "Longitudinal monthly cost momentum & departmental CAGR", "MoM Trend Line Chart (Mean & Median), CAGR Bar Chart, Inflation Metric Cards", "Date Range, Dept Multi-select", "US-01, US-02"),
        ("UI-03", "Outlier & Anomaly Audit", "Tukey IQR and Z-score outlier identification table", "Z-Score Scatter Plot, Tukey Fences Boxplot, Highlighted Outlier Table with CSV Export", "Z-threshold Slider, Dept Filter", "US-06, US-09"),
        ("UI-04", "Procedure & LOS Variance", "Inter-procedure variance analysis and LOS scatter correlation", "Procedure IQR Box Whiskers, LOS vs Cost Scatter Plot, Correlation Metric Card", "Department, Complexity, Gender", "US-03, US-04, US-05"),
        ("UI-05", "Forecasting & Simulation", "90-day moving average cost projection and price cap simulation", "Quarterly Forecast Line with Confidence Intervals, Scenario Adjustment Slider", "Horizon (Quarter), Price Adj %", "US-07, US-11"),
        ("UI-06", "Data Explorer & Quality", "Tabular dataset viewer, schema summary, and data health audit", "Filterable Interactive Table, Column Profiler, Null & Duplicate Check Cards", "All dimension filters, Text Search", "US-12, US-13")
    ]
    for idx, vals in enumerate(ui_data):
        row = ui_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(ui_table, ui_widths)

    doc.add_paragraph()

    # Section 2: ASCII Wireframes
    doc.add_heading("2. ASCII Structural Screen Wireframes", level=1)
    
    wireframe_text = (
        "+---------------------------------------------------------------------------------------------------+\n"
        "|  [ABES LOGO]  TREATMENT COST TREND ANALYSIS — EXECUTIVE DASHBOARD (UI-01)                        |\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "| SIDEBAR FILTERS         | [ KPI 1: Avg Cost ] [ KPI 2: Total Revenue ] [ KPI 3: ALOS ] [ KPI 4: Outliers ]|\n"
        "| > Date Range Picker     |   ₹1,32,450 (↑6.2%)   ₹4.63 Cr (MoM +4.1%)    4.8 Days (↓0.3)    9 (2.6%)      |\n"
        "| > Department MultiSelect+-------------------------------------------------------------------------+\n"
        "| > Hospital Branch       | CHARGE COMPONENT WATERFALL (Plotly Waterfall Chart)                     |\n"
        "| > Admission Type        | [Room: ₹48K] -> [Surgery: ₹62K] -> [Pharmacy: ₹18K] -> [Lab: ₹14K]      |\n"
        "| > Price Simulation Slidr+-------------------------------------------------------------------------+\n"
        "|                         | FACILITY BENCHMARK COMPARISON      | INSURANCE VS OUT-OF-POCKET SPLIT   |\n"
        "| [Download Filtered CSV] | (Bar Chart: Avg Cost per Bed Day)  | (Donut Chart: Insured 74% / OOP 26%)|\n"
        "+---------------------------------------------------------------------------------------------------+\n\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "|  STATISTICAL OUTLIER DETECTION & AUDIT REVIEW (UI-03)                                             |\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "| SIDEBAR CONTROLS        | [ ALERT: 9 Statistically Anomalous Invoices Detected Beyond Tukey 1.5*IQR ]     |\n"
        "| > Z-Score Threshold (3.0+-------------------------------------------------------------------------+\n"
        "| > Procedure Filter      | SCATTER PLOT: BILLING AMOUNT VS LENGTH OF STAY (High-Cost Outliers in Red)       |\n"
        "| > Review Status Filter  | (X: LOS Days 1-30 | Y: Total Billed Cost ₹50K - ₹6.5 Lakhs | Z-Scores Highlighted)|\n"
        "|                         +-------------------------------------------------------------------------+\n"
        "| [Export Audit Sheet]    | AUDIT TABLE: Patient ID | Procedure | Billed | Expected Q3 | Z-Score | Status |\n"
        "+---------------------------------------------------------------------------------------------------+"
    )
    p_wf = doc.add_paragraph()
    p_wf.paragraph_format.left_indent = Inches(0.1)
    r_wf = p_wf.add_run(wireframe_text)
    r_wf.font.name = "Consolas"
    r_wf.font.size = Pt(7.5)

    doc.add_paragraph()

    # Section 3: Chart Selection Rationale
    doc.add_heading("3. Chart Type Selection Rationale for Operational Questions", level=1)
    
    chart_rationales = [
        ("OQ-01 (MoM Cost Trends)", "Plotly Multi-Line Chart (Mean vs Median)", "Displays longitudinal trajectory while allowing visual comparison between average and outlier-resilient median."),
        ("OQ-02 (Cost Inflation)", "Plotly Horizontal Bar with Trend Slope Indicator", "Enables immediate ranking of clinical departments sorted by annualized CAGR percentage."),
        ("OQ-03 (LOS vs Cost Correlation)", "Plotly Scatter Plot with OLS Linear Trendline", "Exposes bivariate relationship, clustering patterns, and variance spread across age and gender cohorts."),
        ("OQ-04 (Insurance Split)", "Plotly Donut Chart & Stacked Bar", "Provides intuitive proportional breakdown of covered vs patient direct liability."),
        ("OQ-05 (Procedure Variance)", "Plotly Box-and-Whisker Plots", "Visualizes exact median, interquartile range (Q1 to Q3), and statistical Tukey outlier fences."),
        ("OQ-06 (Outlier Detection)", "Plotly Categorical Scatter with Highlighted Badges", "Immediately draws auditor attention to billing points exceeding 3 standard deviations."),
        ("OQ-07 (Cost Forecasting)", "Plotly Area & Line Chart with Confidence Ribbon", "Illustrates historical quarterly actuals alongside projected rolling trend corridor."),
        ("OQ-08 (Cost Breakdown)", "Plotly Waterfall Chart", "Demonstrates the step-by-step additive contribution of individual hospital service line charges.")
    ]
    for oq, c_type, why in chart_rationales:
        p_cr = doc.add_paragraph(style='List Bullet')
        p_cr.add_run(f"{oq} -> {c_type}: ").bold = True
        p_cr.add_run(why)

    # Section 3: Live Production Dashboard Interface
    doc.add_heading("3. Production Deployed Dashboard (Streamlit Cloud)", level=1)
    doc.add_paragraph("High-resolution capture of the live production dashboard (UI-01 Executive Overview) deployed at https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/:")
    p_live = doc.add_paragraph()
    p_live.alignment = WD_ALIGN_PARAGRAPH.CENTER
    from pathlib import Path
    dash_path = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.png"
    if not dash_path.exists():
        dash_path = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.jpg"
    if dash_path.exists():
        p_live.add_run().add_picture(str(dash_path), width=Inches(6.5))

    target_path = "docs/12_UI_Framework.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step12()
