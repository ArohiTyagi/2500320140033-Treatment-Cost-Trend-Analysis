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

def build_step07_trd():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("TECHNICAL REQUIREMENTS DOCUMENT (TRD)")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Architecture & Technical Blueprint Specification | Step 7")
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
        ("Stack Baseline", "Python 3.12, Pandas, NumPy, SciPy, Plotly, Streamlit, Tailwind, SQLite, Pytest"),
        ("Hosting Target", "Streamlit Community Cloud (via GitHub)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Architecture
    doc.add_heading("1. System Architecture & Data Flow", level=1)
    doc.add_paragraph(
        "The application employs a layered data analysis pipeline moving from persistent storage through caching, transformation, "
        "and presentation tiers. The architecture is represented below in Mermaid diagram notation:"
    )

    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.left_indent = Inches(0.2)
    r_arch = p_arch.add_run(
        "flowchart TD\n"
        "    A[(SQLite: data/app.db)] -->|Raw Ingestion| B[Data Access Layer: src/db.py]\n"
        "    B -->|Type Validation & Cleaning| C[Transformation Layer: src/clean.py]\n"
        "    C -->|st.cache_data In-Memory DF| D[Analytics Engine: src/analysis.py]\n"
        "    D -->|MoM Trends, Outliers, CAGR| E[Presentation Layer: app.py & pages/]\n"
        "    E -->|Tailwind CSS Components| F[Interactive Dashboard UI]\n"
        "    E -->|Plotly Visualizations| F"
    )
    r_arch.font.name = "Consolas"
    r_arch.font.size = Pt(9)

    doc.add_paragraph()

    # Section 2: Folder Structure
    doc.add_heading("2. Repository & Directory Structure", level=1)
    doc.add_paragraph("The project follows a standard modular Python data application layout:")

    p_fs = doc.add_paragraph()
    p_fs.paragraph_format.left_indent = Inches(0.2)
    r_fs = p_fs.add_run(
        "033TrendCostAnalysis/\n"
        "├── .streamlit/\n"
        "│   └── config.toml             # Theme settings (Burgundy A51C30, Navy, White)\n"
        "├── app.py                      # Application entry point & Executive Overview\n"
        "├── pages/\n"
        "│   ├── 1_Cost_Trends.py        # MoM & CAGR Inflation analysis (OQ-01, OQ-02)\n"
        "│   ├── 2_Outlier_Detection.py  # Tukey IQR & Z-score anomaly audit (OQ-06)\n"
        "│   ├── 3_Procedure_Variance.py # Inter-procedure variance & LOS correlation (OQ-03, OQ-05)\n"
        "│   ├── 4_Forecasting.py        # 90-day moving average cost projection (OQ-07)\n"
        "│   └── 5_Data_Explorer.py      # Filterable tabular dataset & CSV export\n"
        "├── src/\n"
        "│   ├── __init__.py\n"
        "│   ├── db.py                   # SQLite connection, query caching & view extraction\n"
        "│   ├── clean.py                # Data validation, type casting & outlier flagging\n"
        "│   ├── analysis.py             # Statistical computations (CAGR, IQR, Pearson r, MAPE)\n"
        "│   ├── ui.py                   # Tailwind CDN iframe cards & styling helpers\n"
        "│   └── generate_data.py        # Reproducible synthetic dataset generator\n"
        "├── data/\n"
        "│   ├── app.db                  # Relational SQLite database with integrity constraints\n"
        "│   └── treatment_cost_data.xlsx# Complete workbook with sheets per entity\n"
        "├── tests/\n"
        "│   ├── test_data.py            # SQLite schema integrity & non-null test suite\n"
        "│   ├── test_analysis.py        # Unit tests for CAGR, IQR, correlation logic\n"
        "│   └── test_app.py             # Streamlit AppTest page rendering & navigation\n"
        "├── docs/                       # Project lifecycle deliverables (01 through 18)\n"
        "├── requirements.txt            # Pinned dependencies for Streamlit Cloud\n"
        "└── README.md                   # Setup guide and architecture overview"
    )
    r_fs.font.name = "Consolas"
    r_fs.font.size = Pt(8.5)

    doc.add_paragraph()

    # Section 3: Functional Requirements
    doc.add_heading("3. Functional Requirements (FR)", level=1)
    
    fr_table = doc.add_table(rows=11, cols=4)
    fr_widths = [Inches(1.0), Inches(2.6), Inches(1.4), Inches(1.5)]
    fr_table.rows[0].cells[0].paragraphs[0].add_run("FR ID")
    fr_table.rows[0].cells[1].paragraphs[0].add_run("Functional Requirement Description")
    fr_table.rows[0].cells[2].paragraphs[0].add_run("Linked BR / US")
    fr_table.rows[0].cells[3].paragraphs[0].add_run("Target Module")

    fr_data = [
        ("FR-01", "Query SQLite database and cache analysis view in session memory.", "BR-01 / US-01", "src/db.py"),
        ("FR-02", "Compute monthly mean, median, and MoM growth rate of billed costs.", "BR-01 / US-01", "src/analysis.py"),
        ("FR-03", "Calculate Compound Annual Growth Rate (CAGR) per clinical department.", "BR-02 / US-02", "src/analysis.py"),
        ("FR-04", "Calculate Pearson correlation coefficient between LOS and total cost.", "BR-03 / US-03", "src/analysis.py"),
        ("FR-05", "Compute insurance coverage ratio and out-of-pocket patient percentage.", "BR-04 / US-04", "src/analysis.py"),
        ("FR-06", "Filter cost outliers using Tukey's fences (1.5 * IQR) and Z-score (>3.0).", "BR-05 / US-06", "src/analysis.py"),
        ("FR-07", "Compute interquartile variance (IQR, Q1, Q3) across surgical procedures.", "BR-06 / US-05", "src/analysis.py"),
        ("FR-08", "Generate 90-day rolling quarterly cost forecasts with confidence bands.", "BR-07 / US-07", "src/analysis.py"),
        ("FR-09", "Decompose patient costs into Room, Surgery, Pharmacy, and Lab charges.", "BR-08 / US-08", "src/analysis.py"),
        ("FR-10", "Render responsive Tailwind KPI summary metric cards in Streamlit.", "BR-01 / US-01", "src/ui.py")
    ]
    for idx, vals in enumerate(fr_data):
        row = fr_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(fr_table, fr_widths)

    doc.add_paragraph()

    # Section 4: Data Layer
    doc.add_heading("4. Data Layer & Caching Strategy", level=1)
    doc.add_paragraph(
        "To guarantee zero latency on Streamlit Community Cloud, data queries are executed once upon startup using "
        "`@st.cache_data(ttl=3600)`. SQLite foreign keys are enforced via `PRAGMA foreign_keys = ON;`. "
        "A pre-joined view `v_treatment_cost_analytics` flattens fact and dimension tables into an analytical structure "
        "enabling instantaneous Pandas column filtering without recurring relational joins."
    )

    # Section 5: Statistical Methods
    doc.add_heading("5. Statistical Methodology for Operational Questions", level=1)
    
    stat_methods = [
        ("OQ-01 (MoM Cost Trend)", "Group-by calendar month; Mean = sum(x)/n; Median = 50th percentile; MoM Delta % = ((Cost_t - Cost_t-1) / Cost_t-1) * 100."),
        ("OQ-02 (Cost Inflation CAGR)", "CAGR = ((Cost_End / Cost_Start) ^ (1 / n_years)) - 1; Linear regression slope on monthly mean cost."),
        ("OQ-03 (LOS Correlation)", "Pearson r = cov(LOS, Cost) / (sigma_LOS * sigma_Cost); Spearman rank order for skewed non-normal cohorts."),
        ("OQ-04 (Insurance Split)", "Coverage Ratio = (Insurance_Paid / Total_Billed) * 100; OOP Ratio = (Out_Of_Pocket / Total_Billed) * 100."),
        ("OQ-05 (Procedure Variance)", "IQR = Q3 (75th percentile) - Q1 (25th percentile); Standard Deviation = sqrt(variance)."),
        ("OQ-06 (Outlier Detection)", "Tukey Fences: Lower = Q1 - 1.5 * IQR; Upper = Q3 + 1.5 * IQR. Z-Score: Z = (x - mu) / sigma > 3.0."),
        ("OQ-07 (Cost Projection)", "3-period rolling average with exponential dampening; MAPE = (1/n) * sum(|Actual - Forecast| / Actual) * 100."),
        ("OQ-08 (Cost Breakdown)", "Waterfall additive decomposition: Total = Room + Surgery + Pharmacy + Lab."),
        ("OQ-09 (Admission Types)", "One-way ANOVA (F-statistic) and Kruskal-Wallis H-test for non-parametric group variance."),
        ("OQ-10 (Facility Benchmark)", "Average Cost Per Bed Day = Total Cost / LOS; Variance = (Facility_Avg - Network_Median) / Network_Median.")
    ]
    for oq, method in stat_methods:
        p_sm = doc.add_paragraph(style='List Bullet')
        p_sm.add_run(f"{oq}: ").bold = True
        p_sm.add_run(method)

    doc.add_paragraph()

    # Section 6: UI & Styling
    doc.add_heading("6. User Interface & Styling Architecture", level=1)
    doc.add_paragraph(
        "The interface incorporates Tailwind CSS components injected into Streamlit via `st.components.v1.html` and custom CSS. "
        "The primary color palette features ABES Crimson (`#A51C30`), Deep Navy (`#1B365D`), Slate Gray (`#5E5E5B`), and Card Cream (`#F4F1F1`). "
        "All visual components adapt responsively to varying screen resolutions."
    )

    # Section 7: Performance, Error Handling & Security
    doc.add_heading("7. Performance, Error Handling & Security", level=1)
    doc.add_paragraph(
        "- Performance: Sub-second response times achieved through columnar Pandas operations and pre-computed analytical views.\n"
        "- Error Handling: Graceful fallback exception handling with informative UI error banners if database records are empty.\n"
        "- Security: Serverless, read-only analytical mode over SQLite; completely synthetic dataset with zero HIPAA/PII compliance liabilities."
    )

    # Section 8: requirements.txt
    doc.add_heading("8. Pinned Dependency Specifications (requirements.txt)", level=1)
    p_req = doc.add_paragraph()
    p_req.paragraph_format.left_indent = Inches(0.2)
    r_req = p_req.add_run(
        "streamlit>=1.38.0\n"
        "pandas>=2.2.0\n"
        "numpy>=1.26.0\n"
        "plotly>=5.24.0\n"
        "scipy>=1.13.0\n"
        "statsmodels>=0.14.0\n"
        "openpyxl>=3.1.0\n"
        "python-docx>=1.1.0\n"
        "python-pptx>=1.0.0\n"
        "pytest>=8.3.0"
    )
    r_req.font.name = "Consolas"
    r_req.font.size = Pt(9)

    doc.add_paragraph()

    # Section 9: Deployment Requirements
    doc.add_heading("9. Deployment Requirements", level=1)
    doc.add_paragraph(
        "1. Host: Streamlit Community Cloud.\n"
        "2. Repository: Public GitHub repository containing `app.py`, `pages/`, `src/`, `data/app.db`, and `requirements.txt`.\n"
        "3. Main file path: `app.py`.\n"
        "4. Python version: 3.11 or 3.12."
    )

    # Section 10: Technical Risks
    doc.add_heading("10. Technical Risks & Guardrails", level=1)
    doc.add_paragraph(
        "Memory footprint exhaustion is prevented by maintaining the dataset size under 5,000 rows (~2 MB SQLite footprint). "
        "Browser cache invalidation is handled by setting strict TTL timeouts on Streamlit data loaders."
    )

    target_path = "docs/07_TRD.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step07_trd()
