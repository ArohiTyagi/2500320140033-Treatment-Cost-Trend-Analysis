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

def build_step16():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("STREAMLIT COMMUNITY CLOUD DEPLOYMENT RUNBOOK")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Continuous Deployment Runbook & Troubleshooting Protocol | Step 16")
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
        ("Target Platform", "Streamlit Community Cloud (Connected to GitHub)"),
        ("Production URL Pattern", "https://2500320140033-treatment-cost-analysis.streamlit.app")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Pre-Deployment Checklist
    doc.add_heading("1. Pre-Deployment Repository Audit Checklist", level=1)
    checks = [
        ("Root app.py Presence", "app.py is located at the workspace root directory acting as entrypoint."),
        ("Pinned dependencies", "requirements.txt pins exact runtime dependencies (streamlit, pandas, plotly, scipy, etc.)."),
        ("Data Persistence", "data/app.db is committed into the repository under 10 MB limits."),
        ("Relative Pathing", "All file references leverage Path(__file__).resolve().parent without hardcoded absolute C: paths."),
        ("Git Hygiene", ".gitignore prevents committing local .venv, __pycache__, and temporary pytest artifacts.")
    ]
    for c_title, c_desc in checks:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(f"{c_title}: ").bold = True
        p.add_run(c_desc)

    doc.add_paragraph()

    # Section 2: Step-by-Step Deployment Guide
    doc.add_heading("2. Step-by-Step Deployment Instructions", level=1)
    steps = [
        ("Step 1: GitHub Push", "Push all workspace commits to GitHub repository: https://github.com/arohityagi/033-treatment-cost-analysis"),
        ("Step 2: Streamlit Sign In", "Navigate to share.streamlit.io and sign in with the authenticated GitHub account."),
        ("Step 3: New App Configuration", "Click 'New App', select repository, set branch to 'main', and specify main file path as 'app.py'."),
        ("Step 4: App URL Customization", "Set custom subdomain: '2500320140033-treatment-cost-analysis.streamlit.app'."),
        ("Step 5: Launch & Live Audit", "Click 'Deploy' and monitor the real-time build logs for 2-3 minutes until the live dashboard renders.")
    ]
    for s_title, s_desc in steps:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(f"{s_title}: ").bold = True
        p.add_run(s_desc)

    doc.add_paragraph()

    # Section 3: Troubleshooting Protocols
    doc.add_heading("3. Troubleshooting Common Streamlit Cloud Failure Modes", level=1)
    
    t_table = doc.add_table(rows=4, cols=3)
    t_widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    t_table.rows[0].cells[0].paragraphs[0].add_run("Error Signature")
    t_table.rows[0].cells[1].paragraphs[0].add_run("Root Cause")
    t_table.rows[0].cells[2].paragraphs[0].add_run("Resolution Action")

    t_data = [
        ("ModuleNotFoundError", "Missing dependency in requirements.txt.", "Add missing library to requirements.txt and git push."),
        ("FileNotFoundError: data/app.db", "Absolute pathing used or data directory not tracked in git.", "Ensure Path(__file__).resolve().parent / 'data' / 'app.db' is used and database is committed."),
        ("Streamlit App Exceeded Memory", "In-memory cache exceeded 1 GB container limit.", "Enforce @st.cache_data(ttl=3600) and keep SQLite dataset lean (< 5,000 rows).")
    ]
    for idx, vals in enumerate(t_data):
        row = t_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(t_table, t_widths)

    target_path = "docs/16_Deploy_Guide.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step16()
