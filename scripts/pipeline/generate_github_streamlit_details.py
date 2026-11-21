"""
Generate GitHub & Streamlit Details PDF â€” Step 7 LMS Submission
Roll No: 2500320140033 | Arohi Tyagi
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs"
ROLL = "2500320140033"
NAME = "ArohiTyagi"

RED   = RGBColor(0xA5, 0x1C, 0x30)
DARK  = RGBColor(0x15, 0x15, 0x14)
GRAY  = RGBColor(0x7F, 0x7F, 0x7F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE  = RGBColor(0x1B, 0x36, 0x5D)

GITHUB_URL   = "https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis"
STREAMLIT_URL = "https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/"

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_margins(section):
    section.top_margin = Cm(2); section.bottom_margin = Cm(2)
    section.left_margin = Cm(2); section.right_margin = Cm(2)
    section.header_distance = Cm(1.25); section.footer_distance = Cm(1.25)

def add_header(doc):
    hdr = doc.sections[0].header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.clear()
    r = p.add_run("ABES Engineering College | GitHub & Streamlit Details | Treatment Cost Trend Analysis")
    r.font.size = Pt(8); r.font.color.rgb = GRAY; r.font.name = "Calibri"

def add_footer(doc):
    ftr = doc.sections[0].footer
    ftr.is_linked_to_previous = False
    p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.clear()
    r = p.add_run("Arohi Tyagi | 2500320140033     Page ")
    r.font.size = Pt(8); r.font.color.rgb = GRAY; r.font.name = "Calibri"
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    r2 = OxmlElement('w:r')
    r2.append(fldChar1); r2.append(instrText); r2.append(fldChar2)
    p._p.append(r2)

def h1(doc, text):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '280'); sp.set(qn('w:after'), '100')
    pPr.append(sp)
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(13); r.font.color.rgb = RED; r.font.name = "Georgia"

def label_val(doc, label, val, url=False):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '40'); sp.set(qn('w:after'), '80')
    pPr.append(sp)
    r1 = p.add_run(f"{label}:  ")
    r1.bold = True; r1.font.size = Pt(10.5); r1.font.name = "Calibri"; r1.font.color.rgb = DARK
    r2 = p.add_run(val)
    r2.font.size = Pt(10.5); r2.font.name = "Calibri"
    r2.font.color.rgb = BLUE if url else DARK
    if url: r2.underline = True

def screenshot_placeholder(doc, label, instruction):
    tbl = doc.add_table(rows=2, cols=1)
    tbl.style = 'Table Grid'
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    hdr_cell = tbl.rows[0].cells[0]
    set_cell_bg(hdr_cell, "A51C30")
    hdr_cell.text = label
    for para in hdr_cell.paragraphs:
        for run in para.runs:
            run.bold = True; run.font.size = Pt(9); run.font.color.rgb = WHITE; run.font.name = "Calibri"
    body_cell = tbl.rows[1].cells[0]
    set_cell_bg(body_cell, "F5F5F5")
    body_cell.paragraphs[0].clear()
    body_cell._tc.get_or_add_tcPr()
    # Set fixed height for the placeholder box
    trPr = tbl.rows[1]._tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), '4320')  # ~7.5 cm
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)
    p = body_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '1200')
    pPr.append(sp)
    r = p.add_run(f"[ {instruction} ]")
    r.italic = True; r.font.size = Pt(10); r.font.color.rgb = GRAY; r.font.name = "Calibri"
    doc.add_paragraph()

def main():
    doc = Document()
    section = doc.sections[0]
    set_margins(section)
    add_header(doc)
    add_footer(doc)

    # Cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '800'); sp.set(qn('w:after'), '80')
    pPr.append(sp)
    logo_file = DOCS / "abes_logo.png"
    if not logo_file.exists():
        logo_file = DOCS / "abes_logo.jpg"
    if logo_file.exists():
        p.add_run().add_picture(str(logo_file), width=Inches(2.2))
        p_tag = doc.add_paragraph()
        p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tag = p_tag.add_run("ABES Engineering College, Ghaziabad")
        r_tag.bold = True; r_tag.font.size = Pt(10); r_tag.font.color.rgb = RED; r_tag.font.name = "Georgia"
    else:
        r = p.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
        r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RED; r.font.name = "Georgia"

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("GITHUB & STREAMLIT DETAILS")
    r2.bold = True; r2.font.size = Pt(22); r2.font.color.rgb = DARK; r2.font.name = "Georgia"

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run("Treatment Cost Trend Analysis  â€”  Step 7")
    r3.bold = True; r3.font.size = Pt(16); r3.font.color.rgb = RED; r3.font.name = "Georgia"

    doc.add_paragraph()

    # Quick reference table
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = 'Table Grid'
    tbl.autofit = False
    tbl.columns[0].width = Inches(2.0)
    tbl.columns[1].width = Inches(4.5)
    for label, val in [
        ("Learner",           "Arohi Tyagi"),
        ("Roll No",           "2500320140033"),
        ("Project ID",         "P106"),
        ("Domain",             "Data Analyst"),
        ("Programme / Batch",  "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("Mentor",            "Mr. Dinesh"),
        ("Roll No",           "2500320140033"),
        ("Mentor",            "Mr. Dinesh"),
        ("GitHub URL",        GITHUB_URL),
        ("Streamlit URL",     STREAMLIT_URL),
        ("Date",              datetime.date.today().strftime("%d/%m/%Y")),
    ]:
        row = tbl.add_row()
        row.cells[0].text = label
        row.cells[1].text = val
        for i, cell in enumerate(row.cells):
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9.5); run.font.name = "Calibri"; run.font.color.rgb = DARK
                    if i == 0: run.bold = True

    doc.add_page_break()

    # Section 1: GitHub
    h1(doc, "1. GitHub Repository")
    label_val(doc, "Repository URL",   GITHUB_URL, url=True)
    label_val(doc, "Owner",            "ArohiTyagi")
    label_val(doc, "Branch",           "main")
    label_val(doc, "Visibility",       "Public")
    label_val(doc, "Last Commit",      "Fix: mentor Mr. Dinesh, domain Data Analyst, regenerate all docs + PDFs")
    label_val(doc, "Total Commits",    "5")
    label_val(doc, "Key Files at Root","app.py, requirements.txt, README.md, .gitignore")

    doc.add_paragraph()
    screenshot_placeholder(doc,
        "Screenshot â€” GitHub Repository File List",
        "Paste screenshot of https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis showing the file tree"
    )

    doc.add_page_break()

    # Section 2: Streamlit
    h1(doc, "2. Streamlit Community Cloud Deployment")
    label_val(doc, "Live App URL",   STREAMLIT_URL, url=True)
    label_val(doc, "Platform",       "Streamlit Community Cloud (share.streamlit.io)")
    label_val(doc, "GitHub Repo",    GITHUB_URL, url=True)
    label_val(doc, "Branch",         "main")
    label_val(doc, "Main File",      "app.py")
    label_val(doc, "Pages Verified", "Executive Overview Â· Cost Trends Â· Outlier Detection Â· Procedure Analysis Â· Forecasting Â· Data Explorer")

    doc.add_paragraph()
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dash_img_file = DOCS / "dashboard_screenshot.png"
    if not dash_img_file.exists():
        dash_img_file = DOCS / "dashboard_screenshot.jpg"
    if dash_img_file.exists():
        p_img.add_run().add_picture(str(dash_img_file), width=Inches(6.5))
    else:
        screenshot_placeholder(doc,
            "Screenshot — Live Streamlit App (Executive Overview)",
            "Screenshot of the live app opened in an incognito window"
        )

    out_name = f"{ROLL}_{NAME}_GitHub_Streamlit_Details.docx"
    out_path = DOCS / out_name
    doc.save(str(out_path))
    print(f"Saved: docs/{out_name}")
    return out_path

if __name__ == "__main__":
    main()
