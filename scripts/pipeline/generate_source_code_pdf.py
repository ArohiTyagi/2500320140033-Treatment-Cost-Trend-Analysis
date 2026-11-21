"""
Generate Source Code PDF — Step 5 LMS Submission
Combines all project .py files into one Word doc then saves.
Roll No: 2500320140033 | Arohi Tyagi
"""
import os
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

CODE_FILES = [
    ("app.py",                             ROOT / "app.py"),
    ("pages/1_Cost_Trends.py",             ROOT / "pages/1_Cost_Trends.py"),
    ("pages/2_Outlier_Detection.py",       ROOT / "pages/2_Outlier_Detection.py"),
    ("pages/3_Procedure_Analysis.py",      ROOT / "pages/3_Procedure_Analysis.py"),
    ("pages/4_Forecasting.py",             ROOT / "pages/4_Forecasting.py"),
    ("pages/5_Data_Explorer.py",           ROOT / "pages/5_Data_Explorer.py"),
    ("src/db.py",                          ROOT / "src/db.py"),
    ("src/clean.py",                       ROOT / "src/clean.py"),
    ("src/analysis.py",                    ROOT / "src/analysis.py"),
    ("src/ui.py",                          ROOT / "src/ui.py"),
    ("src/generate_data.py",               ROOT / "src/generate_data.py"),
    ("tests/test_data.py",                 ROOT / "tests/test_data.py"),
    ("tests/test_analysis.py",             ROOT / "tests/test_analysis.py"),
    ("tests/test_app.py",                  ROOT / "tests/test_app.py"),
]

RED    = RGBColor(0xA5, 0x1C, 0x30)
DARK   = RGBColor(0x15, 0x15, 0x14)
GRAY   = RGBColor(0x7F, 0x7F, 0x7F)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_page_border(doc):
    section = doc.sections[0]
    sectPr = section._sectPr
    pgBorders = OxmlElement('w:pgBorders')
    pgBorders.set(qn('w:offsetFrom'), 'page')
    for edge in ('top','left','bottom','right'):
        bd = OxmlElement(f'w:{edge}')
        bd.set(qn('w:val'), 'none')
        pgBorders.append(bd)
    sectPr.append(pgBorders)

def set_margins(section):
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(2)
    section.right_margin  = Cm(2)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)

def add_header(doc):
    section = doc.sections[0]
    hdr = section.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.clear()
    run = p.add_run("ABES Engineering College | Source Code | Treatment Cost Trend Analysis")
    run.font.size = Pt(8)
    run.font.color.rgb = GRAY
    run.font.name = "Calibri"

def add_footer(doc):
    section = doc.sections[0]
    ftr = section.footer
    ftr.is_linked_to_previous = False
    p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    run = p.add_run("Arohi Tyagi | 2500320140033")
    run.font.size = Pt(8)
    run.font.color.rgb = GRAY
    run.font.name = "Calibri"
    run2 = p.add_run("     Page ")
    run2.font.size = Pt(8)
    run2.font.color.rgb = GRAY
    run2.font.name = "Calibri"
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    r = OxmlElement('w:r')
    r.append(fldChar1); r.append(instrText); r.append(fldChar2)
    p._p.append(r)

def cover_page(doc):
    # Title block
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    spb = OxmlElement('w:spacing'); spb.set(qn('w:before'), '1680'); spb.set(qn('w:after'), '120')
    pPr.append(spb)
    r = p.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RED; r.font.name = "Georgia"

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("SOURCE CODE")
    r2.bold = True; r2.font.size = Pt(26); r2.font.color.rgb = DARK; r2.font.name = "Georgia"

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run("Treatment Cost Trend Analysis")
    r3.bold = True; r3.font.size = Pt(18); r3.font.color.rgb = RED; r3.font.name = "Georgia"

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run("app.py · pages · src · tests")
    r4.font.size = Pt(10); r4.font.color.rgb = GRAY; r4.font.name = "Calibri"

    # Details table
    doc.add_paragraph()
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = 'Table Grid'
    tbl.autofit = False
    tbl.columns[0].width = Inches(2.0)
    tbl.columns[1].width = Inches(4.5)
    rows_data = [
        ("Learner", "Arohi Tyagi"),
        ("Roll No", "2500320140033"),
        ("Programme / Batch", "MCA 2nd Year Data Analyst — ABES_OND26_MCA"),
        ("Mentor", "Mr. Dinesh"),
        ("GitHub Repository", "https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis"),
        ("Total Files", f"{len(CODE_FILES)} Python source files"),
        ("Date", datetime.date.today().strftime("%d/%m/%Y")),
    ]
    for label, val in rows_data:
        row = tbl.add_row()
        row.cells[0].text = label
        row.cells[1].text = val
        for i, cell in enumerate(row.cells):
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9.5)
                    run.font.name = "Calibri"
                    run.font.color.rgb = DARK
                    if i == 0:
                        run.bold = True
    doc.add_page_break()

def add_file_section(doc, label, filepath, first=False):
    if not first:
        doc.add_page_break()

    # Filename banner row
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.rows[0].cells[0]
    cell.text = ""
    set_cell_bg(cell, "A51C30")
    p = cell.paragraphs[0]
    run = p.add_run(label)
    run.bold = True
    run.font.size = Pt(10)
    run.font.name = "Consolas"
    run.font.color.rgb = WHITE

    doc.add_paragraph()  # small gap

    try:
        content = filepath.read_text(encoding="utf-8")
    except Exception as e:
        content = f"# Could not read file: {e}"

    lines = content.split("\n")
    for i, line in enumerate(lines, 1):
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), '240')
        sp.set(qn('w:lineRule'), 'auto')
        pPr.append(sp)

        # Line number
        rn = p.add_run(f"{i:4d}  ")
        rn.font.size = Pt(8)
        rn.font.name = "Consolas"
        rn.font.color.rgb = GRAY

        # Code
        rc = p.add_run(line.rstrip())
        rc.font.size = Pt(8.5)
        rc.font.name = "Consolas"
        rc.font.color.rgb = DARK

def main():
    doc = Document()
    section = doc.sections[0]
    set_margins(section)
    add_header(doc)
    add_footer(doc)
    cover_page(doc)

    for idx, (label, path) in enumerate(CODE_FILES):
        print(f"  Adding: {label}")
        add_file_section(doc, label, path, first=(idx == 0))

    out_name = f"{ROLL}_{NAME}_Source_Code.docx"
    out_path = DOCS / out_name
    doc.save(str(out_path))
    print(f"\nSaved: docs/{out_name}")
    return out_path

if __name__ == "__main__":
    main()
