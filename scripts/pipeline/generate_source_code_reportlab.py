"""
Generate Source Code PDF using ReportLab — fast, no Word required.
Roll No: 2500320140033 | Arohi Tyagi
"""
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 PageBreak, Table, TableStyle, HRFlowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import datetime, os

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")
DOCS = ROOT / "docs"
OUT  = DOCS / "2500320140033_ArohiTyagi_Source_Code.pdf"

RED    = colors.HexColor("#A51C30")
DARK   = colors.HexColor("#151514")
GRAY   = colors.HexColor("#7F7F7F")
LGRAY  = colors.HexColor("#F5F5F5")
WHITE  = colors.white

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

W, H = A4
MARGIN = 2 * cm

def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(GRAY)
    # Header
    canvas.drawRightString(W - MARGIN, H - MARGIN * 0.65,
        "ABES Engineering College | Source Code | Treatment Cost Trend Analysis")
    # Footer
    canvas.drawCentredString(W / 2, MARGIN * 0.5,
        f"Arohi Tyagi | 2500320140033     Page {doc.page}")
    canvas.restoreState()

def build_pdf():
    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN + 0.5*cm, bottomMargin=MARGIN,
    )

    story = []

    # Cover
    styles = getSampleStyleSheet()
    cover_title = ParagraphStyle("cov_title",
        fontName="Helvetica-Bold", fontSize=22, textColor=DARK,
        spaceAfter=8, alignment=1)
    cover_sub = ParagraphStyle("cov_sub",
        fontName="Helvetica-Bold", fontSize=15, textColor=RED,
        spaceAfter=6, alignment=1)
    cover_meta = ParagraphStyle("cov_meta",
        fontName="Helvetica", fontSize=9, textColor=GRAY,
        spaceAfter=4, alignment=1)
    normal = ParagraphStyle("norm",
        fontName="Helvetica", fontSize=10, textColor=DARK, spaceAfter=4)

    from reportlab.platypus import Image as RLImage
    logo_p = ROOT / "docs" / "abes_logo.jpg"
    if logo_p.exists():
        story.append(RLImage(str(logo_p), width=4.5*cm, height=1.3*cm))
        story.append(Spacer(1, 0.5*cm))
    else:
        story.append(Spacer(1, 2*cm))
    story.append(Paragraph("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad",
        ParagraphStyle("hdr", fontName="Helvetica-Bold", fontSize=10, textColor=RED,
                       spaceAfter=12, alignment=1)))
    story.append(Paragraph("SOURCE CODE", cover_title))
    story.append(Paragraph("Treatment Cost Trend Analysis — Step 5", cover_sub))
    story.append(Paragraph("app.py · pages · src · tests", cover_meta))
    story.append(Spacer(1, 0.5*cm))

    # Cover table
    cover_data = [
        ["Learner",           "Arohi Tyagi"],
        ["Roll No",           "2500320140033"],
        ["Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"],
        ["Project ID",         "P106"],
        ["Domain",             "Data Analyst"],
        ["Mentor",            "Mr. Dinesh"],
        ["Total Files",       f"{len(CODE_FILES)} Python source files"],
        ["Date",              datetime.date.today().strftime("%d/%m/%Y")],
    ]
    ct = Table(cover_data, colWidths=[4.5*cm, 11*cm])
    ct.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), "Helvetica"),
        ("FONTSIZE",   (0,0),(-1,-1), 9.5),
        ("FONTNAME",   (0,0),(0,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",  (0,0),(-1,-1), DARK),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0,0),(-1,-1), [WHITE, LGRAY]),
        ("TOPPADDING",  (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING", (0,0),(-1,-1), 8),
    ]))
    story.append(ct)
    story.append(PageBreak())

    # Code sections
    file_hdr_style = ParagraphStyle("fhdr",
        fontName="Helvetica-Bold", fontSize=10, textColor=WHITE,
        backColor=RED, leftIndent=6, spaceAfter=0,
        leading=16, borderPadding=(4, 6, 4, 6))

    ln_style = ParagraphStyle("ln",
        fontName="Courier", fontSize=7.5, textColor=DARK,
        spaceAfter=0, spaceBefore=0, leading=10)

    for idx, (label, path) in enumerate(CODE_FILES):
        print(f"  Processing: {label}")

        # Banner
        banner_data = [[Paragraph(f"  {label}", ParagraphStyle(
            "bann", fontName="Courier-Bold", fontSize=10, textColor=WHITE,
            leading=14))]]
        bt = Table(banner_data, colWidths=[W - 2*MARGIN])
        bt.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), RED),
            ("TOPPADDING",    (0,0),(-1,-1), 6),
            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ]))
        story.append(bt)
        story.append(Spacer(1, 4))

        try:
            content = path.read_text(encoding="utf-8")
        except Exception as e:
            content = f"# Error reading file: {e}"

        lines = content.split("\n")
        # Batch lines into chunks for efficiency
        CHUNK = 50
        for chunk_start in range(0, len(lines), CHUNK):
            chunk = lines[chunk_start:chunk_start + CHUNK]
            rows = []
            for i, line in enumerate(chunk, chunk_start + 1):
                # escape XML special chars
                safe = (line.rstrip()
                        .replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;"))
                # keep leading spaces using non-breaking space
                safe = safe.replace("    ", "\u00a0\u00a0\u00a0\u00a0")
                safe = safe.replace("  ", "\u00a0\u00a0")
                lnum = Paragraph(
                    f'<font color="#999999">{i:4d}</font>  {safe}',
                    ln_style)
                rows.append([lnum])

            tbl = Table(rows, colWidths=[W - 2*MARGIN])
            tbl.setStyle(TableStyle([
                ("FONTNAME",     (0,0),(-1,-1), "Courier"),
                ("FONTSIZE",     (0,0),(-1,-1), 7.5),
                ("LEADING",      (0,0),(-1,-1), 10),
                ("TOPPADDING",   (0,0),(-1,-1), 1),
                ("BOTTOMPADDING",(0,0),(-1,-1), 1),
                ("LEFTPADDING",  (0,0),(-1,-1), 4),
                ("ROWBACKGROUNDS",(0,0),(-1,-1), [WHITE, colors.HexColor("#FAFAFA")]),
                ("LINEBELOW",    (0,-1),(-1,-1), 0.3, colors.HexColor("#EEEEEE")),
            ]))
            story.append(tbl)

        story.append(Spacer(1, 0.3*cm))
        if idx < len(CODE_FILES) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    sz = OUT.stat().st_size / (1024 * 1024)
    print(f"\n[OK] {OUT.name}  ({sz:.2f} MB)")

if __name__ == "__main__":
    print(f"Generating Source Code PDF via ReportLab...")
    build_pdf()
