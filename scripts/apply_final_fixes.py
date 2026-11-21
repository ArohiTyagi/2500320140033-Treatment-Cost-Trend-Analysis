"""
Update all project scripts to enforce:
1. Domain: "Data Analyst" only
2. Project ID: "P106"
3. Mentor: "Mr. Dinesh"
4. Batch: "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"
5. Roll No: "2500320140033" (fix any 2500320140033 typo)
6. Logo: Insert ONLY where required (Cover pages / placeholders), NEVER in every page header!
7. Dashboard Photo: Insert where required (GitHub/Streamlit details doc, Presentation slide 7, UI Framework, Step 17)
"""
from pathlib import Path
import re

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")

# 1. Fix Step 17 script
step17_path = ROOT / "scripts/pipeline/generate_step17_abes_project_document.py"
step17_code = step17_path.read_text(encoding="utf-8")

# Fix logo on cover: replace [Insert ABES logo here] with actual picture
logo_old = '''    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_after = Pt(5)
    r_logo = p_logo.add_run("[Insert ABES logo here]")
    r_logo.font.name = "Calibri"
    r_logo.font.size = Pt(9)
    r_logo.font.italic = True
    r_logo.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)'''

logo_new = '''    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(8)
    p_logo.paragraph_format.space_after = Pt(8)
    r_logo = p_logo.add_run()
    logo_file = Path(__file__).resolve().parents[2] / "docs" / "abes_logo.jpg"
    if logo_file.exists():
        r_logo.add_picture(str(logo_file), width=Inches(2.4))
    else:
        r_logo.text = "[ABES Engineering College]"'''

step17_code = step17_code.replace(logo_old, logo_new)

# Fix cover data in Step 17
step17_code = re.sub(
    r'\("Domain",\s*"[^"]*"\)',
    '("Domain", "Data Analyst")',
    step17_code
)
step17_code = re.sub(
    r'\("Programme / Batch",\s*"[^"]*"\)',
    '("Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01")',
    step17_code
)
step17_code = re.sub(
    r'\("GitHub repository",\s*"[^"]*"\)',
    '("GitHub repository", "https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis")',
    step17_code
)
step17_code = re.sub(
    r'\("Live app \(Streamlit\)",\s*"[^"]*"\)',
    '("Live app (Streamlit)", "https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/")',
    step17_code
)

# Fix dashboard screenshot in Section 10 UI Framework of Step 17
sec10_old = '''    for idx, vals in enumerate(t10_data, 1):
        row = t10.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t10, t10_widths, has_header=True)

    doc.add_page_break()'''

sec10_new = '''    for idx, vals in enumerate(t10_data, 1):
        row = t10.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t10, t10_widths, has_header=True)

    # Embed Live Production Dashboard Screenshot
    p_dash_hdr = doc.add_paragraph()
    p_dash_hdr.paragraph_format.space_before = Pt(14)
    p_dash_hdr.paragraph_format.space_after = Pt(4)
    r_dh = p_dash_hdr.add_run("Live Production Dashboard Interface (Executive Overview — UI-01)")
    r_dh.font.name = "Georgia"
    r_dh.font.size = Pt(11)
    r_dh.font.bold = True
    r_dh.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    p_dash_img = doc.add_paragraph()
    p_dash_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_di = p_dash_img.add_run()
    dash_file = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.jpg"
    if dash_file.exists():
        r_di.add_picture(str(dash_file), width=Inches(6.5))

    doc.add_page_break()'''

if sec10_old in step17_code:
    step17_code = step17_code.replace(sec10_old, sec10_new)

# Fix appendix texts
step17_code = step17_code.replace("DOMAIN: Data Analyst\n", "DOMAIN: Data Analyst\n")
step17_code = step17_code.replace("BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01\n", "BATCH: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01\n")
step17_code = step17_code.replace("PROJECT: <treatment cost trend analysis> | DOMAIN: <Data analysis>\n", "PROJECT: <treatment cost trend analysis> | DOMAIN: <Data Analyst>\n")

step17_path.write_text(step17_code, encoding="utf-8")
print("Updated: scripts/pipeline/generate_step17_abes_project_document.py")

# 2. Fix Step 18 Presentation script
step18_path = ROOT / "scripts/pipeline/generate_step18_abes_presentation.py"
step18_code = step18_path.read_text(encoding="utf-8")

# Fix roll number typo
step18_code = step18_code.replace("2500320140033", "2500320140033")
step18_code = step18_code.replace("https://2500320140033-treatment-cost-analysis.streamlit.app", "https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/")

# Slide 1 lines:
slide1_lines_old = '''    lines = [
        "Arohi Tyagi | Roll No 2500320140033",
        "MCA 2nd Year Data Analyst | ABES Engineering College",
        "Mentor: Mr. Dinesh",
        "20 November 2026"
    ]'''

slide1_lines_new = '''    lines = [
        "Arohi Tyagi | Roll No: 2500320140033",
        "Project ID: P106 | Domain: Data Analyst",
        "Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",
        "Mentor: Mr. Dinesh | ABES Engineering College",
        "21 November 2026"
    ]'''
step18_code = step18_code.replace(slide1_lines_old, slide1_lines_new)

# Slide 1: add logo picture instead of [ABES logo] text
slide1_logo_old = '''def add_header_dark(slide):
    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    tb_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_l = tb_logo.text_frame
    tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_l = tf_l.paragraphs[0]
    p_l.text = "[ABES logo]"
    p_l.font.name = "Calibri"
    p_l.font.size = Pt(10)
    p_l.font.italic = True
    p_l.font.color.rgb = C_PINK'''

slide1_logo_new = '''def add_header_dark(slide):
    logo_file = Path(__file__).resolve().parents[2] / "docs" / "abes_logo.jpg"
    if logo_file.exists():
        slide.shapes.add_picture(str(logo_file), Inches(0.5), Inches(0.32), width=Inches(1.8))
    else:
        tb_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
        tf_l = tb_logo.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.text = "ABES Engineering College"
        p_l.font.name = "Calibri"
        p_l.font.size = Pt(10)
        p_l.font.color.rgb = C_PINK'''
step18_code = step18_code.replace(slide1_logo_old, slide1_logo_new)

# Slide 7: add actual dashboard picture
slide7_dash_old = '''    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = C_CARD
    frame.line.color.rgb = hex_to_rgb("BFBFBF")
    frame.line.width = Pt(1)
    tf_fr = frame.text_frame
    tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_fr = tf_fr.paragraphs[0]
    p_fr.alignment = PP_ALIGN.CENTER
    p_fr.text = "[Interactive Multipage Dashboard: Overview & Cost Waterfall]"
    p_fr.font.name = "Calibri"
    p_fr.font.size = Pt(13)
    p_fr.font.italic = True
    p_fr.font.color.rgb = C_GREY'''

slide7_dash_new = '''    dash_file = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.jpg"
    if dash_file.exists():
        s7.shapes.add_picture(str(dash_file), Inches(0.5), Inches(1.25), width=Inches(5.8), height=Inches(3.7))
    else:
        frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
        frame.fill.solid()
        frame.fill.fore_color.rgb = C_CARD
        frame.line.color.rgb = hex_to_rgb("BFBFBF")
        frame.line.width = Pt(1)
        tf_fr = frame.text_frame
        tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_fr = tf_fr.paragraphs[0]
        p_fr.alignment = PP_ALIGN.CENTER
        p_fr.text = "[Interactive Multipage Dashboard: Overview & Cost Waterfall]"
        p_fr.font.name = "Calibri"
        p_fr.font.size = Pt(13)
        p_fr.font.italic = True
        p_fr.font.color.rgb = C_GREY'''
step18_code = step18_code.replace(slide7_dash_old, slide7_dash_new)

step18_path.write_text(step18_code, encoding="utf-8")
print("Updated: scripts/pipeline/generate_step18_abes_presentation.py")

# 3. Fix GitHub & Streamlit Details script
gh_path = ROOT / "scripts/pipeline/generate_github_streamlit_details.py"
gh_code = gh_path.read_text(encoding="utf-8")

# Cover table values
gh_code = re.sub(
    r'\("Learner",\s*"[^"]*"\),',
    '("Learner",           "Arohi Tyagi"),\n        ("Roll No",           "2500320140033"),\n        ("Project ID",         "P106"),\n        ("Domain",             "Data Analyst"),\n        ("Programme / Batch",  "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),\n        ("Mentor",            "Mr. Dinesh"),',
    gh_code
)

# Cover logo
cover_logo_add = '''    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:before'), '800'); sp.set(qn('w:after'), '80')
    pPr.append(sp)
    logo_file = DOCS / "abes_logo.jpg"
    if logo_file.exists():
        p.add_run().add_picture(str(logo_file), width=Inches(2.2))
        p_tag = doc.add_paragraph()
        p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tag = p_tag.add_run("ABES Engineering College, Ghaziabad")
        r_tag.bold = True; r_tag.font.size = Pt(10); r_tag.font.color.rgb = RED; r_tag.font.name = "Georgia"
    else:
        r = p.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
        r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RED; r.font.name = "Georgia"'''

gh_code = re.sub(
    r'p = doc\.add_paragraph\(\)\s+p\.alignment = WD_ALIGN_PARAGRAPH\.CENTER\s+pPr = p\._p\.get_or_add_pPr\(\)\s+sp = OxmlElement\(\'w:spacing\'\); sp\.set\(qn\(\'w:before\'\), \'1200\'\); sp\.set\(qn\(\'w:after\'\), \'120\'\)\s+pPr\.append\(sp\)\s+r = p\.add_run\("HCLTech Career Shaper Programme \| ABES Engineering College, Ghaziabad"\)\s+r\.bold = True; r\.font\.size = Pt\(10\); r\.font\.color\.rgb = RED; r\.font\.name = "Georgia"',
    cover_logo_add,
    gh_code
)

# Replace Streamlit screenshot placeholder with real image
dash_embed = '''    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dash_img_file = DOCS / "dashboard_screenshot.jpg"
    if dash_img_file.exists():
        p_img.add_run().add_picture(str(dash_img_file), width=Inches(6.5))
    else:
        screenshot_placeholder(doc,
            "Screenshot — Live Streamlit App (Executive Overview)",
            "Screenshot of the live app opened in an incognito window"
        )'''

gh_code = re.sub(
    r'screenshot_placeholder\(doc,\s*"Screenshot[^"]*Live Streamlit App[^"]*",\s*"[^"]*"\s*\)',
    dash_embed,
    gh_code
)

gh_path.write_text(gh_code, encoding="utf-8")
print("Updated: scripts/pipeline/generate_github_streamlit_details.py")

# 4. Fix Step 12 UI Framework script to include live dashboard image
ui_path = ROOT / "scripts/pipeline/generate_step12_ui_framework.py"
ui_code = ui_path.read_text(encoding="utf-8")
if "dashboard_screenshot.jpg" not in ui_code:
    sec3_add = '''
    # Section 3: Live Production Dashboard Interface
    doc.add_heading("3. Production Deployed Dashboard (Streamlit Cloud)", level=1)
    doc.add_paragraph("High-resolution capture of the live production dashboard (UI-01 Executive Overview) deployed at https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/:")
    p_live = doc.add_paragraph()
    p_live.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dash_path = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.jpg"
    if dash_path.exists():
        p_live.add_run().add_picture(str(dash_path), width=Inches(6.5))
'''
    ui_code = ui_code.replace('doc.save("docs/12_UI_Framework.docx")', sec3_add + '\n    doc.save("docs/12_UI_Framework.docx")')
    ui_path.write_text(ui_code, encoding="utf-8")
    print("Updated: scripts/pipeline/generate_step12_ui_framework.py")

# 5. Fix Source Code ReportLab script
sc_path = ROOT / "scripts/pipeline/generate_source_code_reportlab.py"
sc_code = sc_path.read_text(encoding="utf-8")
sc_code = sc_code.replace(
    '["Programme / Batch", "MCA 2nd Year Data Analyst — ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"],',
    '["Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"],\n        ["Project ID",         "P106"],\n        ["Domain",             "Data Analyst"],'
)
# Add logo on cover page
if "RLImage(str(ROOT / 'docs' / 'abes_logo.jpg')" not in sc_code and "abes_logo.jpg" not in sc_code:
    sc_code = sc_code.replace(
        'story.append(Spacer(1, 2*cm))',
        'from reportlab.platypus import Image as RLImage\n    logo_p = ROOT / "docs" / "abes_logo.jpg"\n    if logo_p.exists():\n        story.append(RLImage(str(logo_p), width=4.5*cm, height=1.3*cm))\n        story.append(Spacer(1, 0.5*cm))\n    else:\n        story.append(Spacer(1, 2*cm))'
    )
sc_path.write_text(sc_code, encoding="utf-8")
print("Updated: scripts/pipeline/generate_source_code_reportlab.py")

# 6. Fix Step 05 BRD script
brd_path = ROOT / "scripts/pipeline/generate_step05_brd.py"
brd_code = brd_path.read_text(encoding="utf-8")
brd_code = brd_code.replace("2500320140033", "2500320140033")
brd_path.write_text(brd_code, encoding="utf-8")
print("Updated: scripts/pipeline/generate_step05_brd.py")

# 7. Global batch & domain cleanup across all .py and .md files
BATCH_STR = "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"
replacements = [
    ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01", BATCH_STR),
    ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01", BATCH_STR),
    ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01", BATCH_STR),
    ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01", BATCH_STR),
    ("2500320140033", "2500320140033"),
    ("Data Analyst", "Data Analyst"),
    ("Data Analyst", "Data Analyst"),
    ("Data Analyst", "Data Analyst"),
]

for f in ROOT.rglob("*"):
    if f.suffix not in (".py", ".md") or ".git" in str(f) or ".venv" in str(f):
        continue
    try:
        text = f.read_text(encoding="utf-8")
        orig = text
        for old, new in replacements:
            text = text.replace(old, new)
        if text != orig:
            f.write_text(text, encoding="utf-8")
            print(f"  Cleaned text in: {f.name}")
    except Exception as e:
        pass

print("\nAll script updates completed successfully.")
