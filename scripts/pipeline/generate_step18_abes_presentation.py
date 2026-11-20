import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

# Brand Colors
C_RED = hex_to_rgb("A51C30")
C_BLACK = hex_to_rgb("151514")
C_DARK = hex_to_rgb("3A3A38")
C_GREY = hex_to_rgb("5E5E5B")
C_CARD = hex_to_rgb("F4F1F1")
C_PINK = hex_to_rgb("D9B5BC")
C_WHITE = hex_to_rgb("FFFFFF")
C_BORDER = hex_to_rgb("D9D9D9")

def add_header_dark(slide):
    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    tb_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_l = tb_logo.text_frame
    tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_l = tf_l.paragraphs[0]
    p_l.text = "[ABES logo]"
    p_l.font.name = "Calibri"
    p_l.font.size = Pt(10)
    p_l.font.italic = True
    p_l.font.color.rgb = C_PINK

    # Program text (2.2,.35,7.3,.5) 11 pink, right, middle, default insets
    tb_prog = slide.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_p = tb_prog.text_frame
    tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_p = tf_p.paragraphs[0]
    p_p.alignment = PP_ALIGN.RIGHT
    p_p.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p_p.font.name = "Calibri"
    p_p.font.size = Pt(11)
    p_p.font.color.rgb = C_PINK

def add_content_chrome(slide, title_text, slide_num):
    # Title (.5,.3,9,.75) Cambria 28 bold black, centred, middle, default insets
    tb_t = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf_t = tb_t.text_frame
    tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    p_t.text = title_text
    p_t.font.name = "Cambria"
    p_t.font.size = Pt(28)
    p_t.font.bold = True
    p_t.font.color.rgb = C_BLACK

    # Footer (.5,5.2,7.5,.3) 9 grey, middle, default insets
    tb_f = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_f = tb_f.text_frame
    tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_f = tf_f.paragraphs[0]
    p_f.text = "HCLTech Career Shaper ABES Engineering College | 033 - Treatment Cost Trend Analysis"
    p_f.font.name = "Calibri"
    p_f.font.size = Pt(9)
    p_f.font.color.rgb = C_GREY

    # Slide number (9,5.2,.5,.3) 9, colour 8A8A86, right, default insets
    tb_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_n = tb_num.text_frame
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_n = tf_n.paragraphs[0]
    p_n.alignment = PP_ALIGN.RIGHT
    p_n.text = str(slide_num)
    p_n.font.name = "Calibri"
    p_n.font.size = Pt(9)
    p_n.font.color.rgb = hex_to_rgb("8A8A86")

def add_card(slide, x, y, w, h):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_CARD
    shape.line.color.rgb = C_CARD
    return shape

def add_badge(slide, x, y, text):
    # Circle .42 red; t centred, 12 bold white
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_RED
    shape.line.color.rgb = C_RED
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = str(text)
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    return shape

def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text):
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_text)
    
    # HEAD (x+.75, y+.17, w-.9, .45): 15 bold black, middle
    tb_h = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
    tf_h = tb_h.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(14)
    p_h.font.bold = True
    p_h.font.color.rgb = C_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85): 13 dark, top
    tb_b = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(12)
    p_b.font.color.rgb = C_DARK

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    add_card(slide, x, y, w, h)
    # HEAD (x+.2, y+.17, w-.35, .45)
    tb_h = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
    tf_h = tb_h.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(14)
    p_h.font.bold = True
    p_h.font.color.rgb = C_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85)
    tb_b = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(12)
    p_b.font.color.rgb = C_DARK

def set_speaker_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: DARK TITLE SLIDE
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.background
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_BLACK
    add_header_dark(s1)

    # Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    tb_t1 = s1.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t1 = tb_t1.text_frame
    tf_t1.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t1 = tf_t1.paragraphs[0]
    p_t1.text = "033 - Treatment Cost Trend Analysis"
    p_t1.font.name = "Cambria"
    p_t1.font.size = Pt(36)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line
    tb_l1 = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_l1 = tb_l1.text_frame
    tf_l1.word_wrap = True
    
    lines = [
        "Arohi Tyagi | Roll No 2500320140033",
        "MCA 2nd Year Data Analyst | ABES Engineering College",
        "Mentor: Prof. Alpana Lodhi",
        "20 November 2026"
    ]
    for idx, l in enumerate(lines):
        p = tf_l1.paragraphs[0] if idx == 0 else tf_l1.add_paragraph()
        p.text = l
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.color.rgb = C_CARD if idx < 3 else C_PINK
        if idx > 0:
            p.space_before = Pt(4)

    set_speaker_notes(s1, "Good morning respected evaluators. My name is Arohi Tyagi, roll number 2500320140033. Today I present my MCA 2nd Year capstone project, Treatment Cost Trend Analysis, developed under the mentorship of Prof. Alpana Lodhi. This platform provides hospital administrators and financial leaders with clear visibility into treatment cost inflation, bed occupancy efficiency, and anomalous billing patterns.")

    # =========================================================================
    # SLIDE 2: CONTENT "Problem and How Might We"
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s2, "Problem and How Might We", 2)

    # BADGE CARD (.5,1.25,4.3,3.7) P, Problem statement: 2-3 sentences
    prob_text = (
        "Healthcare administrators, clinical heads, and insurance adjudicators lack an integrated, "
        "data-driven analytical tool to track treatment cost inflation, evaluate departmental expenditure "
        "variance, and detect anomalous billing patterns.\n\n"
        "This absence leads to severe budget unpredictability, inefficient clinical resource allocation, "
        "and elevated out-of-pocket financial burdens on patients."
    )
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem Statement", prob_text)

    # BADGE CARD (5.2,1.25,4.3,1.75) 1, HMW-01: statement
    hmw1_text = "How might we empower hospital financial controllers to track month-over-month treatment cost trends across departments so that budget overruns are caught proactively?"
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Trend Tracking", hmw1_text)

    # BADGE CARD (5.2,3.2,4.3,1.75) 2, HMW-02: statement
    hmw2_text = "How might we equip claims managers with automated outlier detection to instantly flag abnormal billing amounts and avoid audit penalties?"
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Outlier Audit", hmw2_text)

    set_speaker_notes(s2, "On this slide we establish the core business problem and design-thinking opportunity statements. Hospitals face steep medical inflation but lack timely cost tracking. We framed ten How Might We statements; HMW-01 focuses on proactive monthly cost drift tracking, while HMW-02 targets automated billing anomaly detection to eliminate manual auditing delays.")

    # =========================================================================
    # SLIDE 3: CONTENT "Stakeholders and users"
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s3, "Stakeholders and users", 3)

    sh_cards = [
        (0.5, "1", "Primary User", "Hospital CFO / Finance Lead\n\nNeeds: Longitudinal expenditure tracking and quarterly budget forecasting.\n\nInfluence: High"),
        (3.55, "2", "Secondary User", "Medical Billing & Claims Lead\n\nNeeds: Automated billing outlier detection before insurance claim submission.\n\nInfluence: Medium"),
        (6.6, "3", "Mentor & Evaluator", "Prof. Alpana Lodhi (ABES)\n\nNeeds: Methodological rigor, data integrity, automated testing, and rubric compliance.\n\nInfluence: High")
    ]
    for x, badge, head, body in sh_cards:
        add_badge_card(s3, x, 1.35, 2.85, 3.5, badge, head, body)

    set_speaker_notes(s3, "Here we map the three primary stakeholders who interact with or evaluate the system. Our primary user is the Hospital CFO who requires accurate budget forecasting. The secondary user is the Claims Auditor who flags anomalous patient invoices. Finally, our mentor and evaluator ensures adherence to high software engineering and academic standards.")

    # =========================================================================
    # SLIDE 4: CONTENT "Requirements and traceability"
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s4, "Requirements and traceability", 4)

    # TABLE at (.5,1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2
    t_shape = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    table = t_shape.table
    col_w = [1.0, 3.4, 1.0, 1.2, 1.2, 1.2]
    for idx, w in enumerate(col_w):
        table.columns[idx].width = Inches(w)

    headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_RED
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    t_data = [
        ("HMW-01", "Track MoM treatment cost trends (mean & median)", "Must", "US-01", "TC-F-01", "UI-02"),
        ("HMW-02", "Compute annual compound inflation rate (CAGR)", "Must", "US-02", "TC-F-02", "UI-02"),
        ("HMW-03", "Analyze Length of Stay (LOS) vs cost correlation", "Must", "US-03", "TC-F-03", "UI-04"),
        ("HMW-06", "Automated outlier detection (Tukey 1.5*IQR)", "Must", "US-06", "TC-F-06", "UI-03")
    ]
    for r_idx, row_vals in enumerate(t_data, 1):
        for c_idx, val in enumerate(row_vals):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(11)
            p.font.color.rgb = C_BLACK

    # Footnote .3 below table
    tb_fn = s4.shapes.add_textbox(Inches(0.5), Inches(4.3), Inches(9.0), Inches(0.35))
    tf_fn = tb_fn.text_frame
    p_fn = tf_fn.paragraphs[0]
    p_fn.text = "Full matrix: Project Document, section 4"
    p_fn.font.name = "Calibri"
    p_fn.font.size = Pt(12)
    p_fn.font.italic = True
    p_fn.font.color.rgb = C_GREY

    set_speaker_notes(s4, "Traceability is the backbone of our engineering process. Every single HMW statement links directly to a formal business requirement, an agile user story, an automated test case, and a dedicated UI screen. This guarantees complete functional coverage and zero orphaned requirements.")

    # =========================================================================
    # SLIDE 5: CONTENT "Data and preparation"
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s5, "Data and preparation", 5)

    stats = [
        (0.5, "350", "Inpatient Fact Rows"),
        (2.8, "32", "Analytical Columns"),
        (5.1, "46 Mo", "Longitudinal Period"),
        (7.4, "100%", "Referential Parity")
    ]
    for x, val, lbl in stats:
        # Value y 1.3 h .8 Cambria 30 bold red, middle
        tb_v = s5.shapes.add_textbox(Inches(x), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_v = tb_v.text_frame
        tf_v.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_v = tf_v.paragraphs[0]
        p_v.alignment = PP_ALIGN.CENTER
        p_v.text = val
        p_v.font.name = "Cambria"
        p_v.font.size = Pt(30)
        p_v.font.bold = True
        p_v.font.color.rgb = C_RED

        # Label y 2.1 h .35, 12 grey, middle
        tb_l = s5.shapes.add_textbox(Inches(x), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_l = tb_l.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.alignment = PP_ALIGN.CENTER
        p_l.text = lbl
        p_l.font.name = "Calibri"
        p_l.font.size = Pt(12)
        p_l.font.color.rgb = C_GREY

    # PLAIN CARD (.5,2.7,4.3,2.25) Source
    src_text = (
        "Dataset Source: Synthetic relational healthcare billing database generated with fixed random seed 42.\n\n"
        "Row Grain: One record represents one complete inpatient hospitalization stay, containing patient demographics, "
        "clinical procedures, length of stay, and itemized billing charges."
    )
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source & Grain", src_text)

    # PLAIN CARD (5.2,2.7,4.3,2.25) Cleaning
    cln_text = (
        "Top 3 Data Quality Fixes:\n"
        "1. Missing Schemes: Imputed unrecorded intake records as 'Unspecified'.\n"
        "2. Minimum Stay: Enforced positive lower bound (LOS >= 1 day).\n"
        "3. Financial Parity: Validated that Room + Surgery + Meds + Lab equals Total Billed Cost across 100% of rows."
    )
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Data Cleaning & Integrity", cln_text)

    set_speaker_notes(s5, "Our data architecture is built on a 350-row synthetic inpatient billing fact table spanning 46 months from 2023 to 2026. Each row captures a hospitalization event. We implemented automated data cleaning in src/clean.py ensuring zero negative stays, imputing missing insurance intake types, and enforcing strict additive cost balance across all four billing components.")

    # =========================================================================
    # SLIDE 6: CONTENT "Solution and technology"
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s6, "Solution and technology", 6)

    # 5 steps, x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2
    steps = [
        (0.5, "Raw Data\n(SQLite / xlsx)", False),
        (2.35, "Clean & KPIs\n(Python, pandas)", False),
        (4.2, "Storage Mart\n(SQLite app.db)", False),
        (6.05, "Dashboard\n(Streamlit UI)", True), # Red fill, white text
        (7.9, "Cloud Deploy\n(GitHub -> Cloud)", False)
    ]
    for x, text, is_red in steps:
        shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(1.6), Inches(1.2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_RED if is_red else C_CARD
        shape.line.color.rgb = C_RED if is_red else C_CARD
        tf = shape.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = text
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE if is_red else C_BLACK

    # Arrows between steps
    arrows = [2.1, 3.95, 5.8, 7.65]
    for ax in arrows:
        tb_a = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_a = tb_a.text_frame
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_a = tf_a.paragraphs[0]
        p_a.alignment = PP_ALIGN.CENTER
        p_a.text = ">"
        p_a.font.name = "Calibri"
        p_a.font.size = Pt(18)
        p_a.font.bold = True
        p_a.font.color.rgb = C_RED

    # Tools line (.5,3.15,9,.5) 13 dark, middle
    tb_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_tools = tb_tools.text_frame
    tf_tools.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tools = tf_tools.paragraphs[0]
    p_tools.text = "Tools: Python 3.12 · Pandas · NumPy · SciPy · Plotly · Streamlit · Tailwind CSS · SQLite · Pytest"
    p_tools.font.name = "Calibri"
    p_tools.font.size = Pt(13)
    p_tools.font.bold = True
    p_tools.font.color.rgb = C_DARK

    # Run command (.5,3.75,9,.5) 13 italic grey, middle
    tb_cmd = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_cmd = tb_cmd.text_frame
    tf_cmd.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_cmd = tf_cmd.paragraphs[0]
    p_cmd.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p_cmd.font.name = "Calibri"
    p_cmd.font.size = Pt(13)
    p_cmd.font.italic = True
    p_cmd.font.color.rgb = C_GREY

    set_speaker_notes(s6, "Here is our end-to-end pipeline. Raw data flows into an in-memory cleaning pipeline, is persisted in SQLite, and is rendered as interactive Plotly charts in Streamlit. The entire technology stack consists of 100% free and open-source tools, requiring zero external server infrastructure.")

    # =========================================================================
    # SLIDE 7: CONTENT "Dashboard"
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s7, "Dashboard", 7)

    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF
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
    p_fr.font.color.rgb = C_GREY

    # BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4
    screens = [
        (1.3, "1", "UI-01", "Executive Overview & Waterfall"),
        (2.2, "2", "UI-02", "Longitudinal Trends & CAGR"),
        (3.1, "3", "UI-03", "Statistical Outlier Audit"),
        (4.0, "4", "UI-04", "Procedure Variance & LOS Scatter")
    ]
    for y, b_num, code, purpose in screens:
        add_badge(s7, 6.7, y, b_num)
        tb_s = s7.shapes.add_textbox(Inches(7.25), Inches(y), Inches(2.25), Inches(0.42))
        tf_s = tb_s.text_frame
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_s = tf_s.paragraphs[0]
        r_c = p_s.add_run()
        r_c.text = f"{code} "
        r_c.font.bold = True
        r_c.font.size = Pt(12)
        r_c.font.color.rgb = C_RED
        r_p = p_s.add_run()
        r_p.text = purpose
        r_p.font.size = Pt(11)
        r_p.font.color.rgb = C_BLACK

    set_speaker_notes(s7, "This slide illustrates our Streamlit user interface. The left panel shows the main dashboard featuring the KPI cards and charge component waterfall chart. On the right are our four primary screens: Executive Overview, Cost Trends, Outlier Detection, and Procedure Variance analysis.")

    # =========================================================================
    # SLIDE 8: CONTENT "Key findings"
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s8, "Key findings", 8)

    findings = [
        (0.5, "7.8%", "Cardiology & Oncology Cost Inflation", "Evidence: Annualized CAGR of mean cost from 2023 to 2026."),
        (3.55, "r = 0.68", "Length of Stay Drives Total Cost", "Evidence: Pearson correlation test yields r = 0.68 (p < 0.001)."),
        (6.6, "2.6%", "Statistical Billing Anomaly Rate", "Evidence: 9 claims flagged via Tukey IQR [Q3 + 1.5*IQR] and Z > 2.5.")
    ]
    for x, num, meaning, evidence in findings:
        add_card(s8, x, 1.3, 2.85, 3.6)
        
        # Number y 1.5 h 1 Cambria 40 bold red, middle
        tb_num = s8.shapes.add_textbox(Inches(x + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = tb_num.text_frame
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_n = tf_n.paragraphs[0]
        p_n.alignment = PP_ALIGN.CENTER
        p_n.text = num
        p_n.font.name = "Cambria"
        p_n.font.size = Pt(36)
        p_n.font.bold = True
        p_n.font.color.rgb = C_RED

        # Meaning y 2.6 h .9, 14 bold black, top
        tb_m = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True
        p_m = tf_m.paragraphs[0]
        p_m.text = meaning
        p_m.font.name = "Calibri"
        p_m.font.size = Pt(13)
        p_m.font.bold = True
        p_m.font.color.rgb = C_BLACK

        # Evidence y 3.6 h 1.1, 12 grey, top
        tb_e = s8.shapes.add_textbox(Inches(x + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = tb_e.text_frame
        tf_e.word_wrap = True
        p_e = tf_e.paragraphs[0]
        p_e.text = evidence
        p_e.font.name = "Calibri"
        p_e.font.size = Pt(11)
        p_e.font.color.rgb = C_GREY

    set_speaker_notes(s8, "Here are our three core empirical findings: First, Cardiology and Oncology exhibit the fastest medical cost inflation at 7.8% annually. Second, Length of Stay has a strong correlation of 0.68 with total patient billing. Third, automated outlier auditing flags exactly 2.6% of hospital claims, highlighting potential unbundled charges.")

    # =========================================================================
    # SLIDE 9: CONTENT "Testing and validation"
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s9, "Testing and validation", 9)

    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    tb_sc = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_sc = tb_sc.text_frame
    tf_sc.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_sc = tf_sc.paragraphs[0]
    p_sc.alignment = PP_ALIGN.CENTER
    p_sc.text = "16 / 16"
    p_sc.font.name = "Cambria"
    p_sc.font.size = Pt(44)
    p_sc.font.bold = True
    p_sc.font.color.rgb = C_RED

    # (.5,2.4,2.6,.9) 13 grey, middle
    tb_st = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_st = tb_st.text_frame
    tf_st.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_st = tf_st.paragraphs[0]
    p_st.alignment = PP_ALIGN.CENTER
    p_st.text = "test cases passed\npytest: 16 | manual: 10"
    p_st.font.name = "Calibri"
    p_st.font.size = Pt(13)
    p_st.font.color.rgb = C_GREY

    # TABLE at (3.4,1.35), column widths 1.1, 3.9, 1.1
    t_shape9 = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.5))
    tbl9 = t_shape9.table
    tbl9.columns[0].width = Inches(1.1)
    tbl9.columns[1].width = Inches(3.9)
    tbl9.columns[2].width = Inches(1.1)

    t9_headers = ["Test", "Checks", "Result"]
    for c_idx, h in enumerate(t9_headers):
        cell = tbl9.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_RED
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    t9_rows = [
        ("TC-T-02", "Primary key uniqueness on fact_hospital_billing", "100% Passed"),
        ("TC-T-04", "Cost additive balance (Room+Surg+Pharm+Lab == Total)", "100% Passed"),
        ("TC-T-06", "Outlier detection algorithm isolates extreme charges", "100% Passed")
    ]
    for r_idx, (test_id, check, res) in enumerate(t9_rows, 1):
        tbl9.cell(r_idx, 0).text = test_id
        tbl9.cell(r_idx, 1).text = check
        tbl9.cell(r_idx, 2).text = res
        for c_idx in range(3):
            cell = tbl9.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Calibri"
            p.font.size = Pt(11)
            p.font.color.rgb = C_BLACK

    set_speaker_notes(s9, "Quality assurance was rigorously enforced using automated pytest suites. All 16 automated technical tests passed with 100% success rate in under 3 seconds. The table highlights our key assertions verifying primary key uniqueness, referential integrity, and mathematical charge balance across all hospital admissions.")

    # =========================================================================
    # SLIDE 10: CONTENT "Recommendation, limitations and next steps"
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_content_chrome(s10, "Recommendation, limitations and next steps", 10)

    recs = [
        (0.5, "R", "Recommendations", "Standardize clinical pathways for procedures exceeding 7 days LOS to reduce bed occupancy costs.\n\nAudit surgical implant suppliers in Cardiology to curb 7.8% annual inflation."),
        (3.58, "L", "Limitations", "Analysis is conducted on a synthetic offline relational dataset.\n\nDirect live electronic health record (EHR) bidirectional telemetry is deferred to future enterprise releases."),
        (6.65, "N", "Next Steps", "Incorporate machine learning models (Random Forest / XGBoost) for 30-day readmission risk prediction.\n\nIntegrate live HL7 FHIR messaging connectors for hospital EHR sync.")
    ]
    for x, badge, head, body in recs:
        add_badge_card(s10, x, 1.25, 2.85, 3.7, badge, head, body)

    set_speaker_notes(s10, "In conclusion, we recommend standardizing clinical pathways to curb unnecessary length of stay and capping surgical implant markups. A realistic limitation is that our capstone uses synthetic data without live EHR integration. As next steps, we plan to incorporate predictive machine learning models for hospital readmissions.")

    # =========================================================================
    # SLIDE 11: DARK LIVE DEMO & Q&A
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.background
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = C_BLACK
    add_header_dark(s11)

    # Title (.5,1.55,9,1.5)
    tb_t11 = s11.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t11 = tb_t11.text_frame
    tf_t11.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t11 = tf_t11.paragraphs[0]
    p_t11.text = "Live demo and Q&A"
    p_t11.font.name = "Cambria"
    p_t11.font.size = Pt(36)
    p_t11.font.bold = True
    p_t11.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7)
    tb_l11 = s11.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_l11 = tb_l11.text_frame
    tf_l11.word_wrap = True
    
    end_lines = [
        "App: https://2500320140033-treatment-cost-analysis.streamlit.app",
        "Code: https://github.com/arohityagi/033-treatment-cost-analysis",
        "Thank you"
    ]
    for idx, el in enumerate(end_lines):
        p = tf_l11.paragraphs[0] if idx == 0 else tf_l11.add_paragraph()
        p.text = el
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.color.rgb = C_CARD if idx < 2 else C_PINK
        if idx > 0:
            p.space_before = Pt(6)

    set_speaker_notes(s11, "I will now open the live Streamlit dashboard for interactive demonstration and welcome any questions from the examination committee. Thank you very much for your time and guidance.")

    # Save presentation
    target_path = "docs/033_Presentation.pptx"
    prs.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_presentation()
