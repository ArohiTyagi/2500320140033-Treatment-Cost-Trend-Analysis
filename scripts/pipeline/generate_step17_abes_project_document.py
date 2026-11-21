import docx
from pathlib import Path
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_element(name):
    return OxmlElement(name)

def set_cell_margins_and_border(cell, top=60, bottom=60, left=90, right=90, border_color="BFBFBF"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)
    
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def set_cell_fill(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def format_abes_table(table, col_widths_dxa, has_header=True):
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblPr = table._tbl.tblPr
    
    # Set fixed table layout and total width 9638 dxa
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="9638" w:type="dxa"/>')
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(tblW)
    tblPr.append(tblLayout)

    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        is_hdr = (i == 0 and has_header)
        if is_hdr:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            
        for j, cell in enumerate(row.cells):
            if j < len(col_widths_dxa):
                w_dxa = col_widths_dxa[j]
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>')
                tcPr.append(tcW)
                
            set_cell_margins_and_border(cell, top=60, bottom=60, left=90, right=90, border_color="BFBFBF")
            
            if is_hdr:
                set_cell_fill(cell, "A51C30")
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Calibri"
                        r.font.size = Pt(9)
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Calibri"
                        r.font.size = Pt(9)
                        r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)

def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    return p

def add_label(doc, text, color_hex="151514", before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(color_hex)
    return p

def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(0.375) # ~540 dxa
    for r in p.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    # in case run wasn't created with text
    if not p.runs:
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    else:
        p.runs[0].text = text
    return p

def add_code(doc, code_text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run(code_text)
    r.font.name = "Consolas"
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
    return p

def build_abes_document():
    doc = docx.Document()
    
    # Page Setup: A4, 2 cm margins, header/footer 1.25 cm
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(2.0)
        s.right_margin = Cm(2.0)
        s.header_distance = Cm(1.25)
        s.footer_distance = Cm(1.25)
        
        # Header setup
        hdr = s.header
        p_hdr = hdr.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_hdr.paragraph_format.space_after = Pt(0)
        r_hdr = p_hdr.add_run("ABES Engineering College | Project Document | P106")
        r_hdr.font.name = "Calibri"
        r_hdr.font.size = Pt(8)
        r_hdr.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

        # Footer setup
        ftr = s.footer
        p_ftr = ftr.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ftr.paragraph_format.space_after = Pt(0)
        r_ftr = p_ftr.add_run("Arohi Tyagi | 2500320140033     Page ")
        r_ftr.font.name = "Calibri"
        r_ftr.font.size = Pt(8)
        r_ftr.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
        
        # Add Page field
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        p_ftr._p.append(fldSimple)

    # Styles Setup
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    h1 = doc.styles['Heading 1']
    h1.font.name = 'Georgia'
    h1.font.size = Pt(15)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(7)
    h1.paragraph_format.keep_with_next = True

    h2 = doc.styles['Heading 2']
    h2.font.name = 'Georgia'
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0x15, 0x15, 0x14)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(5)
    h2.paragraph_format.keep_with_next = True

    # =========================================================================
    # PAGE 1: COVER
    # =========================================================================
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(5)
    r1 = p1.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(8)
    p_logo.paragraph_format.space_after = Pt(8)
    r_logo = p_logo.add_run()
    logo_file = Path(__file__).resolve().parents[2] / "docs" / "abes_logo.png"
    if not logo_file.exists():
        logo_file = Path(__file__).resolve().parents[2] / "docs" / "abes_logo.jpg"
    if logo_file.exists():
        r_logo.add_picture(str(logo_file), width=Inches(2.4))
    else:
        r_logo.text = "[ABES Engineering College]"

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(70)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("PROJECT DOCUMENT")
    r_title.font.name = "Georgia"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x15, 0x15, 0x14)

    p_proj = doc.add_paragraph()
    p_proj.paragraph_format.space_before = Pt(0)
    p_proj.paragraph_format.space_after = Pt(12)
    r_proj = p_proj.add_run("P106 - Treatment Cost Trend Analysis")
    r_proj.font.name = "Georgia"
    r_proj.font.size = Pt(18)
    r_proj.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)
    # Add bottom border to paragraph
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="A51C30"/></w:pBdr>')
    p_proj._p.get_or_add_pPr().append(pBdr)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(24)
    p_desc.paragraph_format.space_after = Pt(5)
    r_desc = p_desc.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
    r_desc.font.name = "Calibri"
    r_desc.font.size = Pt(10)
    r_desc.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

    # Cover Details Table: Field 2900 | Details 6738 dxa
    cover_table = doc.add_table(rows=8, cols=2)
    cover_widths = [2900, 6738]
    cover_data = [
        ("Project ID & Title", "P106 - Treatment Cost Trend Analysis"),
        ("Domain", "Data Analyst"),
        ("Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("Learner(s) & Roll No(s)", "Arohi Tyagi | Roll No: 2500320140033"),
        ("Mentor", "Mr. Dinesh"),
        ("GitHub repository", "https://github.com/ArohiTyagi/2500320140033-Treatment-Cost-Trend-Analysis"),
        ("Live app (Streamlit)", "https://2500320140033-treatment-cost-trend-analysis-bafa6sbr9kx7s8zfmu.streamlit.app/"),
        ("Version & date", "v1.0 - 20/11/2026")
    ]
    for idx, (fld, val) in enumerate(cover_data):
        row = cover_table.rows[idx]
        row.cells[0].paragraphs[0].add_run(fld).bold = True
        row.cells[1].paragraphs[0].add_run(val)
    format_abes_table(cover_table, cover_widths, has_header=False)

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: TABLE OF CONTENTS (Word TOC Field)
    # =========================================================================
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(0)
    p_toc.paragraph_format.space_after = Pt(0)
    # Insert native Word TOC field
    fldSimpleTOC = parse_xml(r'<w:fldSimple %s w:instr="TOC \h \o &quot;1-1&quot;"/>' % nsdecls('w'))
    p_toc._p.append(fldSimpleTOC)
    
    # Set updateFields on open so Word automatically refreshes page numbers
    settingsPr = doc.settings.element
    updateFields = parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls('w'))
    settingsPr.append(updateFields)

    doc.add_page_break()

    # =========================================================================
    # SECTION 1: PROJECT DESCRIPTION
    # =========================================================================
    doc.add_heading("1. Project Description", level=1)
    t1 = doc.add_table(rows=7, cols=2)
    t1_widths = [2400, 7238]
    t1_data = [
        ("Overview", "An interactive Data Analyst and Hospital Billing Analytics dashboard that tracks longitudinal inpatient treatment cost trends, detects anomalous charges, evaluates Length of Stay dynamics, and projects future expenditures."),
        ("Problem statement", "Healthcare administrators, clinical heads, and insurance adjudicators lack an integrated, data-driven analytical tool to track treatment cost inflation, evaluate departmental expenditure variance, and detect anomalous billing patterns, leading to budget unpredictability, inefficient financial resource allocation, and elevated out-of-pocket burdens on patients."),
        ("Intended users", "Hospital Chief Financial Officers (weekly/monthly budget variance), Clinical Department Heads (bi-weekly procedure cost audits), and Health Insurance TPA Claims Managers (daily/weekly anomaly review)."),
        ("In scope", "Synthetic longitudinal multi-year hospitalization billing data (SQLite), exploratory data analysis, MoM cost inflation, procedure IQR variance, outlier detection (Tukey's fences), rolling 90-day forecasting, and Tailwind-styled Streamlit UI."),
        ("Out of scope", "Live HL7/FHIR EHR hospital server synchronization, real-time payment gateway transaction settlements, and diagnostic computer vision disease classification."),
        ("Assumptions", "Synthetic data realistically models private/corporate Indian hospital billing dynamics; end-users utilize standard modern desktop browsers; monetary figures standardized in INR."),
        ("Dataset", "Synthetic multi-year relational dataset spanning 2023-01 to 2026-10 (350 inpatient fact admissions, 4 dimensions, 32-column analytical mart view) generated with fixed random seed 42.")
    ]
    for idx, (item, desc) in enumerate(t1_data):
        row = t1.rows[idx]
        row.cells[0].paragraphs[0].add_run(item).bold = True
        row.cells[1].paragraphs[0].add_run(desc)
    format_abes_table(t1, t1_widths, has_header=False)

    # =========================================================================
    # SECTION 2: HOW MIGHT WE (HMW) STATEMENTS
    # =========================================================================
    doc.add_heading("2. How Might We (HMW) Statements", level=1)
    add_note(doc, "Format: How might we <action> for <user> so that <outcome>?")
    
    t2 = doc.add_table(rows=11, cols=2)
    t2_widths = [1300, 8338]
    t2.rows[0].cells[0].paragraphs[0].add_run("ID")
    t2.rows[0].cells[1].paragraphs[0].add_run("HMW statement")
    
    t2_data = [
        ("HMW-01", "How might we empower hospital financial controllers to track month-over-month treatment cost trends across departments so that budget overruns are caught proactively?"),
        ("HMW-02", "How might we help clinical heads identify departments and diagnosis groups experiencing the highest cost inflation so that clinical pathway reviews can be prioritized?"),
        ("HMW-03", "How might we enable hospital administrators to examine how Length of Stay impacts total patient costs across demographic groups so that bed occupancy efficiency is maximized?"),
        ("HMW-04", "How might we allow billing leads and insurers to measure the out-of-pocket payment burden on patients versus insurance reimbursements so that financial hardship is transparent?"),
        ("HMW-05", "How might we help medical departments isolate procedures with excessive cost variance so that standardization protocols can be established?"),
        ("HMW-06", "How might we equip claims managers with automated outlier detection to instantly flag abnormal billing amounts and avoid audit penalties?"),
        ("HMW-07", "How might we help executive leadership project next-quarter treatment cost expenditures based on historical trajectories so that capital allocation is predictable?"),
        ("HMW-08", "How might we provide finance teams with an atomic component breakdown (room, surgery, pharmacy, lab) of inpatient bills to pinpoint the primary drivers of cost escalation?"),
        ("HMW-09", "How might we help operations managers assess the cost premium associated with Emergency versus Elective admissions to better plan emergency room staffing and resources?"),
        ("HMW-10", "How might we benchmark average daily treatment costs across hospital branches and facilities so that organizational operational parity is achieved?")
    ]
    for idx, (hid, stmt) in enumerate(t2_data, 1):
        row = t2.rows[idx]
        row.cells[0].paragraphs[0].add_run(hid)
        row.cells[1].paragraphs[0].add_run(stmt)
    format_abes_table(t2, t2_widths, has_header=True)

    # =========================================================================
    # SECTION 3: STAKEHOLDERS
    # =========================================================================
    doc.add_heading("3. Stakeholders", level=1)
    t3 = doc.add_table(rows=7, cols=5)
    t3_widths = [1000, 2400, 1900, 3138, 1200]
    t3.rows[0].cells[0].paragraphs[0].add_run("ID")
    t3.rows[0].cells[1].paragraphs[0].add_run("Stakeholder")
    t3.rows[0].cells[2].paragraphs[0].add_run("Role")
    t3.rows[0].cells[3].paragraphs[0].add_run("Interest")
    t3.rows[0].cells[4].paragraphs[0].add_run("Influence")
    
    t3_data = [
        ("SH-01", "Hospital CFO / Finance Lead", "Executive Sponsor", "Cost containment, margin optimization, budget forecasting", "High"),
        ("SH-02", "Clinical Department Heads", "Operational User", "Treatment cost efficacy, standard care protocols vs. cost", "High"),
        ("SH-03", "Medical Billing & Claims Head", "Primary User", "Billing integrity, outlier detection, claim dispute prevention", "Medium"),
        ("SH-04", "Health Insurance / TPA Analyst", "External Auditor", "Claim validity, actuarial cost benchmarking, tariff negotiation", "Medium"),
        ("SH-05", "Lead Data Analyst (Arohi Tyagi)", "Lead Developer", "Model accuracy, interactive UX, statistical rigor", "High"),
        ("SH-06", "Academic Mentor (Mr. Dinesh)", "Evaluator / Guide", "Technical completeness, academic rigor, ABES standards", "High")
    ]
    for idx, vals in enumerate(t3_data, 1):
        row = t3.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t3, t3_widths, has_header=True)

    # =========================================================================
    # SECTION 4: TRACEABILITY MATRIX
    # =========================================================================
    doc.add_heading("4. Traceability Matrix", level=1)
    add_note(doc, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")
    
    t4 = doc.add_table(rows=13, cols=9)
    t4_widths = [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040]
    t4_headers = ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"]
    for c_idx, h in enumerate(t4_headers):
        t4.rows[0].cells[c_idx].paragraphs[0].add_run(h)
        
    t4_data = [
        ("HMW-01", "BR-01", "Aggregate & track MoM treatment cost trends", "SH-01", "FEAT-01", "Must", "US-01", "TC-F-01", "UI-02"),
        ("HMW-02", "BR-02", "Compute annual compound inflation rate (CAGR)", "SH-02", "FEAT-02", "Must", "US-02", "TC-F-02", "UI-02"),
        ("HMW-03", "BR-03", "Evaluate Length of Stay (LOS) vs cost correlation", "SH-02", "FEAT-03", "Must", "US-03", "TC-F-03", "UI-04"),
        ("HMW-04", "BR-04", "Decompose insurance reimbursement vs out-of-pocket", "SH-04", "FEAT-04", "Must", "US-04", "TC-F-04", "UI-01"),
        ("HMW-05", "BR-06", "Analyze procedure-level IQR cost variance", "SH-02", "FEAT-05", "Must", "US-05", "TC-F-05", "UI-04"),
        ("HMW-06", "BR-05", "Automated statistical outlier detection (Tukey IQR)", "SH-03", "FEAT-06", "Must", "US-06", "TC-F-06", "UI-03"),
        ("HMW-07", "BR-07", "Project rolling 90-day cost expenditures", "SH-01", "FEAT-07", "Should", "US-07", "TC-F-07", "UI-05"),
        ("HMW-08", "BR-08", "Decompose patient bills into constituent charges", "SH-01", "FEAT-08", "Must", "US-08", "TC-F-08", "UI-01"),
        ("HMW-09", "BR-09", "Compare Emergency vs Elective admission cost", "SH-02", "FEAT-09", "Should", "US-09", "TC-F-09", "UI-01"),
        ("HMW-10", "BR-10", "Benchmark daily cost and stay across facilities", "SH-01", "FEAT-10", "Should", "US-10", "TC-F-10", "UI-01"),
        ("HMW-01", "BR-11", "Simulate revenue impact under fee cap adjustments", "SH-01", "FEAT-11", "Could", "US-11", "TC-F-07", "UI-05"),
        ("HMW-06", "BR-12", "Real-time EHR HL7 API bidirectional sync", "SH-03", "FEAT-12", "Won't", "N/A", "N/A", "N/A")
    ]
    for idx, vals in enumerate(t4_data, 1):
        row = t4.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t4, t4_widths, has_header=True)

    # =========================================================================
    # SECTION 5: USER STORIES
    # =========================================================================
    doc.add_heading("5. User Stories", level=1)
    t5 = doc.add_table(rows=7, cols=4)
    t5_widths = [900, 3400, 4138, 1200]
    t5_headers = ["US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"]
    for c_idx, h in enumerate(t5_headers):
        t5.rows[0].cells[c_idx].paragraphs[0].add_run(h)
        
    t5_data = [
        ("US-01", "As a CFO, I want to track monthly mean and median treatment costs so that budget overruns are detected.", "Given billing data is loaded, When viewed on Trends page, Then line chart renders mean & median with MoM growth.", "Must"),
        ("US-02", "As a Clinical Head, I want to view departmental CAGR inflation so that high-cost specialties are prioritized.", "Given multi-year records, When department is selected, Then annualized CAGR is displayed with trend slope.", "Must"),
        ("US-03", "As a Medical Director, I want to evaluate LOS vs cost correlation so that bed utilization is optimized.", "Given patient stays, When Procedure page is opened, Then scatter plot shows regression slope with Pearson r card.", "Must"),
        ("US-04", "As a TPA Claims Lead, I want to view insurance vs out-of-pocket split so that patient burden is visible.", "Given billing ledger, When viewing settlement charts, Then donut chart shows percentage covered vs out-of-pocket.", "Must"),
        ("US-05", "As a Clinical Head, I want procedure IQR boxplots so that care protocol variance can be minimized.", "Given surgical records, When procedure boxplot renders, Then median, Q1, Q3, and Tukey whiskers are shown.", "Must"),
        ("US-06", "As a Billing Auditor, I want automated outlier flags so that suspicious invoices are reviewed before submission.", "Given invoice ledger, When Outlier page opens, Then claims with Z > 2.5 or exceeding 1.5*IQR are flagged in table.", "Must")
    ]
    for idx, vals in enumerate(t5_data, 1):
        row = t5.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t5, t5_widths, has_header=True)

    # =========================================================================
    # SECTION 6: BUSINESS REQUIREMENTS DOCUMENT (BRD)
    # =========================================================================
    doc.add_heading("6. Business Requirements Document (BRD)", level=1)
    
    doc.add_heading("6.1 Business objectives", level=2)
    add_bullet(doc, "Deliver transparent, interactive longitudinal visibility into inpatient treatment expenditures across 6 clinical departments.")
    add_bullet(doc, "Automate statistical anomaly detection to reduce billing auditing turnaround time by 60% and mitigate claim rejections.")
    add_bullet(doc, "Empower hospital executive leadership with rolling 90-day predictive expenditure forecasts for working capital planning.")
    add_bullet(doc, "Quantify patient financial liability by measuring the exact proportion of out-of-pocket payments across insurance categories.")

    doc.add_heading("6.2 Success KPIs", level=2)
    t6_kpi = doc.add_table(rows=6, cols=5)
    t6_kpi_widths = [1900, 2100, 2100, 1700, 1838]
    t6_kpi_headers = ["KPI", "Numerator", "Denominator", "Period", "Target / note"]
    for c_idx, h in enumerate(t6_kpi_headers):
        t6_kpi.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t6_kpi_data = [
        ("Average Treatment Cost", "Sum of Total Billed Cost", "Total Patient Discharges", "Monthly", "< ₹1,35,000 / stay"),
        ("Outlier Detection Rate", "Flagged High-Cost Claims", "Total Hospital Invoices", "Monthly", "2.0% - 4.5% of cohort"),
        ("Out-of-Pocket Share %", "Sum of Out-of-Pocket Paid", "Total Billed Healthcare Cost", "Quarterly", "< 28.0% of total"),
        ("Departmental CAGR", "Current Period Dept Mean Cost", "Base Period Dept Mean Cost", "Annual", "6.0% - 8.5% p.a."),
        ("Average Length of Stay", "Sum of Inpatient Bed Days", "Total Patient Admissions", "Monthly", "3.8 - 5.2 Days")
    ]
    for idx, vals in enumerate(t6_kpi_data, 1):
        row = t6_kpi.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t6_kpi, t6_kpi_widths, has_header=True)

    doc.add_heading("6.3 Business requirements", level=2)
    t6_br = doc.add_table(rows=7, cols=4)
    t6_br_widths = [1100, 5338, 1300, 1900]
    t6_br_headers = ["BR-ID", "Requirement", "Priority", "Source (stakeholder)"]
    for c_idx, h in enumerate(t6_br_headers):
        t6_br.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t6_br_data = [
        ("BR-01", "The business needs to aggregate and track month-over-month inpatient billing totals, means, and medians so that leadership can spot systemic financial drift.", "Must", "SH-01 (CFO)"),
        ("BR-02", "The business needs to compute annual compound inflation rates (CAGR) by clinical department so that high-inflation specialties can be identified.", "Must", "SH-02 (Clinical)"),
        ("BR-03", "The business needs to evaluate the relationship between Length of Stay (LOS) and total billed expenditure across patient demographics so that clinical bed turnover is optimized.", "Must", "SH-02 (Clinical)"),
        ("BR-04", "The business needs to quantify the financial split between insurance reimbursements and patient out-of-pocket costs so that affordability and insurance adequacy can be assessed.", "Must", "SH-04 (TPA)"),
        ("BR-05", "The business needs to automatically identify statistical billing outliers beyond normal distribution thresholds so that suspicious claims can be audited before billing finalization.", "Must", "SH-03 (Billing)"),
        ("BR-06", "The business needs to decompose patient bills into distinct cost categories (room, surgery, medicine, lab) so that the principal operational cost drivers are visible.", "Must", "SH-01 (CFO)")
    ]
    for idx, vals in enumerate(t6_br_data, 1):
        row = t6_br.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t6_br, t6_br_widths, has_header=True)

    doc.add_heading("6.4 Constraints and out of scope", level=2)
    add_bullet(doc, "Memory resource constraint: The application must operate within 1 GB RAM on Streamlit Community Cloud.")
    add_bullet(doc, "Standalone offline scope: Excludes live bidirectional HL7/FHIR EHR hospital database integration.")
    add_bullet(doc, "Ethical compliance: Strictly zero real personal patient health data; 100% synthetic relational dataset.")

    # =========================================================================
    # SECTION 7: TECHNICAL REQUIREMENTS DOCUMENT (TRD)
    # =========================================================================
    doc.add_heading("7. Technical Requirements Document (TRD)", level=1)
    
    doc.add_heading("7.1 Technology stack (all free)", level=2)
    t7_tech = doc.add_table(rows=9, cols=3)
    t7_tech_widths = [2000, 3800, 3838]
    t7_tech_headers = ["Layer", "Tool", "Why"]
    for c_idx, h in enumerate(t7_tech_headers):
        t7_tech.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t7_tech_data = [
        ("Language", "Python 3.12 / 3.15", "Industry standard for data engineering, statistical modeling, and analytical rapid prototyping."),
        ("Data processing", "Pandas 2.2+ & NumPy 1.26+", "High-performance vectorized dataframe manipulation, aggregation, and mathematical operations."),
        ("Storage", "SQLite 3.x (data/app.db)", "Serverless, zero-configuration relational database enforcing ACID compliance and foreign keys."),
        ("Charts", "Plotly Express 5.24+", "Interactive WebGL/HTML5 visualizations featuring zoom, pan, hover tooltips, and responsive layout."),
        ("App / UI", "Streamlit 1.38+ with Tailwind CDN", "Rapid pure-Python multipage web framework styled with modern Tailwind CSS card components."),
        ("Testing", "Pytest 8.3+", "Automated test runner supporting parameterized unit, functional, and schema integrity validation."),
        ("Version control", "Git & GitHub", "Distributed version management, branch tracking, and automated deployment hook to cloud."),
        ("Editor", "VS Code / Antigravity IDE", "Advanced agentic development environment with integrated debugging and terminal control.")
    ]
    for idx, vals in enumerate(t7_tech_data, 1):
        row = t7_tech.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t7_tech, t7_tech_widths, has_header=True)

    doc.add_heading("7.2 Functional requirements", level=2)
    t7_fr = doc.add_table(rows=7, cols=3)
    t7_fr_widths = [1100, 6338, 2200]
    t7_fr_headers = ["FR-ID", "Requirement", "Linked BR / US"]
    for c_idx, h in enumerate(t7_fr_headers):
        t7_fr.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t7_fr_data = [
        ("FR-01", "Query SQLite database and cache analysis view in session memory.", "BR-01 / US-01"),
        ("FR-02", "Compute monthly mean, median, and MoM growth rate of billed costs.", "BR-01 / US-01"),
        ("FR-03", "Calculate Compound Annual Growth Rate (CAGR) per clinical department.", "BR-02 / US-02"),
        ("FR-04", "Calculate Pearson correlation coefficient between LOS and total cost.", "BR-03 / US-03"),
        ("FR-05", "Compute insurance coverage ratio and out-of-pocket patient percentage.", "BR-04 / US-04"),
        ("FR-06", "Filter cost outliers using Tukey's fences (1.5 * IQR) and Z-scores.", "BR-05 / US-06")
    ]
    for idx, vals in enumerate(t7_fr_data, 1):
        row = t7_fr.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t7_fr, t7_fr_widths, has_header=True)

    doc.add_heading("7.3 Non-functional requirements", level=2)
    t7_nfr = doc.add_table(rows=5, cols=2)
    t7_nfr_widths = [1100, 8538]
    t7_nfr_headers = ["NFR-ID", "Requirement"]
    for c_idx, h in enumerate(t7_nfr_headers):
        t7_nfr.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t7_nfr_data = [
        ("NFR-01", "Response Time: Dashboard pages and Plotly visualizations must render within 1.5 seconds on a dataset of 3,000+ records."),
        ("NFR-02", "Reliability: Zero runtime crashes or unhandled null exceptions; 100% passing automated test suite in pytest."),
        ("NFR-03", "Portability: Lightweight standalone repository runnable locally on any OS with pip install -r requirements.txt."),
        ("NFR-04", "Security & Privacy: Contains 100% synthetic non-PII healthcare billing data; read-only database analytical access.")
    ]
    for idx, vals in enumerate(t7_nfr_data, 1):
        row = t7_nfr.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t7_nfr, t7_nfr_widths, has_header=True)

    doc.add_heading("7.4 Architecture and deployment", level=2)
    add_label(doc, "Raw data (data/raw) -> src/generate_data.py -> SQLite (data/app.db) -> Streamlit app (app.py) -> user", color_hex="A51C30", before=0)
    
    t7_dep = doc.add_table(rows=4, cols=3)
    t7_dep_widths = [2200, 4838, 2600]
    t7_dep_headers = ["Stage", "Command / action", "Done when"]
    for c_idx, h in enumerate(t7_dep_headers):
        t7_dep.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t7_dep_data = [
        ("1. Local", "streamlit run app.py && pytest tests/ -v", "App renders locally on localhost:8501 and all 16 tests pass."),
        ("2. GitHub", "git add . && git commit -m 'Release v1.0' && git push", "Repository updated with main branch, requirements.txt, and app.db."),
        ("3. Streamlit Cloud", "share.streamlit.io -> Deploy from repo '2500320140033-treatment-cost-trend-analysis'", "Live public HTTPS URL accessible with 0 build errors.")
    ]
    for idx, vals in enumerate(t7_dep_data, 1):
        row = t7_dep.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t7_dep, t7_dep_widths, has_header=True)

    # =========================================================================
    # SECTION 8: DATABASE DESIGN
    # =========================================================================
    doc.add_heading("8. Database Design", level=1)
    add_note(doc, "Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")

    t8 = doc.add_table(rows=11, cols=5)
    t8_widths = [1700, 2300, 2100, 1400, 2138]
    t8_headers = ["Table / sheet", "Column", "Type", "Key", "Description"]
    for c_idx, h in enumerate(t8_headers):
        t8.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t8_data = [
        ("dim_departments", "dept_id", "INTEGER", "PK", "Unique department identifier"),
        ("dim_departments", "dept_name", "TEXT", "None", "Cardiology, Orthopedics, Oncology, etc."),
        ("dim_procedures", "procedure_id", "INTEGER", "PK", "Unique procedure identifier"),
        ("dim_procedures", "dept_id", "INTEGER", "FK", "References dim_departments(dept_id)"),
        ("dim_facilities", "facility_id", "INTEGER", "PK", "Unique hospital location identifier"),
        ("dim_patients", "patient_id", "INTEGER", "PK", "Unique registered patient identifier"),
        ("fact_hospital_billing", "billing_id", "INTEGER", "PK", "Unique inpatient billing invoice ID"),
        ("fact_hospital_billing", "patient_id", "INTEGER", "FK", "References dim_patients(patient_id)"),
        ("fact_hospital_billing", "procedure_id", "INTEGER", "FK", "References dim_procedures(procedure_id)"),
        ("fact_hospital_billing", "total_billed_cost", "REAL", "None", "Total billed hospitalization cost")
    ]
    for idx, vals in enumerate(t8_data, 1):
        row = t8.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t8, t8_widths, has_header=True)

    add_label(doc, "Relationship: dim_departments (1) -- (many) dim_procedures; dim_procedures (1) -- (many) fact_hospital_billing; dim_facilities (1) -- (many) fact_hospital_billing; dim_patients (1) -- (many) fact_hospital_billing.", before=5)
    
    add_label(doc, "SQLite DDL:", before=5)
    add_code(doc, 
        "CREATE TABLE fact_hospital_billing (\n"
        "    billing_id INTEGER PRIMARY KEY,\n"
        "    patient_id INTEGER NOT NULL REFERENCES dim_patients(patient_id),\n"
        "    procedure_id INTEGER NOT NULL REFERENCES dim_procedures(procedure_id),\n"
        "    facility_id INTEGER NOT NULL REFERENCES dim_facilities(facility_id),\n"
        "    admission_date TEXT NOT NULL,\n"
        "    discharge_date TEXT NOT NULL,\n"
        "    los_days INTEGER NOT NULL CHECK(los_days >= 1),\n"
        "    admission_type TEXT NOT NULL CHECK(admission_type IN ('Emergency', 'Elective', 'Urgent')),\n"
        "    room_charges REAL NOT NULL CHECK(room_charges >= 0),\n"
        "    surgery_charges REAL NOT NULL CHECK(surgery_charges >= 0),\n"
        "    pharmacy_charges REAL NOT NULL CHECK(pharmacy_charges >= 0),\n"
        "    lab_charges REAL NOT NULL CHECK(lab_charges >= 0),\n"
        "    total_billed_cost REAL NOT NULL CHECK(total_billed_cost > 0),\n"
        "    insurance_paid REAL NOT NULL CHECK(insurance_paid >= 0),\n"
        "    out_of_pocket_paid REAL NOT NULL CHECK(out_of_pocket_paid >= 0),\n"
        "    insurance_type TEXT,\n"
        "    is_outlier INTEGER NOT NULL DEFAULT 0 CHECK(is_outlier IN (0, 1))\n"
        ");"
    )

    # =========================================================================
    # SECTION 9: TEST CASES
    # =========================================================================
    doc.add_heading("9. Test Cases", level=1)
    t9 = doc.add_table(rows=8, cols=6)
    t9_widths = [900, 1500, 2400, 2638, 1100, 1100]
    t9_headers = ["TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail"]
    for c_idx, h in enumerate(t9_headers):
        t9.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t9_data = [
        ("TC-F-01", "BR-01 / US-01", "Filter by department on Overview page.", "Cards & charts update strictly to chosen department.", "Matched", "Pass"),
        ("TC-F-02", "BR-02 / US-02", "Navigate to Trends page; view CAGR chart.", "Departments ranked by CAGR with percentage values.", "Matched", "Pass"),
        ("TC-F-03", "BR-03 / US-03", "Open Procedure page; inspect LOS scatter.", "Scatter shows regression slope with Pearson r.", "Matched", "Pass"),
        ("TC-F-06", "BR-05 / US-06", "Open Outlier page; adjust Z-score slider.", "Table dynamically lists claims exceeding Z-score.", "Matched", "Pass"),
        ("TC-T-01", "BR-01 / US-12", "Execute test_database_file_exists.", "Assert data/app.db exists on disk.", "Found", "Pass"),
        ("TC-T-04", "BR-08 / US-08", "Execute test_cost_additive_balance.", "Assert Room+Surgery+Pharm+Lab == Total Cost.", "0 diff", "Pass"),
        ("TC-T-05", "BR-01 / US-01", "Execute test_kpi_computation.", "Accurate mean, median, total revenue calculation.", "Matched", "Pass")
    ]
    for idx, vals in enumerate(t9_data, 1):
        row = t9.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t9, t9_widths, has_header=True)

    # =========================================================================
    # SECTION 10: UI FRAMEWORK
    # =========================================================================
    doc.add_heading("10. UI Framework", level=1)
    add_note(doc, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")

    t10 = doc.add_table(rows=7, cols=6)
    t10_widths = [800, 1200, 1700, 2900, 1700, 1338]
    t10_headers = ["UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories"]
    for c_idx, h in enumerate(t10_headers):
        t10.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t10_data = [
        ("UI-01", "Overview", "Executive KPI summary & cost waterfall", "Tailwind Header, 4 KPI cards, Waterfall Chart, Branch Bar Chart", "Dept, Facility, Year, Acuity", "US-01, US-08, US-10"),
        ("UI-02", "Cost Trends", "Longitudinal monthly trend & CAGR inflation", "MoM Multi-line Chart, CAGR Ranking Bar, Data Table", "Specialty, Date Range", "US-01, US-02"),
        ("UI-03", "Outliers", "Tukey IQR & Z-score billing audit", "Z-score Scatter, Highlighted Table, CSV Export Button", "Z-threshold, IQR factor", "US-06, US-09"),
        ("UI-04", "Procedures", "Procedure IQR variance & LOS correlation", "Procedure Boxplots, LOS Scatter Regression, Correlation Cards", "Department, Acuity", "US-03, US-05"),
        ("UI-05", "Forecast", "Rolling quarterly projection & price caps", "Quarterly Line with Confidence Ribbon, Price Adjustment Slider", "Horizon, Fee Delta %", "US-07, US-11"),
        ("UI-06", "Explorer", "Relational database mart & quality check", "Searchable Grid, Entity Tabs, Schema Summary, CSV Download", "Search text, Entity Tabs", "US-12, US-13")
    ]
    for idx, vals in enumerate(t10_data, 1):
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
    dash_file = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.png"
    if not dash_file.exists():
        dash_file = Path(__file__).resolve().parents[2] / "docs" / "dashboard_screenshot.jpg"
    if dash_file.exists():
        r_di.add_picture(str(dash_file), width=Inches(6.5))

    doc.add_page_break()

    # =========================================================================
    # APPENDIX A & B
    # =========================================================================
    doc.add_heading("Appendix A - Prompt to generate this document", level=1)
    add_label(doc, "MY PREFERENCES", color_hex="A51C30")
    add_code(doc,
        "PROJECT ID & TITLE: P106 - Treatment Cost Trend Analysis\n"
        "DOMAIN: Data Analyst\n"
        "LEARNER(S) & ROLL NO(S): Arohi Tyagi | Roll No: 2500320140033\n"
        "MENTOR: Mr. Dinesh\n"
        "BATCH: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01\n"
        "INTENDED USERS: Hospital CFO, Clinical Department Heads, Health Insurance TPA Claims Managers\n"
        "PROBLEM IN MY WORDS: Healthcare costs escalate unpredictably due to medical inflation and unmonitored length of stay, while billing anomalies cause claim rejections.\n"
        "DATASET: Synthetic longitudinal hospitalization records (350 fact rows, 4 dimensions, 32 columns)\n"
        "KPIS I WANT: Average Cost, Median Cost, Departmental CAGR, Outlier Ratio, Out-of-Pocket Share, ALOS\n"
        "FEATURES I WANT: Executive KPI dashboard, Cost Component Waterfall, MoM Line Charts, Tukey Outlier Detector, LOS Regression, Rolling 90-day Forecast\n"
        "STORAGE: SQLite (data/app.db)\n"
        "TEAM SIZE & DEADLINE: 1 Developer (Arohi Tyagi) | November 2026"
    )

    add_label(doc, "INSTRUCTIONS FOR THE AI", color_hex="A51C30", before=10)
    add_code(doc,
        "# STEP 0 — Kickoff\n"
        "PROJECT: <treatment cost trend analysis> | DOMAIN: <Data analysis>\n"
        "ME: <arohi tyagi>, Roll <2500320140033>, MCA 2nd yr DA, ABES | MENTOR: <Mr. Dinesh>\n"
        "PRODUCT: Python + Streamlit app (EDA, statistics, Plotly charts, Tailwind UI), SQLite data, pytest tests, on Streamlit Community Cloud.\n"
        "FINAL SUBMISSION: Project Document (ABES Word template) + ABES slide deck.\n"
        "STEPS: 1 Project understanding 2 HMW 3 Needs analysis (BR) 4 Traceability 5 BRD 6 PRD 7 TBD 8 Database design 9 ER diagram 10 Dataset generation 11 User stories 12 UI framework 13 Test cases 14 Build plan ID 15 Build, test, fix 16 Deploy 17 Project Document 18 Presentation"
    )

    doc.add_heading("Appendix B - AI tools disclosure", level=1)
    t_app = doc.add_table(rows=4, cols=3)
    t_app_widths = [2600, 3500, 3538]
    t_app_headers = ["Tool", "Used for", "How the output was checked"]
    for c_idx, h in enumerate(t_app_headers):
        t_app.rows[0].cells[c_idx].paragraphs[0].add_run(h)
    t_app_data = [
        ("Antigravity AI (Gemini 3.8 Flash)", "Solution architecture design, python script drafting, and documentation synthesis.", "Reviewed manually against ABES rubrics, verified through automated pytest test suites."),
        ("Python-docx Engine", "Automated compilation of institutional Word document adhering to exact dxa table column specifications.", "Inspected rendered output in Microsoft Word; confirmed page breaks and margins."),
        ("Pytest Test Runner", "Verification of database schema integrity, mathematical formulas, and analytical pipelines.", "Ran automated command line test suite; confirmed 16 of 16 tests passing cleanly.")
    ]
    for idx, vals in enumerate(t_app_data, 1):
        row = t_app.rows[idx]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    format_abes_table(t_app, t_app_widths, has_header=True)

    target_path = "docs/P106_Project_Document.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_abes_document()
