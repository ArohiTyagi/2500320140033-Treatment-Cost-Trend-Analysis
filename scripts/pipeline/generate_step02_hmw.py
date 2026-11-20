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

def build_step02():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("TREATMENT COST TREND ANALYSIS")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Step 2: How Might We (HMW) Opportunity Framework")
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
        ("Academic Mentor", "Prof. Alpana Lodhi"),
        ("Deliverable Step", "STEP 2 — Design Thinking HMW Framework"),
        ("Linked Operational Questions", "OQ-01 through OQ-10 (100% Coverage)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Stakeholder Pain Points
    doc.add_heading("1. Stakeholder Pain Points Analysis", level=1)
    doc.add_paragraph(
        "Through design-thinking discovery across the hospital ecosystem, 15 core pain points were identified across the six primary stakeholders:"
    )

    pp_data = [
        ("SH-01: Hospital CFO / Finance Lead", [
            "Lacks visibility into monthly medical inflation and departmental expenditure drift until post-period accounting closes.",
            "Inability to forecast upcoming quarterly inpatient billing requirements accurately.",
            "Struggles to determine whether revenue variance stems from increased patient volume or rising unit procedure costs."
        ]),
        ("SH-02: Clinical Department Heads", [
            "No clear data linking Length of Stay (LOS) to unwarranted cost escalations per specialty.",
            "Difficulty comparing treatment cost variance across procedures to enforce standard clinical care pathways.",
            "Limited visibility into cost disparities between Emergency and Elective admissions."
        ]),
        ("SH-03: Medical Billing & Claims Head", [
            "Severe manual auditing workload in detecting billing anomalies and rogue invoice spikes before insurance submission.",
            "Frequent claim dispute rejections due to unbundled or out-of-bracket diagnostic and room charges.",
            "Absence of an automated statistical outlier filter to flag suspicious patient accounts."
        ]),
        ("SH-04: Health Insurance & TPA Analyst", [
            "Lack of transparency into the patient out-of-pocket ratio vs. covered charges across insurance policy categories.",
            "Difficulty benchmarking hospital billing rates against regional median treatment costs."
        ]),
        ("SH-05: Lead Data Analyst (Arohi Tyagi)", [
            "Need to ensure analytical algorithms remain robust against noisy, skewed healthcare distributions.",
            "Need to deliver sub-second dashboard rendering on Streamlit Cloud without complex server management."
        ]),
        ("SH-06: Academic Mentor (Prof. Alpana Lodhi)", [
            "Ensuring rigorous traceability between business pain points, quantitative models, and ABES curriculum requirements.",
            "Verifying that all operational questions are backed by reproducible data pipelines and automated tests."
        ])
    ]

    for sh, points in pp_data:
        p_sh = doc.add_paragraph()
        r_sh = p_sh.add_run(sh)
        r_sh.bold = True
        r_sh.font.color.rgb = RGBColor(44, 62, 80)
        for pt in points:
            p_pt = doc.add_paragraph(style='List Bullet')
            p_pt.add_run(pt)

    doc.add_paragraph()

    # Section 2: HMW Statements Table
    doc.add_heading("2. How Might We (HMW) Opportunity Statements", level=1)
    doc.add_paragraph(
        "Pain points were transformed into 10 precisely scoped 'How might we' opportunity statements designed to guide solution development without presuming technical architecture."
    )

    hmw_table = doc.add_table(rows=11, cols=5)
    hmw_widths = [Inches(0.8), Inches(2.4), Inches(1.3), Inches(1.2), Inches(0.8)]
    hmw_table.rows[0].cells[0].paragraphs[0].add_run("HMW ID")
    hmw_table.rows[0].cells[1].paragraphs[0].add_run("How Might We Statement")
    hmw_table.rows[0].cells[2].paragraphs[0].add_run("Stakeholder")
    hmw_table.rows[0].cells[3].paragraphs[0].add_run("Source Pain Point")
    hmw_table.rows[0].cells[4].paragraphs[0].add_run("Linked OQ")

    hmw_rows = [
        ("HMW-01", "How might we empower hospital financial controllers to track month-over-month treatment cost trends across departments so that budget overruns are caught proactively?", "SH-01 (CFO)", "Delayed visibility into departmental drift", "OQ-01"),
        ("HMW-02", "How might we help clinical heads identify departments and diagnosis groups experiencing the highest cost inflation so that clinical pathway reviews can be prioritized?", "SH-02 (Clinical)", "Unmonitored procedure inflation", "OQ-02"),
        ("HMW-03", "How might we enable hospital administrators to examine how Length of Stay impacts total patient costs across demographic groups so that bed occupancy efficiency is maximized?", "SH-02 (Clinical)", "No clear data linking LOS to cost", "OQ-03"),
        ("HMW-04", "How might we allow billing leads and insurers to measure the out-of-pocket payment burden on patients versus insurance reimbursements so that financial hardship is transparent?", "SH-04 (TPA / SH-03)", "Lack of transparency in out-of-pocket split", "OQ-04"),
        ("HMW-05", "How might we help medical departments isolate procedures with excessive cost variance so that standardization protocols can be established?", "SH-02 (Clinical)", "Cost disparities across surgical procedures", "OQ-05"),
        ("HMW-06", "How might we equip claims managers with automated outlier detection to instantly flag abnormal billing amounts and avoid audit penalties?", "SH-03 (Billing)", "Manual auditing workload & billing spikes", "OQ-06"),
        ("HMW-07", "How might we help executive leadership project next-quarter treatment cost expenditures based on historical trajectories so that capital allocation is predictable?", "SH-01 (CFO)", "Inability to forecast inpatient budget", "OQ-07"),
        ("HMW-08", "How might we provide finance teams with an atomic component breakdown (room, surgery, pharmacy, lab) of inpatient bills to pinpoint the primary drivers of cost escalation?", "SH-01 (CFO)", "Unbundled charge obscurity", "OQ-08"),
        ("HMW-09", "How might we help operations managers assess the cost premium associated with Emergency versus Elective admissions to better plan emergency room staffing and resources?", "SH-02 (Clinical)", "Emergency vs elective cost disparity", "OQ-09"),
        ("HMW-10", "How might we benchmark average daily treatment costs across hospital branches and facilities so that organizational operational parity is achieved?", "SH-01 (CFO)", "Cross-facility cost disparity", "OQ-10")
    ]

    for idx, vals in enumerate(hmw_rows):
        row = hmw_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(hmw_table, hmw_widths)

    doc.add_paragraph()

    # Section 3: Priority HMWs
    doc.add_heading("3. Priority HMW Selection & Strategic Rationale", level=1)
    doc.add_paragraph(
        "To maximize analytical value for the Minimum Viable Product (MVP), the top five HMW statements were selected as PRIORITY based on business impact and feasibility:"
    )

    priorities = [
        ("HMW-01 [PRIORITY 1 - Trend Tracking]", "Core foundation of project; provides baseline longitudinal cost visibility across all clinical specialties."),
        ("HMW-06 [PRIORITY 2 - Outlier Detection]", "Highest direct ROI for billing operations by automatically mitigating revenue leakage, fraud, and claim disputes."),
        ("HMW-03 [PRIORITY 3 - LOS & Demographics]", "Crucial clinical metric correlating bed utilization with cost escalation to optimize hospital bed turnover."),
        ("HMW-04 [PRIORITY 4 - Insurance vs. Out-of-Pocket]", "Vital for healthcare affordability and TPA policy negotiations, providing multi-stakeholder social and economic value."),
        ("HMW-07 [PRIORITY 5 - Predictive Cost Forecasting]", "Transforms reactive reporting into proactive financial planning for hospital CFOs and executive leadership.")
    ]

    for title, rationale in priorities:
        p_pr = doc.add_paragraph(style='List Bullet')
        r_t = p_pr.add_run(f"{title}: ")
        r_t.bold = True
        p_pr.add_run(rationale)

    doc.add_paragraph()

    # Section 4: Coverage Verification
    doc.add_heading("4. Operational Question Coverage Verification", level=1)
    p_cov = doc.add_paragraph()
    r_cov = p_cov.add_run("Coverage Check Status: ")
    r_cov.bold = True
    p_cov.add_run(
        "PASSED (10/10). Every Operational Question (OQ-01 through OQ-10) is directly addressed by at least one dedicated HMW statement. "
        "Zero operational gaps exist between Step 1 discovery and Step 2 opportunity framing."
    )

    target_path = "docs/02_HMW_Statements.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step02()
