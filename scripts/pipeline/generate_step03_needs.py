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

def build_step03():
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
    r_sub = p_sub.add_run("Step 3: Business Needs Analysis & Requirements Specification")
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
        ("Author / Roll No", "Arohi Tyagi | Roll No: 2500320140033 | MCA 2nd Yr DA, ABES"),
        ("Academic Mentor", "Mr. Dinesh"),
        ("Deliverable Step", "STEP 3 — Needs Analysis & Business Requirements (BR)"),
        ("Prioritization Standard", "MoSCoW (Must: ~58%, Should: 25%, Could: 8%, Won't: 8%)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        meta_table.rows[idx].cells[0].paragraphs[0].add_run(lbl).bold = True
        meta_table.rows[idx].cells[1].paragraphs[0].add_run(val)
    style_table(meta_table, meta_widths, header_bg="2C3E50", alt_bg="F8F9FA")

    doc.add_paragraph()

    # Section 1: Business Requirements Definition
    doc.add_heading("1. Business Requirements Matrix", level=1)
    doc.add_paragraph(
        "Each requirement expresses WHAT the enterprise needs to accomplish and the corresponding business benefit, "
        "independent of implementation mechanisms or code structures."
    )

    br_table = doc.add_table(rows=13, cols=6)
    br_widths = [Inches(0.7), Inches(2.3), Inches(1.1), Inches(0.8), Inches(0.7), Inches(1.4)]
    br_table.rows[0].cells[0].paragraphs[0].add_run("BR ID")
    br_table.rows[0].cells[1].paragraphs[0].add_run("Business Requirement Statement")
    br_table.rows[0].cells[2].paragraphs[0].add_run("Stakeholder")
    br_table.rows[0].cells[3].paragraphs[0].add_run("Source")
    br_table.rows[0].cells[4].paragraphs[0].add_run("Priority")
    br_table.rows[0].cells[5].paragraphs[0].add_run("Measurement Method")

    br_data = [
        ("BR-01", "The business needs to aggregate and track month-over-month inpatient billing totals, means, and medians so that leadership can spot systemic financial drift.", "SH-01 (CFO)", "HMW-01", "Must", "Accuracy of calculated monthly mean/median against audited billing ledger."),
        ("BR-02", "The business needs to compute annual compound and annualized inflation rates by clinical department so that high-inflation specialties can be identified.", "SH-02 (Clinical)", "HMW-02", "Must", "Departmental CAGR matching empirical cost delta across sequential fiscal quarters."),
        ("BR-03", "The business needs to evaluate the relationship between Length of Stay (LOS) and total billed expenditure across patient demographics so that clinical bed turnover is optimized.", "SH-02 (Clinical)", "HMW-03", "Must", "Statistical correlation coefficient (r) and cost-per-day distribution validation."),
        ("BR-04", "The business needs to quantify the financial split between insurance reimbursements and patient out-of-pocket costs so that affordability and insurance adequacy can be assessed.", "SH-04 (TPA)", "HMW-04", "Must", "Out-of-pocket percentage matches (Patient_Pay / Total_Bill) * 100 exactly."),
        ("BR-05", "The business needs to automatically identify statistical billing outliers beyond normal distribution thresholds so that suspicious claims can be audited before billing finalization.", "SH-03 (Billing)", "HMW-06", "Must", "Outlier identification flag matching Tukey's IQR [Q1 - 1.5*IQR, Q3 + 1.5*IQR]."),
        ("BR-06", "The business needs to analyze inter-patient cost variability (IQR and standard deviation) for standard surgical procedures so that clinical protocol compliance can be enforced.", "SH-02 (Clinical)", "HMW-05", "Must", "Comparison of procedure IQR bounds against clinical standard guidelines."),
        ("BR-07", "The business needs to generate short-term rolling 90-day cost projections by clinical specialty so that finance teams can anticipate quarterly working capital needs.", "SH-01 (CFO)", "HMW-07", "Should", "Forecast Mean Absolute Percentage Error (MAPE) < 12% on test validation holdouts."),
        ("BR-08", "The business needs to decompose patient bills into distinct cost categories (room, surgery, medicine, lab) so that the principal operational cost drivers are visible.", "SH-01 (CFO)", "HMW-08", "Must", "Sum of atomic charge components equals Total_Billed_Cost with zero discrepancy."),
        ("BR-09", "The business needs to compare expenditure patterns across admission types (Emergency vs. Elective vs. Urgent) so that emergency care surge pricing can be monitored.", "SH-02 (Clinical)", "HMW-09", "Should", "ANOVA p-value indicating statistically significant cost variance between categories."),
        ("BR-10", "The business needs to benchmark cost performance and average daily charges across multiple hospital branches so that operational cost disparities are normalized.", "SH-01 (CFO)", "HMW-10", "Should", "Facility unit cost comparison variance within +/- 5% of network baseline."),
        ("BR-11", "The business needs to simulate the financial impact of potential price cap adjustments on procedure revenues so that executive scenario planning is enabled.", "SH-01 (CFO)", "HMW-01", "Could", "Simulated revenue delta reflects configured percentage adjustment accurately."),
        ("BR-12", "The business needs automated real-time bi-directional synchronization with external hospital Electronic Health Record (EHR) systems.", "SH-03 (Billing)", "HMW-06", "Won't", "Deferred to future enterprise phase; out-of-scope for standalone analytics MVP.")
    ]

    for idx, vals in enumerate(br_data):
        row = br_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(br_table, br_widths)

    doc.add_paragraph()

    # Section 2: Data Requirements
    doc.add_heading("2. Data Attributes Required by Business Requirements", level=1)
    doc.add_paragraph(
        "To fulfill the requirements without ambiguity, the underlying data layer must provision the following core business attributes:"
    )

    data_map = [
        ("Temporal Attributes", "Admission Date, Discharge Date, Fiscal Year, Admission Month"),
        ("Clinical Entities", "Department Name, Procedure Name, Diagnosis Code / Category, Admission Type (Emergency/Elective/Urgent)"),
        ("Demographic Entities", "Patient Age, Patient Gender, Facility / Hospital Branch Location"),
        ("Financial & Cost Metrics", "Room Rent Charges, Surgery & Doctor Fees, Pharmacy / Medicine Charges, Diagnostic / Lab Charges, Total Billed Cost"),
        ("Insurance & Settlement", "Insurance Provider Type, Covered Amount, Co-pay / Out-of-pocket Amount, Claim Settlement Status")
    ]
    for cat, attrs in data_map:
        p_dm = doc.add_paragraph(style='List Bullet')
        r_c = p_dm.add_run(f"{cat}: ")
        r_c.bold = True
        p_dm.add_run(attrs)

    doc.add_paragraph()

    # Section 3: Gap Analysis
    doc.add_heading("3. Requirements Gap & Alignment Verification", level=1)
    p_gap = doc.add_paragraph()
    p_gap.add_run("1. HMWs with no BR: ").bold = True
    p_gap.add_run("None. All 10 HMW statements map to at least one concrete Business Requirement.\n")
    p_gap.add_run("2. BRs with no HMW: ").bold = True
    p_gap.add_run("None. Every BR originates from an approved stakeholder opportunity statement.\n")
    p_gap.add_run("3. Testability Status: ").bold = True
    p_gap.add_run("100% of Must and Should requirements have objective, mathematically verifiable acceptance metrics.")

    target_path = "docs/03_Needs_Analysis.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step03()
