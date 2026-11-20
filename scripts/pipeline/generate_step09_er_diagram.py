import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from PIL import Image, ImageDraw, ImageFont
import os

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

def generate_er_diagram_image(image_path="docs/09_ER_Diagram.png"):
    width, height = 1200, 750
    img = Image.new("RGB", (width, height), color="#F8FAFC")
    draw = ImageDraw.Draw(img)

    # Colors
    c_header = "#1B365D"
    c_card = "#FFFFFF"
    c_border = "#CBD5E1"
    c_fact_header = "#A51C30"
    c_text = "#1E293B"
    c_line = "#64748B"

    # Entities: (x, y, w, h, title, is_fact, fields)
    tables = [
        (60, 50, 260, 160, "dim_departments", False, [
            "+ dept_id: INT [PK]",
            "  dept_name: TEXT [UQ]",
            "  hod_name: TEXT",
            "  budget_cap_inr: REAL"
        ]),
        (60, 270, 260, 190, "dim_procedures", False, [
            "+ procedure_id: INT [PK]",
            "# dept_id: INT [FK]",
            "  procedure_name: TEXT",
            "  base_cost_inr: REAL",
            "  complexity_level: TEXT"
        ]),
        (450, 60, 300, 360, "fact_hospital_billing", True, [
            "+ billing_id: INT [PK]",
            "# patient_id: INT [FK]",
            "# procedure_id: INT [FK]",
            "# facility_id: INT [FK]",
            "  admission_date: DATE",
            "  discharge_date: DATE",
            "  los_days: INT",
            "  admission_type: TEXT",
            "  room_charges: REAL",
            "  surgery_charges: REAL",
            "  pharmacy_charges: REAL",
            "  lab_charges: REAL",
            "  total_billed_cost: REAL",
            "  insurance_paid: REAL",
            "  out_of_pocket_paid: REAL",
            "  is_outlier: INT"
        ]),
        (880, 50, 260, 180, "dim_facilities", False, [
            "+ facility_id: INT [PK]",
            "  facility_name: TEXT [UQ]",
            "  city: TEXT",
            "  bed_capacity: INT"
        ]),
        (880, 280, 260, 200, "dim_patients", False, [
            "+ patient_id: INT [PK]",
            "  patient_name: TEXT",
            "  age: INT",
            "  gender: TEXT",
            "  city: TEXT"
        ])
    ]

    # Draw connectors first
    # dept to proc
    draw.line([(190, 210), (190, 270)], fill=c_line, width=2)
    draw.text((200, 230), "contains (1:N)", fill=c_line)

    # proc to fact
    draw.line([(320, 350), (450, 300)], fill=c_line, width=2)
    draw.text((340, 310), "performed in (1:N)", fill=c_line)

    # facility to fact
    draw.line([(880, 140), (750, 170)], fill=c_line, width=2)
    draw.text((770, 140), "hosts (1:N)", fill=c_line)

    # patient to fact
    draw.line([(880, 350), (750, 270)], fill=c_line, width=2)
    draw.text((770, 310), "incurs (1:N)", fill=c_line)

    # Draw table cards
    for x, y, w, h, name, is_fact, fields in tables:
        # Card background & border
        draw.rectangle([x, y, x+w, y+h], fill=c_card, outline=c_border, width=2)
        # Header banner
        h_color = c_fact_header if is_fact else c_header
        draw.rectangle([x, y, x+w, y+34], fill=h_color)
        draw.text((x+10, y+8), name, fill="#FFFFFF")

        # Field items
        cur_y = y + 42
        for f in fields:
            draw.text((x+12, cur_y), f, fill=c_text)
            cur_y += 18

    # Title Banner
    draw.text((450, 500), "Relational Star/Snowflake Schema Architecture", fill=c_header)
    draw.text((430, 525), "Treatment Cost Trend Analysis Platform (SQLite 3)", fill=c_line)

    img.save(image_path)
    print(f"Generated ER diagram image: {image_path}")

def build_step09():
    generate_er_diagram_image("docs/09_ER_Diagram.png")

    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("ENTITY RELATIONSHIP (ER) DIAGRAM & ARCHITECTURE")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(27, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Relational Data Modeling & Structural Integrity Specification | Step 9")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(90, 105, 120)

    doc.add_paragraph()

    # Section 1: Visual ER Diagram
    doc.add_heading("1. Visual Entity Relationship Architecture", level=1)
    if os.path.exists("docs/09_ER_Diagram.png"):
        doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].paragraph_format.space_before = Pt(8)
        doc.paragraphs[-1].paragraph_format.space_after = Pt(8)
        run_img = doc.paragraphs[-1].add_run()
        run_img.add_picture("docs/09_ER_Diagram.png", width=Inches(6.2))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap = p_cap.add_run("Figure 9.1: Entity-Relationship Model for Inpatient Treatment Cost Analytics")
    r_cap.font.size = Pt(9)
    r_cap.font.italic = True
    r_cap.font.color.rgb = RGBColor(100, 110, 120)

    doc.add_paragraph()

    # Section 2: Mermaid erDiagram Specification
    doc.add_heading("2. Mermaid Relational Schema Notation", level=1)
    doc.add_paragraph("The canonical database schema is expressed below using standard Mermaid ER diagram syntax:")

    p_mm = doc.add_paragraph()
    p_mm.paragraph_format.left_indent = Inches(0.2)
    r_mm = p_mm.add_run(
        "erDiagram\n"
        "    DIM_DEPARTMENTS ||--o{ DIM_PROCEDURES : \"offers\"\n"
        "    DIM_PROCEDURES ||--o{ FACT_HOSPITAL_BILLING : \"performed_in\"\n"
        "    DIM_FACILITIES ||--o{ FACT_HOSPITAL_BILLING : \"hosts\"\n"
        "    DIM_PATIENTS ||--o{ FACT_HOSPITAL_BILLING : \"incurs\"\n"
        "\n"
        "    DIM_DEPARTMENTS {\n"
        "        int dept_id PK\n"
        "        string dept_name\n"
        "        string hod_name\n"
        "        float budget_cap_inr\n"
        "    }\n"
        "    DIM_PROCEDURES {\n"
        "        int procedure_id PK\n"
        "        int dept_id FK\n"
        "        string procedure_name\n"
        "        float base_cost_inr\n"
        "        string complexity_level\n"
        "    }\n"
        "    DIM_FACILITIES {\n"
        "        int facility_id PK\n"
        "        string facility_name\n"
        "        string city\n"
        "        int bed_capacity\n"
        "    }\n"
        "    DIM_PATIENTS {\n"
        "        int patient_id PK\n"
        "        string patient_name\n"
        "        int age\n"
        "        string gender\n"
        "        string city\n"
        "    }\n"
        "    FACT_HOSPITAL_BILLING {\n"
        "        int billing_id PK\n"
        "        int patient_id FK\n"
        "        int procedure_id FK\n"
        "        int facility_id FK\n"
        "        date admission_date\n"
        "        date discharge_date\n"
        "        int los_days\n"
        "        string admission_type\n"
        "        float room_charges\n"
        "        float surgery_charges\n"
        "        float pharmacy_charges\n"
        "        float lab_charges\n"
        "        float total_billed_cost\n"
        "        float insurance_paid\n"
        "        float out_of_pocket_paid\n"
        "        string insurance_type\n"
        "        int is_outlier\n"
        "    }"
    )
    r_mm.font.name = "Consolas"
    r_mm.font.size = Pt(8.5)

    doc.add_paragraph()

    # Section 3: Plain-English Relationship Statements
    doc.add_heading("3. Plain-English Relationship Declarations", level=1)
    rels = [
        ("Department to Procedure", "One clinical department offers one or many medical procedures, while each procedure belongs strictly to exactly one clinical department."),
        ("Procedure to Fact Billing", "One clinical procedure is performed in zero, one, or many inpatient hospitalization events, while each hospitalization record references exactly one primary medical procedure."),
        ("Facility to Fact Billing", "One hospital facility branch hosts zero, one, or many inpatient admissions, while each billing invoice is issued by exactly one hospital facility."),
        ("Patient to Fact Billing", "One registered patient incurs one or many hospitalization billing episodes, while each billing record belongs strictly to exactly one registered patient.")
    ]
    for name, stmt in rels:
        p_r = doc.add_paragraph(style='List Bullet')
        p_r.add_run(f"{name}: ").bold = True
        p_r.add_run(stmt)

    doc.add_paragraph()

    # Section 4: Business Rules Enforced by the Schema
    doc.add_heading("4. Business Rules Enforced by Relational Architecture", level=1)
    rules = [
        ("Domain Integrity", "Foreign key constraints prevent orphaned invoices; every billing record must reference valid existing primary keys in patient, procedure, and facility master tables."),
        ("Cost Partition Balance", "CHECK constraints guarantee that total_billed_cost > 0 and component charges are non-negative, preserving the additive equation total = room + surgery + pharmacy + lab."),
        ("Temporal Feasibility", "Discharge date must equal or exceed admission date; los_days is constrained to >= 1 day."),
        ("Settlement Integrity", "Out-of-pocket amount cannot exceed total billed cost; insurance_paid + out_of_pocket_paid matches total billed cost."),
        ("Outlier Categorization", "The is_outlier attribute is bounded strictly to boolean binary values (0 or 1).")
    ]
    for r_title, r_desc in rules:
        p_ru = doc.add_paragraph(style='List Bullet')
        p_ru.add_run(f"{r_title}: ").bold = True
        p_ru.add_run(r_desc)

    doc.add_paragraph()

    # Section 5: Design Analysis & Mitigations
    doc.add_heading("5. Database Design Analysis, Identified Issues & Resolved Fixes", level=1)
    
    da_table = doc.add_table(rows=5, cols=3)
    da_widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    da_table.rows[0].cells[0].paragraphs[0].add_run("Potential Design Anti-Pattern")
    da_table.rows[0].cells[1].paragraphs[0].add_run("Identified Risk / Inefficiency")
    da_table.rows[0].cells[2].paragraphs[0].add_run("Engineered Resolution in Current Schema")

    da_data = [
        ("Department FK in Fact Table", "Storing dept_id directly in fact_hospital_billing introduces transitive redundancy since procedure_id already determines dept_id.", "Normalized dim_procedures with dept_id FK; exposed dept_name seamlessly via view v_treatment_cost_analytics."),
        ("Many-to-Many Patient Procedures", "Patients receiving multiple secondary diagnostic tests could lead to non-1NF repeating groups.", "Established procedure_id as the primary clinical treatment driving the inpatient admission; secondary labs aggregated into lab_charges."),
        ("Pre-computed Outlier Flags", "Dynamic outlier filtering based on runtime filters might disagree with static database columns.", "Included is_outlier as a stored statistical baseline flag while allowing Streamlit dynamic recalculation on filtered subsets."),
        ("String-based Date Handling in SQLite", "SQLite lacks native datetime types, risking invalid date queries.", "Enforced ISO-8601 YYYY-MM-DD date text format and leveraged strftime() in the analytical view.")
    ]
    for idx, vals in enumerate(da_data):
        row = da_table.rows[idx+1]
        for c_idx, val in enumerate(vals):
            row.cells[c_idx].paragraphs[0].add_run(val)
    style_table(da_table, da_widths)

    target_path = "docs/09_ER_Diagram.docx"
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_step09()
