"""
Master correction + logo insertion script.
Fixes: Project ID = P106, Mentor = Mr. Dinesh, Domain = Data Analyst
Adds: ABES logo to cover pages of all Word documents
"""
import re
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import subprocess, sys

ROOT  = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")
LOGO  = ROOT / "docs" / "abes_logo.jpg"
DOCS  = ROOT / "docs"

# ── 1. Global text replacements in ALL .py and .md files ──────────────────────
OLD_NEW = [
    ("Mr. Dinesh",   "Mr. Dinesh"),
    ("Mr. Dinesh",         "Mr. Dinesh"),
    ("mr. dinesh",         "mr. dinesh"),
    ("MR. DINESH",         "MR. DINESH"),
    ("Data Analyst",          "Data Analyst"),
    ("data analyst",          "data analyst"),
    ("DATA ANALYST",          "DATA ANALYST"),
    ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01","ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
    ("Project ID: P106",         "Project ID: P106"),
    ("P106 - Treatment Cost",    "P106 - Treatment Cost"),
    ("\"033\"",                 '"P106"'),
    ("'P106'",                   "'P106'"),
    ("Project Document | P106",  "Project Document | P106"),
    ("docs/P106_Project_Document.docx", "docs/P106_Project_Document.docx"),
    ("P106_Project_Document",    "P106_Project_Document"),
    ("P106_Presentation",        "P106_Presentation"),
]

changed = []
for f in ROOT.rglob("*"):
    if f.suffix not in (".py", ".md") or ".git" in str(f) or ".venv" in str(f):
        continue
    try:
        text = f.read_text(encoding="utf-8")
        orig = text
        for old, new in OLD_NEW:
            text = text.replace(old, new)
        if text != orig:
            f.write_text(text, encoding="utf-8")
            changed.append(f.name)
            print(f"  Fixed: {f.name}")
    except Exception as e:
        print(f"  SKIP {f.name}: {e}")

print(f"\nText corrections applied to {len(changed)} files.")

# ── 2. Add logo to existing Word documents ─────────────────────────────────────
WORD_DOCS = [
    DOCS / "01_Project_Understanding.docx",
    DOCS / "02_HMW_Statements.docx",
    DOCS / "03_Needs_Analysis.docx",
    DOCS / "05_BRD.docx",
    DOCS / "07_TRD.docx",
    DOCS / "09_ER_Diagram.docx",
    DOCS / "11_User_Stories.docx",
    DOCS / "12_UI_Framework.docx",
    DOCS / "13_Test_Cases.docx",
    DOCS / "14_BUILD_PLAN.docx",
    DOCS / "15_Application_Source_Code_and_Test_Report.docx",
    DOCS / "16_Deploy_Guide.docx",
    DOCS / "P106_Project_Document.docx",  # will be generated fresh below
]

def insert_logo_in_header(doc_path):
    """Insert ABES logo in the header of an existing docx (left-aligned)."""
    if not doc_path.exists():
        print(f"  SKIP (not found): {doc_path.name}")
        return False
    try:
        doc = Document(str(doc_path))
        for section in doc.sections:
            hdr = section.header
            hdr.is_linked_to_previous = False
            # find or create paragraph
            p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # Insert logo inline
            run_logo = p.add_run()
            run_logo.add_picture(str(LOGO), width=Inches(1.1))
            # Add college name text after logo
            run_text = p.add_run("  ABES Engineering College")
            run_text.font.size = Pt(8)
            run_text.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
            run_text.font.name = "Calibri"
        doc.save(str(doc_path))
        print(f"  Logo added: {doc_path.name}")
        return True
    except Exception as e:
        print(f"  ERROR {doc_path.name}: {e}")
        return False

print("\nInserting logo into document headers...")
for doc_path in WORD_DOCS:
    insert_logo_in_header(doc_path)

print("\nDone — logo inserted. Now regenerating fresh documents with corrections...")
