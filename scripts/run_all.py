"""
Insert ABES logo + apply all corrections, then regenerate every document.
Logo goes on the cover page of each document, right-aligned next to the college name.
"""
import subprocess, sys
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")
LOGO = ROOT / "docs" / "abes_logo.jpg"
DOCS = ROOT / "docs"

# ── Step 1: global text corrections in all scripts ────────────────────────────
OLD_NEW = [
    ("Prof. Kalpana Dwivedi",   "Mr. Dinesh"),
    ("Kalpana Dwivedi",         "Mr. Dinesh"),
    ("Data Analytics",          "Data Analyst"),
    ("ABES_OND26_MCA Batch 033","ABES_OND26_MCA Batch P106"),
    ("Project ID: 033",         "Project ID: P106"),
    ("033 - Treatment Cost",    "P106 - Treatment Cost"),
    ('"033"',                   '"P106"'),
    ("'033'",                   "'P106'"),
    ("Project Document | 033",  "Project Document | P106"),
    ("docs/033_Project_Document.docx", "docs/P106_Project_Document.docx"),
    ("033_Project_Document",    "P106_Project_Document"),
    ("033_Presentation",        "P106_Presentation"),
]

print("── Applying text corrections ──")
for f in ROOT.rglob("*"):
    if f.suffix not in (".py", ".md") or ".git" in str(f) or ".venv" in str(f):
        continue
    if f.name == "run_all.py":
        continue
    try:
        text = f.read_text(encoding="utf-8")
        orig = text
        for old, new in OLD_NEW:
            text = text.replace(old, new)
        if text != orig:
            f.write_text(text, encoding="utf-8")
            print(f"  Fixed: {f.name}")
    except Exception as e:
        pass

# ── Step 2: Regenerate all documents ─────────────────────────────────────────
GENERATORS = [
    ROOT / "scripts/generate_step01_doc.py",
    ROOT / "scripts/pipeline/generate_step02_hmw.py",
    ROOT / "scripts/pipeline/generate_step03_needs.py",
    ROOT / "scripts/pipeline/generate_step04_rtm.py",
    ROOT / "scripts/pipeline/generate_step05_brd.py",
    ROOT / "scripts/pipeline/generate_step07_trd.py",
    ROOT / "scripts/pipeline/generate_step11_user_stories.py",
    ROOT / "scripts/pipeline/generate_step12_ui_framework.py",
    ROOT / "scripts/pipeline/generate_step13_test_cases.py",
    ROOT / "scripts/pipeline/generate_step14_build_plan.py",
    ROOT / "scripts/pipeline/generate_step15_test_report.py",
    ROOT / "scripts/pipeline/generate_step16_deploy.py",
    ROOT / "scripts/pipeline/generate_step17_abes_project_document.py",
    ROOT / "scripts/pipeline/generate_step18_abes_presentation.py",
    ROOT / "scripts/pipeline/generate_validation_report_doc.py",
    ROOT / "scripts/pipeline/generate_github_streamlit_details.py",
]

print("\n── Regenerating documents ──")
for gen in GENERATORS:
    if not gen.exists():
        print(f"  SKIP (not found): {gen.name}")
        continue
    result = subprocess.run(
        [sys.executable, str(gen)],
        capture_output=True, text=True, cwd=str(ROOT)
    )
    if result.returncode == 0:
        last_line = (result.stdout.strip().split("\n") or [""])[-1]
        print(f"  OK: {gen.name}  → {last_line}")
    else:
        err = result.stderr.strip().split("\n")[-1] if result.stderr else "unknown error"
        print(f"  ERROR: {gen.name}: {err}")

# ── Step 3: Insert logo into every generated Word doc ───────────────────────
WORD_DOCS = sorted(DOCS.glob("*.docx"))

def insert_logo_cover(doc_path):
    """Add ABES logo as an inline image at the very top of the document (before page 1 content)."""
    if not doc_path.exists():
        return False
    try:
        doc = Document(str(doc_path))
        # Insert a new paragraph at position 0 with the logo
        # Find first paragraph and insert before it
        first_para = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
        
        # Add a new paragraph at the top
        new_p = doc.add_paragraph()
        new_p._element.getparent().remove(new_p._element)
        first_para._element.getparent().insert(
            list(first_para._element.getparent()).index(first_para._element), 
            new_p._element
        )
        new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        pPr = new_p._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '60')
        pPr.append(sp)
        run = new_p.add_run()
        run.add_picture(str(LOGO), width=Inches(1.5))
        
        doc.save(str(doc_path))
        return True
    except Exception as e:
        print(f"    Logo error on {doc_path.name}: {e}")
        return False

# Also insert in headers for clean look
def insert_logo_header(doc_path):
    """Replace plain-text header with logo + text header."""
    if not doc_path.exists():
        return False
    try:
        doc = Document(str(doc_path))
        for section in doc.sections:
            hdr = section.header
            if hdr.is_linked_to_previous:
                continue
            p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_logo = p.add_run()
            run_logo.add_picture(str(LOGO), width=Inches(0.9))
            run_text = p.add_run("   ABES Engineering College  |  Treatment Cost Trend Analysis  |  P106")
            run_text.font.size = Pt(7.5)
            run_text.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
            run_text.font.name = "Calibri"
        doc.save(str(doc_path))
        return True
    except Exception as e:
        return False

print("\n── Inserting logo into headers ──")
logo_ok = 0
for doc_path in WORD_DOCS:
    # Skip submission-renamed copies (avoid double processing)
    if doc_path.name.startswith("2500320140033"):
        continue
    if insert_logo_header(doc_path):
        print(f"  Logo header: {doc_path.name}")
        logo_ok += 1

print(f"\nLogo inserted into {logo_ok} documents.")
print("\nAll done! Ready to convert to PDF.")
