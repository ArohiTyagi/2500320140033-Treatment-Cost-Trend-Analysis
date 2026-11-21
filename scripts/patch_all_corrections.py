"""Patch all remaining 033/mentor/domain references to P106/Mr. Dinesh/Data Analyst"""
import re
from pathlib import Path

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")

# File-specific patches
patches = {
    "scripts/pipeline/generate_step17_abes_project_document.py": [
        ("Project Document | P106",                                "Project Document | P106"),
        ("P106 - Treatment Cost Trend Analysis",                   "P106 - Treatment Cost Trend Analysis"),
        ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",                              "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("PROJECT ID & TITLE: 033",                              "PROJECT ID & TITLE: P106"),
        ("BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",
         "BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("033-treatment-cost-analysis",                           "2500320140033-treatment-cost-trend-analysis"),
        ("docs/P106_Project_Document.docx",                        "docs/P106_Project_Document.docx"),
        ('"Project ID & Title", "033 -',                          '"Project ID & Title", "P106 -'),
        ("\"Mentor\", \"Prof. Kalpana",                           '"Mentor", "Mr. Dinesh"'),
        ("Mentor: Mr. Dinesh",                         "Mentor: Mr. Dinesh"),
        ("Mr. Dinesh",                                        "Mr. Dinesh"),
    ],
    "scripts/generate_step01_doc.py": [
        ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",  "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("\"033\"",                    '"P106"'),
        ("'P106'",                      "'P106'"),
        ("Mr. Dinesh",            "Mr. Dinesh"),
    ],
    "scripts/pipeline/generate_step18_abes_presentation.py": [
        ("ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",  "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
        ("P106",                        "P106"),
        ("Mr. Dinesh",            "Mr. Dinesh"),
    ],
    "scripts/pipeline/generate_step04_rtm.py": [
        ("Mr. Dinesh",  "Mr. Dinesh"),
    ],
    "scripts/pipeline/generate_step05_brd.py": [
        ("Mr. Dinesh",  "Mr. Dinesh"),
        ("P106",              "P106"),
    ],
}

# Global patch across ALL .py files for remaining refs
global_patches = [
    ("Mr. Dinesh",  "Mr. Dinesh"),
    ("Mr. Dinesh",        "Mr. Dinesh"),
    ("Data Analyst",         "Data Analyst"),
    ("MCA 2nd Year Data Analyst - ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01",
     "MCA 2nd Year Data Analyst - ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 2nd Year_Data Analytics_Batch 01"),
]

changed_files = []

# Apply file-specific patches
for rel_path, patch_list in patches.items():
    p = ROOT / rel_path
    if not p.exists():
        print(f"SKIP (not found): {rel_path}")
        continue
    text = p.read_text(encoding="utf-8")
    orig = text
    for old, new in patch_list:
        text = text.replace(old, new)
    if text != orig:
        p.write_text(text, encoding="utf-8")
        changed_files.append(rel_path)
        print(f"Patched: {rel_path}")

# Apply global patches to all .py files
for f in ROOT.rglob("*.py"):
    if ".git" in str(f) or ".venv" in str(f):
        continue
    text = f.read_text(encoding="utf-8")
    orig = text
    for old, new in global_patches:
        text = text.replace(old, new)
    if text != orig:
        f.write_text(text, encoding="utf-8")
        rel = str(f.relative_to(ROOT))
        if rel not in changed_files:
            changed_files.append(rel)
            print(f"Global patched: {rel}")

print(f"\nTotal files changed: {len(changed_files)}")
print("Done.")
