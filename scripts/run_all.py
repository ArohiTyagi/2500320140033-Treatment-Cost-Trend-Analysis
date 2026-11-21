"""
Master generation script.
Regenerates all project documents with:
- Project ID: P106
- Project Title: Treatment Cost Trend Analysis
- Domain: Data Analyst
- Programme / Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
- Student: Arohi Tyagi (Roll No: 2500320140033)
- Mentor: Mr. Dinesh
- ABES Logo: placed ONLY where required (Cover page / placeholder)
- Dashboard Photo: placed ONLY where required (GitHub/Streamlit details, Presentation Slide 7, UI Framework, Step 17)
- NO logo in headers of running pages!
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")
DOCS = ROOT / "docs"

GENERATORS = [
    ROOT / "scripts/generate_step01_doc.py",
    ROOT / "scripts/pipeline/generate_step02_hmw.py",
    ROOT / "scripts/pipeline/generate_step03_needs.py",
    ROOT / "scripts/pipeline/generate_step04_rtm.py",
    ROOT / "scripts/pipeline/generate_step05_brd.py",
    ROOT / "scripts/pipeline/generate_step07_trd.py",
    ROOT / "scripts/pipeline/generate_step08_db_design.py",
    ROOT / "scripts/pipeline/generate_step09_er_diagram.py",
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
    ROOT / "scripts/pipeline/generate_source_code_reportlab.py",
]

def main():
    print("==================================================")
    print("REGENERATING ALL 18-STEP PROJECT DELIVERABLES")
    print("==================================================")
    
    success = 0
    for gen in GENERATORS:
        if not gen.exists():
            print(f"  [SKIP] Not found: {gen.name}")
            continue
        print(f"Running: {gen.name} ...")
        res = subprocess.run([sys.executable, str(gen)], capture_output=True, text=True, cwd=str(ROOT))
        if res.returncode == 0:
            last = [line for line in res.stdout.strip().split("\n") if line.strip()]
            status = last[-1] if last else "OK"
            print(f"  -> {status}")
            success += 1
        else:
            err = res.stderr.strip() or res.stdout.strip()
            print(f"  -> ERROR in {gen.name}: {err[:160]}")

    print(f"\nCompleted {success}/{len(GENERATORS)} document generators.")
    print("Headers are kept clean (no logos on running pages).")
    print("Logo and Dashboard photos are embedded strictly where required.")

if __name__ == "__main__":
    main()
