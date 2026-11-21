"""
Create Project ZIP for LMS submission — Step 9
Excludes: .venv, venv, __pycache__, .git
Roll No: 2500320140033 | Arohi Tyagi
"""
import zipfile
import os
from pathlib import Path
import datetime

ROOT = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis")
DESKTOP = Path(r"C:\Users\Arohi Tyagi\OneDrive\Desktop")

ZIP_NAME = "2500320140033_ArohiTyagi_Project_ZIP.zip"
ZIP_PATH = ROOT / "docs" / ZIP_NAME

EXCLUDES = {".venv", "venv", "__pycache__", ".git", "node_modules"}

def should_exclude(path: Path) -> bool:
    for part in path.parts:
        if part in EXCLUDES:
            return True
    return False

def create_zip():
    total_files = 0
    total_bytes = 0

    with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for file_path in ROOT.rglob("*"):
            if file_path.is_file():
                rel = file_path.relative_to(ROOT.parent)
                if should_exclude(rel):
                    continue
                # Also skip the zip itself if it somehow exists already
                if file_path == ZIP_PATH:
                    continue
                try:
                    arcname = str(rel).replace("\\", "/")
                    zf.write(file_path, arcname)
                    total_files += 1
                    total_bytes += file_path.stat().st_size
                    if total_files % 20 == 0:
                        print(f"  Added {total_files} files...")
                except Exception as e:
                    print(f"  [SKIP] {rel}: {e}")

    zip_size_mb = ZIP_PATH.stat().st_size / (1024 * 1024)
    print(f"\nZIP created: docs/{ZIP_NAME}")
    print(f"Files zipped:  {total_files}")
    print(f"Original size: {total_bytes / (1024*1024):.2f} MB")
    print(f"ZIP size:      {zip_size_mb:.2f} MB")
    if zip_size_mb > 100:
        print(f"WARNING: ZIP exceeds 100 MB LMS limit ({zip_size_mb:.1f} MB). Upload files separately.")
    else:
        print(f"OK: ZIP is within 100 MB limit.")

if __name__ == "__main__":
    print(f"Creating ZIP: {ZIP_NAME}")
    print(f"Root: {ROOT}")
    print(f"Excluding: {EXCLUDES}")
    create_zip()
