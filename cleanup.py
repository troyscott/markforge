"""
==============================================================================
SCRIPT NAME:   Workspace Cleanup & Reset Tool
DESCRIPTION:   A maintenance utility to reset the processing environment.

               Actions:
               1. DESTRUCTIVE: Completely deletes the Output directory (./out)
                  to prepare for a fresh run.
               2. CLEANUP: Scans the Input directory (./in) for stray temporary
                  PDF chunks (files containing '_temp_chunk_') left behind 
                  if the main converter script crashed or was interrupted.

USAGE:         Run this script before starting a fresh batch of documents
               to ensure no mixed data or ghost files remain.
==============================================================================
"""


import os
import shutil
from pathlib import Path

# CONFIG
OUTPUT_DIR = Path("./out")
INPUT_DIR = Path("./in")

def clean_environment():
    print("🧹 Starting Cleanup...")

    # 1. Clean Output Directory
    if OUTPUT_DIR.exists():
        print(f"   🗑️  Removing output folder: {OUTPUT_DIR}")
        try:
            shutil.rmtree(OUTPUT_DIR)
            print("      ✅ Output folder cleared.")
        except Exception as e:
            print(f"      ❌ Could not delete output folder: {e}")
    else:
        print("   ℹ️  Output folder already empty.")

    # 2. Clean Stray Temp PDF Chunks in Input Directory
    print(f"   🔍 Scanning '{INPUT_DIR}' for stray temp chunks...")
    count = 0
    for root, dirs, files in os.walk(INPUT_DIR):
        for file in files:
            # Look for the naming convention used in the main script
            if "_temp_chunk_" in file and file.endswith(".pdf"):
                file_path = Path(root) / file
                try:
                    os.remove(file_path)
                    print(f"      deleted: {file}")
                    count += 1
                except Exception as e:
                    print(f"      failed to delete {file}: {e}")

    if count == 0:
        print("   ✅ No stray temp files found.")
    else:
        print(f"   ✨ Removed {count} stray temp files.")

    print("\nEnvironment is clean and ready for a fresh run.")

if __name__ == "__main__":
    clean_environment()