"""
generate_final_set.py - The ONE script the user runs.

Steps:
  1. User opens master/Master_Input.xlsx and fills in their company's data.
  2. User runs: python generate_final_set.py
  3. Out comes:  output/Final_Set.docx  (the complete bound AGM document set)

That's it.
"""

import os
import sys
import time

# Allow running this script from anywhere - resolve paths relative to script location
HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE) if os.path.basename(HERE) == "scripts" else HERE
sys.path.insert(0, os.path.join(PROJECT_ROOT, "scripts"))

from loader import load_master
from gen_final_set import build_final_set
from validate_master import validate


def main():
    master_path = os.path.join(PROJECT_ROOT, "master", "Master_Input.xlsx")
    output_path = os.path.join(PROJECT_ROOT, "output", "Final_Set.docx")

    if not os.path.exists(master_path):
        print(f"ERROR: Master input file not found at {master_path}")
        print("Please make sure Master_Input.xlsx is in the master/ folder.")
        sys.exit(1)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    print("=" * 60)
    print("  AGM FINAL SET GENERATOR")
    print("=" * 60)
    print(f"Reading master input:  {master_path}")
    t0 = time.time()
    # load_master also writes back auto-generated share-capital descriptions,
    # so run validation *after* it to avoid false positives on those cells.
    data = load_master(master_path)
    validate(master_path)
    print(f"  Company: {data['company']['name']}")
    print(f"  CIN: {data['company']['cin']}")
    print(f"  AGM: {data['agm']['number']} on {data['agm']['date']}")
    print(f"  Directors: {len(data['directors'])}")
    print(f"  Shareholders: {len(data['shareholders'])}")
    print(f"  Board Meetings: {len(data['board_meetings'])}")

    print(f"\nGenerating Final Set...")
    build_final_set(data, output_path)
    elapsed = time.time() - t0

    print("=" * 60)
    print(f"DONE in {elapsed:.1f} seconds")
    print(f"Open this file in MS Word: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
