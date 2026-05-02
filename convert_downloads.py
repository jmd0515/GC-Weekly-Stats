"""
Convert scraped CSV downloads to XLSX so generate_data.py can read them
unchanged. Preserves the original multi-row header structure (report title row,
date-range row, then the column header row) so smart_read still detects the
right header row.
"""

import csv
import os
import sys

from openpyxl import Workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Map: scraped CSV filename -> XLSX filename expected by generate_data.py.
CONVERSIONS = {
    "Employee_Stats.csv": "Employee_Stats.xlsx",
    "Employee_Return_Stats.csv": "Employee_Return_Stats.xlsx",
}


def convert(src, dst):
    """Read a CSV with ragged rows (report has title rows with fewer columns
    than the data rows) and write to xlsx, preserving structure exactly."""
    wb = Workbook()
    ws = wb.active
    with open(src, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = 0
        for row in reader:
            ws.append(row)
            rows += 1
    wb.save(dst)
    return rows


def main():
    converted = 0
    for csv_name, xlsx_name in CONVERSIONS.items():
        src = os.path.join(SCRIPT_DIR, csv_name)
        if not os.path.exists(src):
            continue
        dst = os.path.join(SCRIPT_DIR, xlsx_name)
        rows = convert(src, dst)
        print(f"Converted {csv_name} -> {xlsx_name} ({rows} rows)")
        converted += 1

    if converted == 0:
        print("ERROR: no CSV downloads found to convert.")
        sys.exit(1)


if __name__ == "__main__":
    main()
