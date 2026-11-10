"""Command line runner: python cli.py --workbook students.xlsx --pdfs statements/ --out output/"""
from __future__ import annotations

import argparse
from pathlib import Path

from fee_processor import process_files


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract fee statement PDFs into an Excel workbook.")
    parser.add_argument("--workbook", required=True, type=Path, help="Workbook with the expected headers in row 1.")
    parser.add_argument("--pdfs", required=True, type=Path, help="Folder containing fee statement PDFs.")
    parser.add_argument("--out", default=Path("output"), type=Path, help="Output folder (default: ./output).")
    parser.add_argument("--append-missing", action="store_true", help="Add rows for students not yet in the workbook.")
    args = parser.parse_args()

    pdfs = sorted(p for p in args.pdfs.glob("*.pdf") if p.is_file())
    if not pdfs:
        print(f"No PDFs found in {args.pdfs}")
        return 1

    wb_out, zip_out, actions = process_files(
        args.workbook,
        pdfs,
        args.out,
        append_missing=args.append_missing,
        progress=lambda r, m: print(f"[{r:>4.0%}] {m}"),
        note=lambda m: print(f"  note: {m}"),
    )

    print("\nSummary")
    for a in actions:
        print(f"  {a['student_number']:<10} row={a['row']!s:<5} {a['action']}")
    print(f"\nWorkbook: {wb_out}\nZIP:      {zip_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
