"""Create fake fee statements and a matching workbook so the tool can be tried without real data.

    python samples/make_samples.py

All names, numbers and amounts here are made up.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent

STUDENTS = [
    # student number, opening, academic, additional, payments, line items
    ("1000001", "1,250.00", "48,600.00", "350.00", "-20,000.00",
     [("Residence Accommodation Fee", "32,400.00"), ("Meal Plan Charge", "18,900.00")]),
    ("1000002", "0.00", "52,100.00", "0.00", "-52,100.00",
     [("Self Catering Accommodation", "28,750.00")]),
    ("1000003", "-430.00", "45,980.00", "1,120.00", "-10,500.00",
     [("Private Rent Allowance", "24,000.00"), ("Meal Plan Charge", "9,450.00"), ("Food Voucher", "1,200.00")]),
]


def make_pdf(path: Path, sn: str, opening: str, academic: str, additional: str, payments: str, items) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    y = 800
    for line in [
        "SAMPLE UNIVERSITY - STUDENT FEE STATEMENT",
        f"Student Number: {sn}",
        "Name: Sample Student",
        "",
        "Summary",
        f"Opening Balance {opening}",
        f"Academic Charges {academic}",
        f"Additional Charges / Credits {additional}",
        f"Payments / Miscellaneous Credits {payments}",
        "",
        "Transactions",
        *[f"{desc} {amt}" for desc, amt in items],
    ]:
        c.drawString(60, y, line)
        y -= 20
    c.save()


def main() -> None:
    pdf_dir = HERE / "statements"
    pdf_dir.mkdir(exist_ok=True)
    for i, (sn, op, ac, ad, pay, items) in enumerate(STUDENTS, start=1):
        # Deliberately unhelpful file names, to show the renaming step.
        make_pdf(pdf_dir / f"statement_download_{i}.pdf", sn, op, ac, ad, pay, items)

    wb = Workbook()
    ws = wb.active
    ws.append(["Student Number", "Total Academic Charges", "Accommodation", "Meal Charges",
               "Total Additional Charges", "Total Payments", "Opening Balance"])
    for sn in ["1000001", "1000002", "1000004"]:  # 1000003 missing, 1000004 has no statement
        ws.append([sn])
    wb.save(HERE / "sample_workbook.xlsx")
    print(f"Created {len(STUDENTS)} PDFs in {pdf_dir} and {HERE / 'sample_workbook.xlsx'}")


if __name__ == "__main__":
    main()
