from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from fee_processor.core import Record, parse_statement_text, write_records_to_workbook

SAMPLE = """Student Number: 1234567
Opening Balance 1,250.00
Academic Charges 48,600.00
Additional Charges / Credits 350.00
Payments / Miscellaneous Credits -20,000.00
Residence Accommodation Fee 32,400.00
Meal Plan Charge 18,900.00
Food Voucher 600.00
"""


def test_parse_statement_text():
    rec = parse_statement_text(SAMPLE)
    assert rec.student_number == "1234567"
    assert rec.opening_balance == 1250.00
    assert rec.academic_charges == 48600.00
    assert rec.additional_charges == 350.00
    assert rec.payments_misc == -20000.00
    assert rec.accommodation == 32400.00
    assert rec.meal == 19500.00


def test_student_number_from_file_name():
    rec = parse_statement_text("Opening Balance 0.00", fallback_name="7654321.pdf")
    assert rec.student_number == "7654321"


def test_missing_student_number_raises():
    with pytest.raises(ValueError):
        parse_statement_text("nothing useful here", fallback_name="scan.pdf")


def test_write_records(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Student Number", "Total Academic Charges", "Accommodation", "Meal Charges",
               "Total Additional Charges", "Total Payments", "Opening Balance"])
    ws.append(["1234567"])
    ws.append(["1111111"])
    path = tmp_path / "wb.xlsx"
    wb.save(path)

    rec = Record("1234567", 10.0, 100.0, 5.0, -50.0, 30.0, 20.0)
    extra = Record("9999999", 0.0, 1.0, 0.0, 0.0, 0.0, 0.0)
    actions = write_records_to_workbook(path, [rec, extra], append_missing=False, note=lambda m: None)

    by_sn = {a["student_number"]: a["action"] for a in actions}
    assert by_sn == {"1234567": "updated", "9999999": "statement_no_match", "1111111": "sheet_no_statement"}

    ws = load_workbook(path).active
    assert [c.value for c in ws[2]] == ["1234567", 100.0, 30.0, 20.0, 5.0, -50.0, 10.0]
