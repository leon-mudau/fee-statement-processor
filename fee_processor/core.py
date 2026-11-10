"""Core logic: read fee statement PDFs, pull out the figures, and write them into an Excel workbook."""
from __future__ import annotations

import re
import zipfile
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from pypdf import PdfReader

# Row 1 headers the workbook may use for each field (matched case and punctuation insensitive).
HEADER_ALIASES = {
    "student_number": ["student number", "studentnumber", "student no", "studentno"],
    "academic_charges": ["total academic charges", "academic charges"],
    "accommodation": ["accommodation"],
    "meal": ["meal charges", "meal"],
    "additional_charges": ["total additional charges", "additional charges / credits", "additional charges"],
    "payments_misc": ["total payments", "payments / miscellaneous credits"],
    "opening_balance": ["opening balance"],
}

FINANCE_KEYS = [
    "academic_charges",
    "accommodation",
    "meal",
    "additional_charges",
    "payments_misc",
    "opening_balance",
]

MONEY = r"-?[\d,]+\.\d{2}"


@dataclass
class Record:
    student_number: str
    opening_balance: float | None
    academic_charges: float | None
    additional_charges: float | None
    payments_misc: float | None
    accommodation: float
    meal: float


def parse_money(raw_text: str) -> float:
    return float(raw_text.replace(",", "").strip())


def pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def parse_statement_text(text: str, fallback_name: str = "") -> Record:
    """Turn the raw text of one statement into a Record."""

    def find_summary(label: str) -> float | None:
        match = re.search(rf"{re.escape(label)}\s+({MONEY})", text, flags=re.IGNORECASE)
        return parse_money(match.group(1)) if match else None

    sn_match = re.search(r"student\s*number\s*[:\-]?\s*(\d{6,10})", text, flags=re.IGNORECASE)
    if sn_match:
        student_number = sn_match.group(1)
    else:
        # Fallback for files already renamed to the student number when text extraction is noisy.
        stem_match = re.fullmatch(r"(\d{6,10})", Path(fallback_name).stem)
        if not stem_match:
            raise ValueError(f"Student number not found in {fallback_name or 'statement'}")
        student_number = stem_match.group(1)

    accommodation = 0.0
    meal = 0.0
    for raw_line in text.splitlines():
        line = raw_line.strip()
        amt_match = re.search(rf"({MONEY})\s*$", line)
        if not amt_match:
            continue
        amount = parse_money(amt_match.group(1))
        lower = line.lower()
        if " meal " in f" {lower} " or "food" in lower:
            meal += amount
        elif any(t in lower for t in ["accom", "rent", "self-catering", "self catering"]):
            accommodation += amount

    return Record(
        student_number=student_number,
        opening_balance=find_summary("Opening Balance"),
        academic_charges=find_summary("Academic Charges"),
        additional_charges=find_summary("Additional Charges / Credits"),
        payments_misc=find_summary("Payments / Miscellaneous Credits"),
        accommodation=round(accommodation, 2),
        meal=round(meal, 2),
    )


def extract_statement_data(pdf_path: Path) -> Record:
    return parse_statement_text(pdf_text(pdf_path), pdf_path.name)


def normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


def map_headers(ws) -> dict[str, int]:
    raw = {idx: normalize_header(ws.cell(row=1, column=idx).value) for idx in range(1, ws.max_column + 1)}
    mapped: dict[str, int] = {}
    for key, aliases in HEADER_ALIASES.items():
        wanted = {normalize_header(a) for a in aliases}
        for col_idx, header in raw.items():
            if header in wanted:
                mapped[key] = col_idx
                break
    missing = [k for k in HEADER_ALIASES if k not in mapped]
    if missing:
        raise ValueError(f"Missing required headers in row 1: {missing}")
    return mapped


def first_empty_row(ws, column_index: int, start_row: int = 2) -> int:
    row = start_row
    while str(ws.cell(row=row, column=column_index).value or "").strip():
        row += 1
    return row


def write_records_to_workbook(
    workbook_path: Path,
    records: list[Record],
    append_missing: bool = False,
    note: Callable[[str], None] = print,
) -> list[dict[str, object]]:
    """Fill the finance columns for each student row. Returns one action per student."""
    wb = load_workbook(workbook_path)
    ws = wb[wb.sheetnames[0]]
    col_map = map_headers(ws)
    student_col = col_map["student_number"]

    existing: dict[str, int] = {}
    for row_idx in range(2, ws.max_row + 1):
        sn = str(ws.cell(row=row_idx, column=student_col).value or "").strip()
        if not sn:
            continue
        if sn in existing:
            note(f"Duplicate student number in sheet: {sn} (rows {existing[sn]} and {row_idx})")
        else:
            existing[sn] = row_idx

    fonts = {key: copy(ws.cell(row=2, column=col_map[key]).font) for key in FINANCE_KEYS}
    actions: list[dict[str, object]] = []
    seen: set[str] = set()

    for rec in records:
        seen.add(rec.student_number)
        if rec.student_number in existing:
            row_idx, action = existing[rec.student_number], "updated"
        elif append_missing:
            row_idx, action = first_empty_row(ws, student_col), "appended"
            existing[rec.student_number] = row_idx
        else:
            note(f"No matching student number in sheet for statement: {rec.student_number}")
            actions.append({"student_number": rec.student_number, "row": None, "action": "statement_no_match"})
            continue

        ws.cell(row=row_idx, column=student_col, value=rec.student_number)
        for key in FINANCE_KEYS:
            cell = ws.cell(row=row_idx, column=col_map[key], value=getattr(rec, key))
            cell.number_format = "#,##0.00"
            cell.font = copy(fonts[key])
        actions.append({"student_number": rec.student_number, "row": row_idx, "action": action})

    for sn in sorted(set(existing) - seen):
        note(f"Student in sheet has no matching fee statement: {sn}")
        actions.append({"student_number": sn, "row": existing[sn], "action": "sheet_no_statement"})

    wb.save(workbook_path)
    return actions


def zip_renamed_statements(pdf_dir: Path, records: list[Record], zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for rec in records:
            f = pdf_dir / f"{rec.student_number}.pdf"
            if f.exists():
                zf.write(f, arcname=f"Fee statements/{rec.student_number}.pdf")


def process_files(
    workbook_path: Path,
    pdf_paths: list[Path],
    output_dir: Path,
    overwrite_original: bool = False,
    append_missing: bool = False,
    progress: Callable[[float, str], None] = lambda r, m: None,
    note: Callable[[str], None] = print,
) -> tuple[Path, Path, list[dict[str, object]]]:
    """Full pipeline: extract every PDF, copy them under the student number, fill the workbook, zip the PDFs."""
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_out = workbook_path if overwrite_original else output_dir / "populated_workbook.xlsx"
    if workbook_out != workbook_path:
        workbook_out.write_bytes(workbook_path.read_bytes())

    renamed_dir = output_dir / "Renamed statements"
    renamed_dir.mkdir(parents=True, exist_ok=True)

    records: list[Record] = []
    total = max(len(pdf_paths), 1)
    for idx, pdf in enumerate(sorted(pdf_paths), start=1):
        progress((idx - 1) / total * 0.7, f"Processing {idx}/{total}: {pdf.name}")
        try:
            rec = extract_statement_data(pdf)
        except Exception as exc:  # keep going so one bad PDF doesn't stop the batch
            note(f"Failed to process {pdf.name}: {exc}")
            continue
        records.append(rec)
        (renamed_dir / f"{rec.student_number}.pdf").write_bytes(pdf.read_bytes())

    if not records:
        note("No valid statements were extracted. The workbook was copied but not changed.")

    records.sort(key=lambda r: r.student_number)
    progress(0.8, "Updating workbook...")
    actions = write_records_to_workbook(workbook_out, records, append_missing, note)

    zip_out = output_dir / "renamed_fee_statements.zip"
    progress(0.9, "Zipping renamed statements...")
    zip_renamed_statements(renamed_dir, records, zip_out)
    progress(1.0, "Complete")
    return workbook_out, zip_out, actions
