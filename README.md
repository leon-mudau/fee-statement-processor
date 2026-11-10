# Fee Statement Processor

A small Python tool that reads batches of student fee statement PDFs, pulls out the key figures and fills them into an Excel workbook. It also renames every PDF to the student's number and zips them, ready to send on.

I built it for a university student funding team that was copying these numbers by hand, one statement at a time. Staff run it as an offline Windows app, so they need no Python, no internet and no coding.

## What it does

1. Reads each PDF and finds the student number, opening balance, academic charges, additional charges and payments.
2. Adds up accommodation and meal lines from the transaction list (rent, residence, self catering, meal plan, food).
3. Matches each statement to the right row in the workbook by student number and fills in the figures.
4. Flags problems instead of failing: statements with no matching student, students with no statement, duplicate rows and unreadable PDFs.
5. Saves a populated copy of the workbook (the original stays untouched) and a ZIP of the PDFs renamed to `<student number>.pdf`.

## Three ways to run it

| Version | File | Who it's for |
|---|---|---|
| Desktop app (Tkinter) | `desktop_app.py` | Staff. Packaged as a Windows `.exe` with PyInstaller |
| Web app (Streamlit) | `streamlit_app.py` | Running in a browser |
| Command line | `cli.py` | Scheduled or scripted runs |

All three share the same logic in `fee_processor/core.py`.

## Try it with sample data

No real student data is included. The sample script creates three made-up statements and a workbook.

```bash
pip install -r requirements.txt
python samples/make_samples.py
python cli.py --workbook samples/sample_workbook.xlsx --pdfs samples/statements --out output
```

Expected result: two students updated, one statement flagged as having no row in the workbook, and one workbook row flagged as having no statement.

Other versions:

```bash
streamlit run streamlit_app.py
python desktop_app.py
```

## Build the offline Windows app

```bash
pip install pyinstaller
pyinstaller --onedir --windowed --name FeeStatementProcessor desktop_app.py
```

Share the `dist/FeeStatementProcessor/` folder with staff.

## Workbook format

Row 1 needs these headers (case and punctuation don't matter): `Student Number`, `Total Academic Charges`, `Accommodation`, `Meal Charges`, `Total Additional Charges`, `Total Payments`, `Opening Balance`.

## Tests

```bash
pip install pytest
pytest
```

## Tech

Python · pypdf · openpyxl · Tkinter · Streamlit · PyInstaller · pytest
