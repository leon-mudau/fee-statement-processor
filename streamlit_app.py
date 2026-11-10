"""Browser version: streamlit run streamlit_app.py"""
from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from fee_processor import process_files

st.set_page_config(page_title="Fee Statement Processor", page_icon="📄", layout="wide")
st.title("Fee Statement Processor")
st.caption("Upload the student workbook and fee statement PDFs, then run the extraction.")

workbook_file = st.file_uploader("Student workbook (.xlsx)", type=["xlsx"])
pdf_files = st.file_uploader("Fee statement PDFs", type=["pdf"], accept_multiple_files=True)
append_missing = st.checkbox("Add rows for students not in the workbook")

if st.button("Run Processing", type="primary", disabled=not (workbook_file and pdf_files)):
    with TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        in_dir = tmp_dir / "in"
        in_dir.mkdir()
        wb_path = tmp_dir / workbook_file.name
        wb_path.write_bytes(workbook_file.getvalue())
        pdf_paths = []
        for f in pdf_files:
            p = in_dir / f.name
            p.write_bytes(f.getvalue())
            pdf_paths.append(p)

        bar = st.progress(0, text="Starting...")
        notes: list[str] = []
        wb_out, zip_out, actions = process_files(
            wb_path,
            pdf_paths,
            tmp_dir / "out",
            append_missing=append_missing,
            progress=lambda r, m: bar.progress(int(r * 100), text=m),
            note=notes.append,
        )

        st.subheader("Result summary")
        st.dataframe(actions, use_container_width=True)
        for n in notes:
            st.warning(n)

        c1, c2 = st.columns(2)
        c1.download_button("Download populated workbook", wb_out.read_bytes(), file_name="populated_workbook.xlsx")
        c2.download_button("Download renamed statements (ZIP)", zip_out.read_bytes(), file_name="renamed_fee_statements.zip")
