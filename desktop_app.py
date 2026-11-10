"""Offline desktop version (Tkinter). Package for staff PCs with: pyinstaller --onedir --windowed desktop_app.py"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import BooleanVar, DoubleVar, StringVar, Tk, filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from fee_processor import process_files


def configure_windows_dpi() -> None:
    """Sharper text on high-DPI Windows screens. Does nothing elsewhere."""
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


class DesktopApp:
    def __init__(self) -> None:
        configure_windows_dpi()
        self.root = Tk()
        self.root.title("Fee Statement Processor")
        self.root.geometry("980x680")
        self.root.minsize(920, 620)

        self.workbook_var = StringVar()
        self.output_var = StringVar(value=str(Path.home() / "Desktop"))
        self.overwrite_var = BooleanVar(value=False)
        self.append_var = BooleanVar(value=False)
        self.progress_var = DoubleVar(value=0.0)
        self.status_var = StringVar(value="Select files, then click Run")
        self.selected_pdfs: list[Path] = []

        self._set_styles()
        self._build_ui()

    def _set_styles(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background="#f6f8fb")
        style.configure("Title.TLabel", background="#f6f8fb", foreground="#153e75", font=("Segoe UI", 20, "bold"))
        style.configure("TLabel", font=("Segoe UI", 10), background="#f6f8fb")
        style.configure("Status.TLabel", font=("Segoe UI", 10, "bold"), foreground="#153e75", background="#f6f8fb")
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))

    def _build_ui(self) -> None:
        frm = ttk.Frame(self.root, padding=16, style="App.TFrame")
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Fee Statement Processor", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frm,
            text="Pick the student workbook and the fee statement PDFs, then click Run. "
            "The original workbook is not changed unless you tick overwrite.",
            wraplength=900,
        ).pack(anchor="w", pady=(2, 10))

        wb = ttk.LabelFrame(frm, text="1) Workbook")
        wb.pack(fill="x", pady=6)
        ttk.Entry(wb, textvariable=self.workbook_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(wb, text="Browse", command=self.pick_workbook).pack(side="right", padx=8, pady=8)

        pdf = ttk.LabelFrame(frm, text="2) Fee statements (PDFs)")
        pdf.pack(fill="x", pady=6)
        self.pdf_label = ttk.Label(pdf, text="No PDFs selected")
        self.pdf_label.pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(pdf, text="Select PDFs", command=self.pick_pdfs).pack(side="right", padx=8, pady=8)

        out = ttk.LabelFrame(frm, text="3) Output folder")
        out.pack(fill="x", pady=6)
        ttk.Entry(out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(out, text="Browse", command=self.pick_output).pack(side="right", padx=8, pady=8)

        ttk.Checkbutton(frm, text="Add rows for students not in the workbook", variable=self.append_var).pack(anchor="w")
        ttk.Checkbutton(frm, text="Overwrite selected workbook (advanced)", variable=self.overwrite_var).pack(anchor="w")

        self.run_btn = ttk.Button(frm, text="Run Processing", command=self.run_clicked)
        self.run_btn.pack(anchor="w", pady=(10, 6))
        ttk.Progressbar(frm, variable=self.progress_var, maximum=100).pack(fill="x", pady=(6, 6))
        ttk.Label(frm, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")

        self.notes_box = ScrolledText(frm, height=6, font=("Consolas", 10), wrap="word", state="disabled")
        self.notes_box.pack(fill="both", pady=(10, 6))

        self.log = ttk.Treeview(frm, columns=("student", "row", "action"), show="headings", height=10)
        for col, label, width in [("student", "Student Number", 180), ("row", "Row", 80), ("action", "Action", 160)]:
            self.log.heading(col, text=label)
            self.log.column(col, width=width)
        self.log.pack(fill="both", expand=True)

    def pick_workbook(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.workbook_var.set(path)

    def pick_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.selected_pdfs = [Path(p) for p in paths]
            self.pdf_label.config(text=f"{len(self.selected_pdfs)} PDF(s) selected")

    def pick_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def set_progress(self, ratio: float, message: str) -> None:
        self.root.after(0, lambda: (self.progress_var.set(ratio * 100), self.status_var.set(message)))

    def append_note(self, message: str) -> None:
        def _append() -> None:
            self.notes_box.configure(state="normal")
            self.notes_box.insert("end", f"- {message}\n")
            self.notes_box.see("end")
            self.notes_box.configure(state="disabled")

        self.root.after(0, _append)

    def run_clicked(self) -> None:
        workbook = Path(self.workbook_var.get().strip()) if self.workbook_var.get().strip() else None
        output = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if not workbook or not workbook.exists():
            messagebox.showerror("Missing workbook", "Please select a valid .xlsx workbook.")
            return
        if not self.selected_pdfs:
            messagebox.showerror("Missing PDFs", "Please select one or more fee statement PDFs.")
            return
        if not output:
            messagebox.showerror("Missing output folder", "Please select an output folder.")
            return

        self.run_btn.config(state="disabled")
        self.log.delete(*self.log.get_children())

        def worker() -> None:
            try:
                wb_out, zip_out, actions = process_files(
                    workbook,
                    self.selected_pdfs,
                    output,
                    overwrite_original=self.overwrite_var.get(),
                    append_missing=self.append_var.get(),
                    progress=self.set_progress,
                    note=self.append_note,
                )
                for a in actions:
                    self.root.after(0, lambda a=a: self.log.insert("", "end", values=(a["student_number"], a["row"], a["action"])))
                messagebox.showinfo("Completed", f"Workbook: {wb_out}\nRenamed ZIP: {zip_out}")
                if sys.platform == "win32" and messagebox.askyesno("Open output folder", "Open the output folder now?"):
                    os.startfile(str(output))
            except Exception:
                self.status_var.set("Failed")
                messagebox.showerror("Processing failed", traceback.format_exc())
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    DesktopApp().root.mainloop()
