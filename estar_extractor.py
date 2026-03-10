#!/usr/bin/env python3
"""
eSTAR Submission Package Extractor
===================================
Extracts all embedded files from an FDA eSTAR submission PDF,
converts any .docx attachments to PDF, and merges everything
into a single navigable PDF with a table of contents, section
dividers, and PDF bookmarks.

Requires: Python 3.8+
Auto-installs dependencies on first run.
"""

import subprocess
import sys
import os
import tempfile
import shutil
import threading
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Dependency management
# ---------------------------------------------------------------------------

REQUIRED_PACKAGES = {
    "pypdf": "pypdf",
    "pdfplumber": "pdfplumber",
    "reportlab": "reportlab",
    "customtkinter": "customtkinter",
    "docx2pdf": "docx2pdf",
}


def check_and_install_dependencies():
    """Check for required packages and install any that are missing."""
    missing = []
    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print(f"Installing missing packages: {', '.join(missing)}")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet"] + missing
        )
        print("All dependencies installed successfully.\n")


check_and_install_dependencies()

# ---------------------------------------------------------------------------
# Now safe to import everything
# ---------------------------------------------------------------------------

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import customtkinter as ctk
from tkinter import filedialog, messagebox

# ---------------------------------------------------------------------------
# Core extraction / merge logic
# ---------------------------------------------------------------------------

# Canonical 510(k) section ordering – files whose names contain these
# substrings (case-insensitive) are sorted in this order.  Anything
# unrecognised goes at the end, in the order it was encountered.
SECTION_ORDER_HINTS = [
    ("coverletter", "Cover Letter"),
    ("cover_letter", "Cover Letter"),
    ("exec-sum", "Executive Summary"),
    ("exec_sum", "Executive Summary"),
    ("executivesummary", "Executive Summary"),
    ("adi-", "Applicant Device Information"),
    ("adi_", "Applicant Device Information"),
    ("-se-", "Substantial Equivalence"),
    ("_se_", "Substantial Equivalence"),
    ("substantial", "Substantial Equivalence"),
    ("spec-draw", "Specifications & Drawings"),
    ("spec_draw", "Specifications & Drawings"),
    ("sys-comp", "System Component List"),
    ("sys_comp", "System Component List"),
    ("pkg-lbl", "Package Labeling"),
    ("pkg_lbl", "Package Labeling"),
    ("-pkg-", "Packaging"),
    ("_pkg_", "Packaging"),
    ("-lbl-", "Labeling"),
    ("_lbl_", "Labeling"),
    ("instruction", "Instructions For Use"),
    ("ifu", "Instructions For Use"),
    ("-sl-", "Standards List"),
    ("_sl_", "Standards List"),
    ("standards", "Standards List"),
    ("biocomp", "Biocompatibility"),
    ("bench", "Bench Testing"),
    ("performance", "Performance Testing"),
    ("hfe", "Human Factors Engineering"),
    ("human factor", "Human Factors Engineering"),
    ("coa", "Certificate of Analysis"),
    ("aca", "Accreditation"),
    ("steril", "Sterilization"),
    ("software", "Software"),
    ("cyber", "Cybersecurity"),
    ("emcemi", "EMC / EMI"),
    ("emc", "EMC / EMI"),
    ("electrical", "Electrical Safety"),
    ("risk", "Risk Analysis"),
    ("qsub", "Q-Sub Traceability"),
    ("pre-sub", "Pre-Submission"),
    ("presub", "Pre-Submission"),
    ("meetingminutes", "Meeting Minutes"),
    ("meeting_minutes", "Meeting Minutes"),
    ("feedback", "FDA Feedback"),
    ("emfb", "FDA Feedback"),
    ("aapm", "AAPM Reference"),
    ("predicate", "Predicate Device"),
    ("truthful", "Truthful & Accurate Statement"),
    ("510ksummary", "510(k) Summary"),
    ("indications", "Indications For Use"),
]


def classify_file(filename: str) -> tuple[int, str]:
    """Return (sort_priority, human_label) for a given filename."""
    lower = filename.lower().replace(" ", "")
    for idx, (hint, label) in enumerate(SECTION_ORDER_HINTS):
        if hint in lower:
            return idx, label
    return 999, Path(filename).stem  # fallback: use filename stem as label


def extract_attachments(estar_path: str, work_dir: str, log_fn=None):
    """
    Extract all embedded files from the eSTAR PDF into *work_dir*.
    Returns a list of (filepath, original_name) tuples.
    """
    reader = PdfReader(estar_path)
    extracted = []

    if not reader.attachments:
        raise ValueError("No embedded files found in this PDF.")

    for name, data_list in reader.attachments.items():
        # eSTAR embeds timestamps as separate "attachment" names – skip them
        stripped = name.strip()
        if not stripped:
            continue
        # Heuristic: timestamps look like "2026-01-24T14:56:25"
        if len(stripped) >= 10 and stripped[4] == "-" and stripped[7] == "-":
            try:
                datetime.fromisoformat(stripped)
                continue
            except ValueError:
                pass

        for data in data_list:
            dest = os.path.join(work_dir, stripped)
            with open(dest, "wb") as f:
                f.write(data)
            extracted.append((dest, stripped))
            if log_fn:
                log_fn(f"  Extracted: {stripped}")

    return extracted


def convert_docx_files(file_list: list, work_dir: str, log_fn=None):
    """
    Convert any .docx files to PDF (in-place replacement in the list).
    Uses docx2pdf on Windows (requires MS Word) with a LibreOffice fallback.
    Returns updated list of (filepath, original_name).
    """
    updated = []
    for fpath, orig_name in file_list:
        if fpath.lower().endswith(".docx"):
            pdf_path = fpath.rsplit(".", 1)[0] + ".pdf"
            converted = False

            # Try docx2pdf first (uses MS Word on Windows)
            try:
                from docx2pdf import convert as docx_convert
                if log_fn:
                    log_fn(f"  Converting (Word): {orig_name}")
                docx_convert(fpath, pdf_path)
                converted = True
            except Exception:
                pass

            # Fallback: LibreOffice
            if not converted:
                lo_path = shutil.which("soffice") or shutil.which("libreoffice")
                if lo_path:
                    if log_fn:
                        log_fn(f"  Converting (LibreOffice): {orig_name}")
                    subprocess.run(
                        [lo_path, "--headless", "--convert-to", "pdf",
                         "--outdir", work_dir, fpath],
                        capture_output=True,
                    )
                    converted = os.path.exists(pdf_path)

            if converted and os.path.exists(pdf_path):
                pdf_name = orig_name.rsplit(".", 1)[0] + ".pdf"
                updated.append((pdf_path, pdf_name))
            else:
                if log_fn:
                    log_fn(f"  WARNING: Could not convert {orig_name} – skipping")
        else:
            updated.append((fpath, orig_name))
    return updated


def _create_divider_pdf(title: str, section_num: int, path: str):
    """Create a single-page section divider PDF."""
    doc = SimpleDocTemplate(
        path, pagesize=letter,
        topMargin=2.8 * inch, bottomMargin=1 * inch,
    )
    styles = getSampleStyleSheet()
    section_style = ParagraphStyle(
        "SectionNum", parent=styles["Normal"],
        fontSize=15, leading=20, alignment=1,
        textColor=HexColor("#8899aa"), spaceBefore=30,
    )
    title_style = ParagraphStyle(
        "DivTitle", parent=styles["Title"],
        fontSize=24, leading=30, alignment=1,
        textColor=HexColor("#1a3c5e"), spaceAfter=20,
    )
    story = [
        Paragraph(f"Section {section_num}", section_style),
        Spacer(1, 18),
        Paragraph(title, title_style),
    ]
    doc.build(story)


def _create_toc_pdf(sections: list, path: str, estar_name: str):
    """Create a table-of-contents page listing all sections."""
    doc = SimpleDocTemplate(
        path, pagesize=letter,
        topMargin=1 * inch, bottomMargin=1 * inch,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TOCTitle", parent=styles["Title"],
        fontSize=22, leading=28, alignment=1,
        textColor=HexColor("#1a3c5e"), spaceAfter=6,
    )
    subtitle_style = ParagraphStyle(
        "TOCSub", parent=styles["Normal"],
        fontSize=13, alignment=1,
        textColor=HexColor("#666666"), spaceAfter=28,
    )
    item_style = ParagraphStyle(
        "TOCItem", parent=styles["Normal"],
        fontSize=11, leading=20, leftIndent=40,
        textColor=HexColor("#333333"),
    )

    story = [
        Paragraph(estar_name, title_style),
        Paragraph("Table of Contents", subtitle_style),
    ]
    for i, (_, label) in enumerate(sections, 1):
        story.append(
            Paragraph(f"<b>Section {i}.</b>&nbsp;&nbsp;{label}", item_style)
        )
    doc.build(story)


def build_merged_pdf(
    estar_path: str,
    output_path: str,
    log_fn=None,
    progress_fn=None,
):
    """
    Main pipeline: extract → convert → sort → merge with TOC & bookmarks.
    *log_fn(msg)* receives status strings.
    *progress_fn(fraction)* receives 0.0–1.0 progress updates.
    """

    if log_fn is None:
        log_fn = print
    if progress_fn is None:
        progress_fn = lambda _: None

    work_dir = tempfile.mkdtemp(prefix="estar_")

    try:
        # --- Step 1: Extract ------------------------------------------------
        log_fn("Step 1/4 — Extracting embedded files …")
        progress_fn(0.05)
        file_list = extract_attachments(estar_path, work_dir, log_fn)
        log_fn(f"  → {len(file_list)} files extracted.\n")
        progress_fn(0.20)

        # --- Step 2: Convert DOCX → PDF -------------------------------------
        log_fn("Step 2/4 — Converting .docx files to PDF …")
        file_list = convert_docx_files(file_list, work_dir, log_fn)
        # Keep only PDFs
        file_list = [(f, n) for f, n in file_list if f.lower().endswith(".pdf")]
        log_fn(f"  → {len(file_list)} PDF sections ready.\n")
        progress_fn(0.40)

        # --- Step 3: Sort into 510(k) order ---------------------------------
        log_fn("Step 3/4 — Sorting into 510(k) order …")
        decorated = []
        for fpath, orig_name in file_list:
            priority, label = classify_file(orig_name)
            decorated.append((priority, label, fpath, orig_name))
        decorated.sort(key=lambda x: x[0])
        sections = [(fpath, label) for _, label, fpath, _ in decorated]

        for i, (_, label) in enumerate(sections, 1):
            log_fn(f"  {i:>2}. {label}")
        log_fn("")
        progress_fn(0.50)

        # --- Step 4: Build merged PDF ---------------------------------------
        log_fn("Step 4/4 — Building merged PDF …")

        estar_name = Path(estar_path).stem.replace("_", " ").replace("-", " ")

        toc_path = os.path.join(work_dir, "_toc.pdf")
        _create_toc_pdf(sections, toc_path, estar_name)

        writer = PdfWriter()

        # Add TOC pages
        toc_reader = PdfReader(toc_path)
        for page in toc_reader.pages:
            writer.add_page(page)
        toc_page_count = len(toc_reader.pages)

        total_sections = len(sections)
        page_cursor = toc_page_count  # for bookmarks

        for idx, (fpath, label) in enumerate(sections):
            frac = 0.50 + 0.48 * (idx / total_sections)
            progress_fn(frac)

            # Divider
            div_path = os.path.join(work_dir, f"_div_{idx}.pdf")
            _create_divider_pdf(label, idx + 1, div_path)
            div_reader = PdfReader(div_path)
            for page in div_reader.pages:
                writer.add_page(page)

            # Bookmark at the divider page
            writer.add_outline_item(label, page_cursor)
            page_cursor += len(div_reader.pages)

            # Actual document
            try:
                doc_reader = PdfReader(fpath)
                for page in doc_reader.pages:
                    writer.add_page(page)
                doc_pages = len(doc_reader.pages)
                log_fn(f"  Added: {label} ({doc_pages} pages)")
                page_cursor += doc_pages
            except Exception as exc:
                log_fn(f"  ERROR reading {label}: {exc}")

        # Write output
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with open(output_path, "wb") as f:
            writer.write(f)

        total_pages = len(writer.pages)
        progress_fn(1.0)
        log_fn(f"\nDone!  {total_pages} pages → {output_path}")
        return output_path, total_pages

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CustomTkinter GUI
# ---------------------------------------------------------------------------

class EstarExtractorApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # --- Window setup ---------------------------------------------------
        self.title("eSTAR Package Extractor")
        self.geometry("780x620")
        self.minsize(680, 520)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self._build_ui()
        self._selected_file: str | None = None
        self._output_path: str | None = None
        self._running = False

    # ---- UI construction ---------------------------------------------------

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 8))

        ctk.CTkLabel(
            header, text="eSTAR Package Extractor",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Extract, convert, and merge all embedded files from an FDA eSTAR submission PDF.",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(2, 0))

        # File selection frame
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(fill="x", padx=24, pady=(12, 6))

        ctk.CTkLabel(
            file_frame, text="eSTAR PDF File:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")

        self._file_entry = ctk.CTkEntry(
            file_frame, placeholder_text="Select an eSTAR PDF …", width=480,
        )
        self._file_entry.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="ew")

        self._browse_btn = ctk.CTkButton(
            file_frame, text="Browse", width=100, command=self._browse_file,
        )
        self._browse_btn.grid(row=1, column=1, padx=(4, 12), pady=(0, 12))

        file_frame.columnconfigure(0, weight=1)

        # Output selection
        out_frame = ctk.CTkFrame(self)
        out_frame.pack(fill="x", padx=24, pady=(6, 6))

        ctk.CTkLabel(
            out_frame, text="Output PDF:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")

        self._out_entry = ctk.CTkEntry(
            out_frame, placeholder_text="(defaults to same folder as input)", width=480,
        )
        self._out_entry.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="ew")

        self._out_btn = ctk.CTkButton(
            out_frame, text="Browse", width=100, command=self._browse_output,
        )
        self._out_btn.grid(row=1, column=1, padx=(4, 12), pady=(0, 12))

        out_frame.columnconfigure(0, weight=1)

        # Run button
        self._run_btn = ctk.CTkButton(
            self, text="Extract & Merge", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run,
        )
        self._run_btn.pack(padx=24, pady=(8, 6))

        # Progress bar
        self._progress = ctk.CTkProgressBar(self, width=400)
        self._progress.pack(padx=24, pady=(4, 6))
        self._progress.set(0)

        # Log area
        self._log_box = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Consolas", size=12))
        self._log_box.pack(fill="both", expand=True, padx=24, pady=(6, 8))
        self._log_box.configure(state="disabled")

        # Footer
        ctk.CTkLabel(
            self,
            text="Supports 510(k) eSTAR submissions  •  Auto-sorts sections  •  Converts .docx to PDF",
            font=ctk.CTkFont(size=11),
            text_color="gray",
        ).pack(pady=(0, 12))

    # ---- Callbacks ---------------------------------------------------------

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select eSTAR PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if path:
            self._selected_file = path
            self._file_entry.delete(0, "end")
            self._file_entry.insert(0, path)

            # Auto-fill output path
            if not self._out_entry.get().strip():
                stem = Path(path).stem
                default_out = str(Path(path).parent / f"{stem}_Complete_Package.pdf")
                self._out_entry.delete(0, "end")
                self._out_entry.insert(0, default_out)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if path:
            self._output_path = path
            self._out_entry.delete(0, "end")
            self._out_entry.insert(0, path)

    def _log(self, msg: str):
        """Thread-safe log append."""
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _set_progress(self, value: float):
        self._progress.set(value)

    def _run(self):
        input_path = self._file_entry.get().strip()
        if not input_path or not os.path.isfile(input_path):
            messagebox.showerror("Error", "Please select a valid eSTAR PDF file.")
            return

        output_path = self._out_entry.get().strip()
        if not output_path:
            stem = Path(input_path).stem
            output_path = str(Path(input_path).parent / f"{stem}_Complete_Package.pdf")
            self._out_entry.delete(0, "end")
            self._out_entry.insert(0, output_path)

        if self._running:
            return
        self._running = True
        self._run_btn.configure(state="disabled", text="Processing …")
        self._progress.set(0)

        # Clear log
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        def worker():
            try:
                result_path, total_pages = build_merged_pdf(
                    estar_path=input_path,
                    output_path=output_path,
                    log_fn=lambda msg: self.after(0, self._log, msg),
                    progress_fn=lambda v: self.after(0, self._set_progress, v),
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"Merged PDF created successfully!\n\n"
                    f"{total_pages} pages → {result_path}",
                ))
            except Exception as exc:
                self.after(0, lambda: self._log(f"\nERROR: {exc}"))
                self.after(0, lambda: messagebox.showerror("Error", str(exc)))
            finally:
                self.after(0, self._finish)

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self):
        self._running = False
        self._run_btn.configure(state="normal", text="Extract & Merge")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = EstarExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
