# eSTAR Package Extractor

A desktop tool that extracts all embedded files from an FDA **eSTAR (electronic Submission Template And Resource)** PDF, converts any `.docx` attachments to PDF, sorts everything into standard 510(k) section order, and merges it all into a single navigable PDF.

The FDA's eSTAR format bundles dozens of documents as embedded attachments inside a single PDF — which most viewers can't display inline. This tool unpacks the entire submission into one readable, bookmarked PDF you can scroll through or share for review.

## Features

- **One-click extraction** of all embedded files from an eSTAR PDF
- **Automatic .docx → PDF conversion** (via Microsoft Word or LibreOffice)
- **Smart section ordering** — recognizes common 510(k) naming conventions and sorts documents into a logical review sequence (cover letter → executive summary → device info → SE → specs → labeling → testing → references)
- **Table of contents** page listing every section
- **Section divider pages** between documents for easy visual navigation
- **PDF sidebar bookmarks** for instant jump-to-section
- **CustomTkinter GUI** with progress bar and real-time log output
- **Auto-installs Python dependencies** on first run

## Screenshots

*Launch the app → browse to an eSTAR PDF → click Extract & Merge → done.*

## Requirements

- **Python 3.8+** — [download here](https://www.python.org/downloads/) (check "Add Python to PATH" during install)
- **Microsoft Word** (recommended) or **LibreOffice** — needed only if the eSTAR contains `.docx` attachments

Python packages are installed automatically on first run:

| Package | Purpose |
|---------|---------|
| `pypdf` | Read/write PDFs, extract embedded files |
| `pdfplumber` | PDF text extraction |
| `reportlab` | Generate TOC and divider pages |
| `customtkinter` | Modern GUI |
| `docx2pdf` | Word → PDF conversion |

## Quick Start

1. Clone or download this repo
2. Double-click **`Run_eSTAR_Extractor.bat`**
3. Browse to your eSTAR PDF
4. Choose an output location (or accept the default)
5. Click **Extract & Merge**

The merged PDF will appear at your chosen output path.

### Running from the command line

```bash
python estar_extractor.py
```

## How It Works

1. **Extract** — reads all embedded file attachments from the eSTAR PDF (skipping timestamp metadata entries)
2. **Convert** — any `.docx` files are converted to PDF using `docx2pdf` (MS Word) with a LibreOffice fallback
3. **Classify & Sort** — each filename is matched against a table of ~40 common 510(k) naming patterns (e.g., `EXEC-SUM`, `ADI`, `SE`, `BIOCOMP`, `HFE`, `BENCH`) and assigned a section priority; unrecognized files go at the end
4. **Merge** — a table of contents is generated, section divider pages are inserted before each document, all pages are combined into a single PDF, and sidebar bookmarks are added for navigation

## Supported Section Types

The classifier recognizes these patterns in filenames (case-insensitive):

| Priority | Pattern keywords | Label |
|----------|-----------------|-------|
| 1 | `coverletter` | Cover Letter |
| 2 | `exec-sum`, `executivesummary` | Executive Summary |
| 3 | `adi` | Applicant Device Information |
| 4 | `se`, `substantial` | Substantial Equivalence |
| 5 | `spec-draw` | Specifications & Drawings |
| 6 | `sys-comp` | System Component List |
| 7–9 | `pkg-lbl`, `pkg`, `lbl` | Packaging / Labeling |
| 10 | `instruction`, `ifu` | Instructions For Use |
| 11 | `sl`, `standards` | Standards List |
| 12 | `biocomp` | Biocompatibility |
| 13 | `bench`, `performance` | Bench / Performance Testing |
| 14 | `hfe`, `human factor` | Human Factors Engineering |
| 15+ | `coa`, `aca`, `steril`, `software`, `cyber`, `emc`, `risk`, etc. | Various |
| Last | `qsub`, `presub`, `meetingminutes`, `feedback`, `aapm` | Pre-sub / Reference docs |

Files that don't match any pattern are placed at the end, in the order they were embedded.

## File Structure

```
estar-extractor/
├── estar_extractor.py          # Main application (GUI + extraction logic)
├── Run_eSTAR_Extractor.bat     # Windows launcher
└── README.md                   # This file
```

## Troubleshooting

**"Python was not found"** — Install Python 3.8+ from python.org and ensure "Add Python to PATH" is checked.

**DOCX files are skipped** — Install Microsoft Word, or install LibreOffice and ensure `soffice` is on your system PATH.

**"No embedded files found"** — The PDF may not be an eSTAR submission, or it may use a non-standard embedding method. Open it in Adobe Acrobat to verify attachments are present.

**Permission errors on pip install** — Run the batch file as Administrator, or manually install packages: `pip install pypdf pdfplumber reportlab customtkinter docx2pdf`

## License

MIT
