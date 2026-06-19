# PDF Tools

A single-file, menu-driven command-line utility for everyday PDF work — convert, merge, split, rotate, protect, unprotect, combine, and extract pages.

## Why this exists

Web tools like iLovePDF are convenient, but they mean uploading your file to a third-party server first — not ideal for anything sensitive (bank statements, invoices, contracts). PDF Tools does the same everyday jobs entirely on your own machine.

- **Local** — files never leave your computer, no account, no upload.
- **Fast** — [pypdf](https://pypdf.readthedocs.io/) manipulates a PDF's internal object structure directly instead of re-rendering pages, so merge/split/rotate/extract are near-instant even on large files.
- **Lightweight** — one script or one `.exe`. No Adobe Acrobat, no browser tab.

## Tools

| # | Tool | What it does |
|---|------|---------------|
| 1 | Convert to PDF | Converts all Excel/Word files in the folder to PDF (all visible sheets) |
| 2 | Merge PDF | Combines all PDFs in the folder into one, alphabetical order |
| 3 | Split PDF | Splits every PDF into individual single-page files |
| 4 | Rotate PDF | Rotates selected pages by 90 / 180 / 270° |
| 5 | Protect PDF | Adds password protection to all PDFs in the folder |
| 6 | Unprotect PDF | Removes password protection |
| 7 | Convert to PDF_Precise | Same as #1, but lets you choose which Excel sheets to include |
| 8 | Combine PDF Pages | Pick and reorder specific pages across multiple PDFs into one output |
| 9 | Extract PDF Pages | Pulls specific pages out of PDFs as separate files |

Full walkthrough — selection syntax (`a1, b2, c4-6`), menu screenshots, and troubleshooting — is in [INSTRUCTIONS.md](INSTRUCTIONS.md).

## Requirements

- Python 3.8+
- `pip install pypdf`
- `pip install pywin32` — only needed for tools **1** and **7**

## Run

```bash
python "PDF Tools.py"
```

## Pre-built Windows .exe 

Or grab a pre-built Windows `.exe` from the [Releases](../../releases) page — no Python install needed.

Windows may show a SmartScreen warning on first run — click "More info" then "Run anyway".

## Platform support

| Tools | Engine | Platform |
|-------|--------|----------|
| 2, 3, 4, 5, 6, 8, 9 | pypdf | Windows / macOS / Linux |
| 1, 7 | Drives a locally installed copy of Microsoft Excel/Word via COM automation | Windows only, with Office installed |

## Output

Every tool writes to its own subfolder next to the script (`Merged\`, `Split\`, `Rotated\`, etc.). Original input files are never modified or deleted.

## Note on Protect PDF

The password you set is included in the output filename (e.g. `Report_pw_ABCd123.pdf`) so you can see it at a glance. Strip it from the filename before sending the file to anyone else.

---

**Author:** Yeoh Huey Sun <br>
**Email:** greyw901@gmail.com <br>
**GitHub:** https://github.com/greyw901
