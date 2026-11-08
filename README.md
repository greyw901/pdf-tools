# PDF Tools

A free, local, single-file, menu-driven command-line utility for everyday PDF work — convert, merge, split, rotate, protect, unprotect, and selectively combine or extract pages.

## Why this exists

Web tools like iLovePDF are convenient, but they mean uploading your file to a third-party server first — not ideal for anything sensitive (bank statements, invoices, contracts). PDF Tools does the same everyday jobs entirely on your own machine.

- **Local** — files never leave your computer, no account, no upload.
- **Fast** — [pypdf](https://pypdf.readthedocs.io/) manipulates a PDF's internal object structure directly instead of re-rendering pages, so merge/split/rotate/extract are near-instant even on large files.
- **Lightweight** — one script or one `.exe`. No Adobe Acrobat, no browser tab.
- **No surprises** — every tool shows you what it is about to process and waits for your confirmation. Nothing runs on "everything in the folder" behind your back.

## How it works

Every tool follows the same four steps:

1. Pick a tool from the menu
2. Pick the files (or pages) — by number for whole files, by code like `a1-a20` for pages and sheets
3. Check the summary and press `[Enter]` to proceed
4. Read the result: `Outputs`, `Skipped`, `Failed` (with the reason)

```
  +-------------------------------+-------------------------------+
  |   PDF Tools v2.0_20261108     |                               |
  +-------------------------------+-------------------------------+
  | BASIC                         | ADVANCE                       |
  +-------------------------------+-------------------------------+
  | [1] - Convert to PDF          | [7] - Convert to PDF Precise  |
  | [2] - Merge PDF               | [8] - Selective Combine       |
  | [3] - Split PDF               | [9] - Selective Extract       |
  | [4] - Rotate PDF              |                               |
  | [5] - Protect PDF             |                               |
  | [6] - Unprotect PDF           |                               |
  +-------------------------------+-------------------------------+
  | [0] - Exit                                                    |
  +---------------------------------------------------------------+
```

Anything in `[ ]` is a key you press. `[0]` always cancels and takes you back.

## Tools

| # | Tool | What it does |
|---|------|---------------|
| 1 | Convert to PDF | Converts the Excel/Word files you pick to PDF (all visible sheets) |
| 2 | Merge PDF | Merges the PDFs you pick into one, in the order you type |
| 3 | Split PDF | Splits the PDFs you pick into single-page files |
| 4 | Rotate PDF | Rotates selected pages right/left by 90 / 180° (`R90`, `L90`, `R180`, `L180`) |
| 5 | Protect PDF | Adds password protection to the PDFs you pick |
| 6 | Unprotect PDF | Removes password protection |
| 7 | Convert to PDF Precise | Like #1, but you pick which Excel sheets or Word pages to include |
| 8 | Selective Combine | Picks specific pages across multiple PDFs into one output, in the order you type |
| 9 | Selective Extract | Pulls specific pages out of PDFs as separate files |

A plain-text copy of the guide, for people who just want the .exe, is in [PDF Tools Instruction.txt](PDF%20Tools%20Instruction.txt).

Full walkthrough — selection syntax (`1, 3, 4-6` for files, `a1, b2, c4-6` for pages), screens, and troubleshooting — is in [INSTRUCTIONS.md](INSTRUCTIONS.md).

## Requirements

- Python 3.8+
- `pip install pypdf`
- `pip install pywin32` — only needed for tools **1** and **7**

## Before you run

1. Create a folder anywhere (e.g. C:\PDF Tools\)
2. Put PDF Tools.py in that folder
3. Put the files you want to process in the same folder

## Run

```bash
python "PDF Tools.py"
```

or

Double-click PDF Tools.py

## Pre-built Windows .exe

Or grab a pre-built Windows `.exe` from the [Releases](../../releases) page — no Python install needed. Put it in the same folder as your files and double-click.

Windows may show a SmartScreen warning on first run — click "More info" then "Run anyway".

## Platform support

| Tools | Engine | Platform |
|-------|--------|----------|
| 2, 3, 4, 5, 6, 8, 9 | pypdf | Windows / macOS / Linux |
| 1, 7 | Drives a locally installed copy of Microsoft Excel/Word via COM automation | Windows only, with Office installed |

## Output

Every tool writes to its own subfolder next to the script (`Merged\`, `Split\`, `Rotated\`, etc.). Original input files are never modified or deleted. If an output name already exists, `_1`, `_2` ... is added instead of overwriting.

## When something fails

A failed file never stops the others. It is listed under `Failed:` with the reason, and choosing `[0] - Save error log and Exit Session` at the end writes `error log_YYYYMMDD.txt` next to the tool — step, error, file path and full traceback, so the problem can be diagnosed without re-describing it.

## Note on Protect PDF

The password you set is included in the output filename (e.g. `Report_pw_ABCd123.pdf`) so you can see it at a glance. Strip it from the filename before sending the file to anyone else. Unprotect PDF removes the `_pw_...` part again when it writes the unlocked copy.

## What's new in v2.0

- Every tool now asks which files (or pages) to use and shows a summary before anything runs.
- Whole files are picked by number (`[1]`, `[2]` ...), pages and sheets by code (`a1-a20`).
- Rotate takes direction + angle (`R90`, `L180`).
- Convert to PDF Precise now works on Word pages as well as Excel sheets.
- "Combine PDF Pages" and "Extract PDF Pages" are now **Selective Combine** and **Selective Extract**.
- Consistent result report across all tools, with the failure reason for each file and a saveable error log.
- Output names use spaces instead of underscores (`Report Merged.pdf`, `Report p3.pdf`).

---

**Author:** Yeoh Huey Sun <br>
**Email:** greyw901@gmail.com <br>
**GitHub:** https://github.com/greyw901
