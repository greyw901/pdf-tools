# PDF Tools — Usage Guide

## Getting Started

1. Get the tool — either:
   - clone/download this repo and run `python "PDF Tools.py"`, or
   - download the pre-built Windows `.exe` from [Releases](../../releases) and double-click it.
2. Put the script/exe in the same folder as the files you want to process (PDF, Excel, or Word).
3. Run it and pick a tool number from the menu.

Every tool follows the same flow, and **nothing is processed until you have seen a summary and pressed `[Enter]`**:

1. Pick a tool
2. Pick the files (or pages)
3. Check the summary, `[Enter]` to proceed
4. Read the result: `Outputs` / `Skipped` / `Failed`

## Main Menu

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

  [1-9] - Select a Tool
  [0] - Exit Session
```

Keys in `[ ]` are what you press. Tools 1, 2, 3, 5, 6 work on whole files, picked by number. Tools 4, 7, 8, 9 work on individual pages or sheets, picked by code (`a1-a20`).

## Output Folders

| Tool | Output |
|------|--------|
| 1 - Convert to PDF | `To PDF\` |
| 2 - Merge PDF | `Merged\` |
| 3 - Split PDF | `Split\` |
| 4 - Rotate PDF | `Rotated\` |
| 5 - Protect PDF | `Protected\` |
| 6 - Unprotect PDF | `Unprotected\` |
| 7 - Convert to PDF Precise | `To PDF\` |
| 8 - Selective Combine | `Combined\` |
| 9 - Selective Extract | `Extract\` |

Input files are never modified or deleted. If an output name already exists, `_1`, `_2` ... is added.

## What Each Tool Does

**1 — Convert to PDF** — Converts the Excel (`.xls .xlsx .xlsm .xlsb`) and Word (`.doc .docx`) files you pick to PDF. Excel: all visible sheets in one PDF per file. Word: the whole document. Output: `name.pdf`. Needs Microsoft Excel and Word.

**2 — Merge PDF** — Merges the PDFs you pick into one, in the order you type (`3, 1` merges file 3 first). Enter = all files in list order. Output: `firstfile Merged.pdf`.

**3 — Split PDF** — Splits each PDF you pick into single-page files. Output: `name p1.pdf`, `name p2.pdf`, ...

**4 — Rotate PDF** — Pick pages, then enter direction + angle: `R90`, `L90`, `R180`, `L180`. Picked pages rotate; the rest stay unchanged. Output: `name Rotated.pdf`.

**5 — Protect PDF** — Adds a password to the PDFs you pick. Output: `name_pw_PASSWORD.pdf` — remove the password from the filename before sending the file to anyone else. Files that already have `_pw_` in the name, or are already encrypted, are skipped.

**6 — Unprotect PDF** — Removes the password from the PDFs you pick. Output: `name.pdf` (the `_pw_PASSWORD` part is dropped). Files that were never password-protected are skipped, not treated as failures.

**7 — Convert to PDF Precise** — Like tool 1, but you pick which Excel **sheets** or which Word **pages** to include. Needs Microsoft Excel and Word.

**8 — Selective Combine** — Pick pages from one or more PDFs and combine them into **one** PDF, in the order you type. A page can be used more than once. Output: `Combined 12 pages.pdf`.

**9 — Selective Extract** — Pick pages; each comma-separated item becomes its own PDF. `a2-3` gives one PDF with pages 2-3; a whole file (`b`) gives one PDF per page. Output: `name p3.pdf`, `name p2-3.pdf`, ...

## How Selection Works

Tools 1, 2, 3, 5, 6 list **files** by number:

```
  [1] - "Report.pdf"
  [2] - "Invoice.pdf"
```

Tools 4, 7, 8, 9 list **pages** (Excel files in tool 7 list sheets):

```
  a Pages: [a1-a20] - "Report.pdf"
  b Pages: [b1-b5] - "Invoice.pdf"
```

| Files (1, 2, 3, 5, 6) | Meaning |
|---|---|
| `1` | file 1 |
| `1, 3` | files 1 and 3 |
| `4-6` | files 4, 5 and 6 |

| Pages (4, 7, 8, 9) | Meaning |
|---|---|
| `a1` | single page |
| `a1-a9` | range of pages within the same file |
| `a` | all pages of file a |
| `a1-a9, b, a14` | mix of pages, ranges, and full files |

Every selection prompt looks like:

```
[Selections] - e.g. [a1, b2, c4-6]
[Enter] - Select all
[0] - Cancel & Return to Menu
```

If part of your input isn't recognized:

```
Invalid Input: "xyz"
[Enter] - Skip Invalid Input & Continue
[N] - Re-enter
[0] - Cancel & Return to Menu
```

After picking, a summary is shown:

```
Rotate summary (Pages: 1; Rotation: R90):
  Report.pdf
    a1 - page 1

[Enter] - Proceed
[N] - Re-enter
[0] - Cancel & Return to Menu
```

`[0]` at any prompt cancels the tool and returns directly to the main menu.

## Unreadable Files

Tools 4, 7, 8, 9 need page or sheet counts, so unreadable files are listed once before the file list:

```
Unreadable: 1
   1. "file1.pdf"   :   PdfStreamError - Stream has ended unexpectedly

[Enter] - Skip & Continue
[0] - Cancel & Return to Menu
```

For the other tools an unreadable file shows up under `Failed` after processing.

## Reading the Result

```
Outputs: 2
   1. "Report Rotated.pdf"
   2. "Invoice Rotated.pdf"

Skipped: 1
   1. "Plain.pdf"   :   no password

Failed: 1
   1. "Broken.pdf"   :   ValueError - wrong password, could not decrypt

Elapsed: 0.4s
```

- **Skipped** — nothing to do for that file; not a problem.
- **Failed** — something went wrong; the reason follows the colon.

## After a Tool Finishes

```
[Enter] - Continue to Menu
[0] - Exit Session
```

When something failed, the second line becomes `[0] - Save error log and Exit Session`. It writes `error log_YYYYMMDD.txt` next to the exe with the step, error, context and traceback for every failure.

## Requirements

- Tools 1 and 7: Microsoft Excel and Microsoft Word must be installed.
- Tools 2, 3, 4, 5, 6, 8, 9: no additional software required.

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| Console closes immediately | Input files missing from folder — place files in the same folder as the script/exe |
| Tool 1 or 7 fails | Microsoft Office not installed, or an Office process is stuck — close any lingering EXCEL.EXE / WINWORD.EXE |
| File listed under "Failed" | Read the reason after the colon; save the error log with `[0]` |
| Wrong password on Unprotect | Check the filename for the `_pw_PASSWORD` hint |
| File listed under "Skipped" | Nothing to do for that file; the reason is shown after the colon |
| Antivirus blocks the `.exe` | Unsigned binary — whitelist it with your antivirus/IT if needed |
| Windows says "Windows protected your PC" | SmartScreen flag on unsigned executables — click "More info" then "Run anyway" |

---

## For Developers

### Single-file layout (`PDF Tools.py`, top to bottom)

| Block | Contents |
|---|---|
| Constants | `VERSION`, `LINE_WIDTH`, `MAX_LISTED`, extensions, `PW_MARK`, `OUT_FOLDERS` |
| Signals | `UserCancelled`, `NoInput`, `ExitSession`, `SkipItem` |
| Paths / naming | `get_script_dir`, `get_unique_path`, `list_files`, `prefix_generator`, `dedupe` |
| File index + selection | `Entry`, `Token`, `Index`, `scan_pdf_index`, `parse_tokens`, `prompt_selection`, `confirm`, `prompt_password`, `unit_summary_lines` |
| Processing + report | `run_items`, `begin`, `finish`, `tail`, `write_error_log`, `make_failure` |
| Office helpers | `OfficeApps`, `excel_export`, `word_export`, `scan_office_index` |
| Tools | nine thin `tool_xxx()` functions |
| Menu | `MENU`, `TOOLS`, `run_tool`, `main` |

Logic used by more than one tool lives in the shared blocks. A `tool_xxx()` holds only its own prompts and its worker.

### The same flow in every tool

```
scan files once -> prompt_selection -> (tool prompt) -> confirm
  -> announce folders -> begin("Gerund") -> run_items -> finish
```

- Selection level is `"file"` (tools 1, 2, 3, 5, 6) or `"unit"` (4, 7, 8, 9). One parser, one set of prompts.
- The scan happens **once**, before the selection loop; `[N]` re-enter never re-scans or re-shows `Unreadable`.
- The `confirm()` summary doubles as the run report's `Inputs` block (on screen before processing, lists what will be processed).
- Workers return output names and never print. `run_items` is the only loop; it files exceptions into `Failed` / `Skipped`.
- `begin()` is the statement directly before `run_items`, so the last line on screen always names the running stage.

### Cancel / exit architecture (exceptions, not return values)

| Exception | Raised when | Caught in |
|---|---|---|
| `UserCancelled` | `[0]` at any prompt | `run_tool()` → back to menu |
| `NoInput` | no files found for the tool | `run_tool()` → normal tail |
| `ExitSession` | `[0]` at the tail | `main()` → close |
| `SkipItem` | worker: nothing to do for the item | `run_items()` |

This replaces the v1.x `"CANCELLED"` string return, which every caller had to remember to check. New mid-tool exit point: `raise UserCancelled()`, never `return`. An unexpected exception inside a tool is caught in `run_tool()`, shown as `FAILED`, and goes to the normal tail so the log can be saved.

### Error log

Written only when `[0]` is pressed at the tail. `make_failure()` calls `traceback.format_exc()` **inside** the `except` block — outside it the traceback is gone. Fields per entry: Step, Error, Context, Traceback. Filename `error log_YYYYMMDD.txt`, same-day re-runs get `_1`, `_2`.

### Output naming

All outputs are deliverables: no run date, words separated by spaces, underscore only for the collision counter. One deliberate exception: Protect PDF keeps `name_pw_PASSWORD.pdf` so the password hint travels with the file; Unprotect strips it. The `_unprotected` suffix and its skip check were removed — Unprotect skips by `reader.is_encrypted`.

### Tool notes

- **Merge** — order is the order typed; output stem is the first file that merged successfully.
- **Protect** — skips when the stem contains `_pw_` or the file is already encrypted. Passwords containing `\ / : * ? " < > |` are rejected (they go into the filename).
- **Unprotect** — not encrypted → Skipped; `decrypt() == 0` → Failed.
- **Convert to PDF Precise** — Excel: unselected visible sheets are hidden, then exported (workbook closed unsaved). Word has no multi-range export, so the whole document is exported to a temp PDF and the picked pages are copied out with pypdf. Page counts come from Word (`ComputeStatistics`); if Word and the PDF ever disagree on pagination, that file lands in `Failed`.
- **Selective Combine** — pages keep typed order and repeats.
- **Selective Extract** — a whole-file token expands to one job per page.
- Office lock files (`~$name.xlsx`) are ignored when listing. Lists longer than `MAX_LISTED` (20) print `... and n more`.

### Path Resolution

```python
script_dir = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, 'frozen', False) else __file__
))
```

`.py` mode resolves to the script's location; `.exe` mode resolves to the exe's location. Don't switch this to `os.getcwd()` or `sys.argv[0]`.

### Versioning

Format: `v{major}.{minor}_{YYYYMMDD}`. Current: `v2.0_20261108`. Update the `VERSION` constant in `PDF Tools.py`, `version_info.txt`, and this file.

### Packaging

```bash
pip install pyinstaller
pyinstaller --onefile --version-file=version_info.txt "PDF Tools.py"
```

`version_info.txt` embeds author/version info into the exe's Windows Properties > Details tab — this is what survives if the `.exe` ever gets forwarded on its own, with no source code or docs attached.

Output: `dist/PDF Tools.exe`. The `build/` folder can be deleted after packaging; keep the generated `.spec` file for faster repeat builds (`pyinstaller "PDF Tools.spec"`).

If tools 1 or 7 crash in the `.exe` but work fine as a `.py` file, add this to `hiddenimports` in the `.spec` file:

```python
['win32com', 'win32com.client', 'win32com.server', 'pywintypes']
```

### Dependencies

| Library | Used by | Install |
|---|---|---|
| pypdf | Tools 2, 3, 4, 5, 6, 7 (Word pages), 8, 9 | `pip install pypdf` |
| pywin32 | Tools 1, 7 | `pip install pywin32` |
| pyinstaller | Packaging only | `pip install pyinstaller` |

Python 3.8+. Microsoft Excel and Word required on the machine running the script/exe, for Tools 1 and 7 only.

### Known Limitations

- No code signing — antivirus may flag the `.exe`.
- Tools 1 and 7 require Microsoft Office on the machine running the tool.
- No auto-update — distribute new versions manually.
- Rotate PDF applies one direction + angle per run. For mixed angles, run the tool again on the output.
- The password prompt is visible by design.

---
**Version:** v2.0_20261108

**Author:** Yeoh Huey Sun
**Email:** greyw901@gmail.com
**GitHub:** https://github.com/greyw901
