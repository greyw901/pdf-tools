# PDF Tools — Usage Guide

## Getting Started

1. Get the tool — either:
   - clone/download this repo and run `python "PDF Tools.py"`, or
   - download the pre-built Windows `.exe` from [Releases](../../releases) and double-click it.
2. Put the script/exe in the same folder as the files you want to process (PDF, Excel, or Word).
3. Run it and pick a tool number from the menu.

## Main Menu

```
  +---------------------------+-----------------------------+
  |   PDF Tools v1.0_20260615 |                             |
  +---------------------------+-----------------------------+
  | BASIC                     | ADVANCE                     |
  +---------------------------+-----------------------------+
  | 1 - Convert to PDF        | 7 - Convert to PDF_Precise  |
  | 2 - Merge PDF             | 8 - Combine PDF Pages       |
  | 3 - Split PDF             | 9 - Extract PDF Pages       |
  | 4 - Rotate PDF            |                             |
  | 5 - Protect PDF           |                             |
  | 6 - Unprotect PDF         |                             |
  +---------------------------+-----------------------------+
  | 0 - Exit                                                |
  +---------------------------------------------------------+

  [Number] - Select a Tool
  [0] - Exit Session
```

BASIC tools run automatically, no selection needed. ADVANCE tools are interactive — you choose what to process.

## Output Folders

| Tool | Output |
|------|--------|
| 1 - Convert to PDF | `To PDF\` subfolder |
| 2 - Merge PDF | `Merged\` subfolder |
| 3 - Split PDF | `Split\` subfolder |
| 4 - Rotate PDF | `Rotated\` subfolder |
| 5 - Protect PDF | `Protected\` subfolder |
| 6 - Unprotect PDF | `Unprotected\` subfolder |
| 7 - Convert to PDF_Precise | `To PDF\` subfolder |
| 8 - Combine PDF Pages | `Combined\` subfolder |
| 9 - Extract PDF Pages | `Extract\` subfolder |

Input files are never modified or deleted.

## What Each Tool Does

**1 — Convert to PDF**
Converts all Excel (`.xls .xlsx .xlsm .xlsb`) and Word (`.doc .docx`) files in the folder to PDF. Excel: all visible sheets combined into one PDF per file. Word: entire document exported as PDF. Requires Microsoft Excel and Word to be installed.

**2 — Merge PDF**
Combines all PDFs in the folder into one, in alphabetical (A-Z) order by filename.
Output: `firstfile_merged.pdf`

**3 — Split PDF**
Splits every PDF into individual single-page files.
Output: `filename_1.pdf`, `filename_2.pdf`, ...

**4 — Rotate PDF**
Select pages to rotate and choose an angle (90/180/270). Selected pages rotate; all other pages stay unchanged.
Output: `filename_rotated.pdf`

**5 — Protect PDF**
Adds password protection to all PDFs in the folder. You'll be prompted for a password.
Output: `filename_pw_PASSWORD.pdf` — remove the password from the filename before sending the file to anyone else.

**6 — Unprotect PDF**
Removes password protection from all PDFs. You'll be prompted for the password.
Output: `filename_unprotected.pdf`. Files that were never password-protected are skipped automatically — not treated as failures.

**7 — Convert to PDF_Precise**
Same as #1, but lets you choose which sheets to include for each Excel file before converting. Requires Microsoft Excel and Word.

**8 — Combine PDF Pages**
Select specific pages from multiple PDFs and combine them into one PDF in your chosen order. Press Enter with no input to combine all files in full (works as a general merge tool too).

**9 — Extract PDF Pages**
Select specific pages to extract from PDFs. Each comma-separated token becomes its own output file. Press Enter with no input to split everything into individual pages.

## How Selection Works (Tools 4, 7, 8, 9)

Files or sheets are listed with codes when the tool starts:

```
a (Pages: a1-a20) - Report.pdf
b (Pages: b1-b5)  - Invoice.pdf
```

Token syntax:

| Input | Meaning |
|-------|---------|
| `a1` | single page |
| `a1-a9` | range of pages within the same file |
| `a` | all pages of file a |
| `a1-a9, b, a14` | mix of pages, ranges, and full files |

(Tool 7's sheet selector uses the same format with plain numbers instead of page codes, e.g. `1, 2, 4-6`.)

Every selection prompt looks like:

```
[Selections] - e.g. a1, b2, c4-6
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

Choosing `0` at any of these steps cancels the current tool and returns you directly to the main menu.

## Unreadable Files

If a file can't be opened (corrupted, encrypted, etc.), all unreadable files are listed together up front:

```
Unreadable PDF:
"file1.pdf"
"file2.pdf"

[Enter] - Skip & Continue
[0] - Cancel & Return to Menu
```

## Skipped vs. Failed Files

Some files are skipped automatically — not a failure, just nothing to do for that file:

- Skipped (already protected) — Tool 5
- Skipped (already unprotected) — Tool 6
- Skipped (no password) — Tool 6

Genuine failures (corrupted file, wrong password, etc.) are listed separately at the end under `Failed:`, one item per line.

## After a Tool Finishes

```
[Enter] - Continue to Menu
[0] - Exit Session
```

## Requirements

- Tools 1 and 7: Microsoft Excel and Microsoft Word must be installed.
- Tools 2, 3, 4, 5, 6, 8, 9: no additional software required.

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| Console closes immediately | Input files missing from folder — place files in the same folder as the script/exe |
| Tool 1 or 7 fails | Microsoft Office not installed, or an Office process is stuck — close any lingering Excel.exe / WINWORD.exe |
| File listed under "Failed" | File corrupted, password-protected, or open in another program |
| Wrong password on Unprotect | Check the filename for the `_pw_PASSWORD` hint |
| File listed under "Skipped (no password)" | Tool 6 only — this file was never password-protected; nothing to do |
| Antivirus blocks the `.exe` | Unsigned binary — whitelist it with your antivirus/IT if needed |

---

## For Developers

### Master Script Structure

| Function | Description |
|---|---|
| `get_script_dir()` | Path resolution for `.py` / `.exe` |
| `get_unique_path()` | Naming collision handler |
| `get_pdf_files()` | Lists `.pdf` files in folder |
| `prefix_generator()` | Yields `a, b, ..., z, aa, ab...` |
| `tool_convert_to_pdf()` | Tool 1 |
| `tool_merge_pdf()` | Tool 2 |
| `tool_split_pdf()` | Tool 3 |
| `tool_protect_pdf()` | Tool 5 |
| `tool_unprotect_pdf()` | Tool 6 |
| `tool_convert_to_pdf_precise()` | Tool 7 |
| `tool_combine_pdf_pages()` | Tool 8 |
| `tool_extract_pdf_pages()` | Tool 9 |
| `tool_rotate_pdf()` | Tool 4 |
| `MENU` | Menu display string |
| `TOOLS` | Dict: number → (name, desc, function) |
| `main()` | Menu loop |

### Path Resolution

```python
script_dir = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, 'frozen', False) else __file__
))
```

`.py` mode resolves to the script's location; `.exe` mode resolves to the exe's location. Don't switch this to `os.getcwd()` or `sys.argv[0]`.

### Cancel / Exit Architecture

Two distinct outcomes are tracked separately:

- **`"CANCELLED"`** — any `tool_xxx()` function returns this string when the user chooses `0` at a mid-process prompt. `main()` checks for it and skips the "Continue to Menu / Exit Session" prompt, looping straight back to the menu.
- **Normal completion** — `tool_xxx()` returns `None`. `main()` then shows the Continue/Exit prompt, where `0` is a genuine `sys.exit(0)`.

Any new mid-tool exit point added to a tool function should also return `"CANCELLED"`, not just `return` or `break` — otherwise cancel will look identical to "completed" to the user.

### File Scanning Pattern (Tools 4, 8, 9)

Rotate PDF, Combine PDF Pages, and Extract PDF Pages all scan the folder's PDFs **exactly once** before entering the selection loop. Unreadable files are collected and shown in a single prompt — not re-triggered on every loop iteration (e.g. when the user presses `N` to re-enter a selection). Keep the scan (`code_map` / `prefix_page_map` / `display_lines`) outside the `while` loop if extending these tools.

### Selection Prompt Convention

Every selection-entry prompt (Tools 4, 7, 8, 9) uses the same three-line shape:

```
[Selections] - e.g. <context-appropriate example>
[Enter] - Select all
[0] - Cancel & Return to Menu
```

Only the example after "e.g." changes (page codes for PDF-page tools, plain numbers for Tool 7's sheet selector).

### Summary Line Conventions

Use "Pages:" / "Output files:" not the grammatically-awkward "N page(s)" / "N output file(s)":

```
Selection summary (Pages: N):
Extract summary (Output files: N):
Rotate summary (Pages: N; Rotation: <angle>):
```

Skipped/Failed/Unreadable lists never comma-join multiple items onto one line — one heading, then one item per line, quoted:

```
Failed:
  "file1.pdf"
  "file2.pdf"
```

### Output Reporting Conventions

Tools with no pre-confirm summary (Merge, Split, Protect, Unprotect): list source filenames being processed, then after processing completes, print `Output: <location>` on its own line, followed by the output filename(s) each on their own line — never on the same line as `Output:`.

Exception: Split only prints the output folder (no filename list — splitting can produce dozens of files). Extract behaves the same way for the same reason.

### Skipped vs. Failed — Three-Way Branch in Tool 6 (Unprotect)

Checked in this order, per file:

1. Filename ends with `_unprotected` → Skipped (already unprotected) — no further processing
2. `reader.is_encrypted` is `False` → Skipped (no password) — no output written, **not** a failure
3. `reader.is_encrypted` is `True` → attempt `decrypt()`; wrong password raises `ValueError` → Failed; correct password → rebuild and write to Output

Four separate buckets (`already_ok`, `no_password`, `success`, `errors`) are tracked and printed under four separate headings — don't merge "no password to begin with" into the same bucket as "failed to unlock."

### Versioning

Format: `v{major}.{minor}_{YYYYMMDD}`. Current: `v1.0_20260615`. Update in the `MENU` string inside `PDF Tools.py`.

### Packaging

```bash
pip install pyinstaller
pyinstaller --onefile "PDF Tools.py"
```

Output: `dist/PDF Tools.exe`. The `build/` folder can be deleted after packaging; keep the generated `.spec` file for faster repeat builds (`pyinstaller "PDF Tools.spec"`).

If tools 1 or 7 crash in the `.exe` but work fine as a `.py` file, add this to `hiddenimports` in the `.spec` file:

```python
['win32com', 'win32com.client', 'win32com.server', 'pywintypes']
```

### Dependencies

| Library | Used by | Install |
|---|---|---|
| pypdf | Tools 2, 3, 4, 5, 6, 8, 9 | `pip install pypdf` |
| pywin32 | Tools 1, 7 | `pip install pywin32` |
| pyinstaller | Packaging only | `pip install pyinstaller` |

Python 3.8+. Microsoft Excel and Word required on the machine running the script/exe, for Tools 1 and 7 only.

### Known Limitations

- No code signing — antivirus may flag the `.exe`.
- Tools 1 and 7 require Microsoft Office on the machine running the tool.
- No auto-update — distribute new versions manually.
- Rotate PDF (Tool 4) applies one angle per run. For mixed angles, run the tool again on the output.

---

**Developed by:** Yeoh Huey Sun
**Version:** v1.0_20260615
