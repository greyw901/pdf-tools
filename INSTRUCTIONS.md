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
3. Check the summary, `[Enter]` to continue
4. Read the result: `Outputs` / `Skipped` / `Failed`

## Main Menu

```
========================================================
PDF Tools
========================================================

+-------------------------------+-------------------------------+
| BASIC                         | ADVANCE                       |
+-------------------------------+-------------------------------+
| [1] Convert to PDF            | [7] Convert to PDF Precise    |
| [2] Merge PDF                 | [8] Selective Combine         |
| [3] Split PDF                 | [9] Selective Extract         |
| [4] Rotate PDF                |                               |
| [5] Protect PDF               |                               |
| [6] Unprotect PDF             |                               |
+-------------------------------+-------------------------------+

--------------------------------------------------------
[#] Select
[0] Exit

Ans:
```

The window is cleared for every screen. Every screen has the same shape: the tool name between `=` lines, what you have already chosen, the question and its choices, a `-` line, the keys you can press, a blank line, then the prompt where you type. Keys in `[ ]` are what you press. Tools 1, 2, 3, 5, 6 work on whole files, picked by number. Tools 4, 7, 8, 9 work on individual pages or sheets, picked by code (`a1-a20`).

## Where the Output Goes

Every tool writes its result **next to the input files**, in the same folder as the exe. The exceptions are **Split PDF** (`Split\`) and **Selective Extract** (`Extract\`), which can produce many files at once.

Input files are never modified or deleted. If an output name already exists, `_1`, `_2` ... is added.

Because results land in the same folder, they show up in the file list the next time you run a tool (for example `Report_merged.pdf`). Move finished files out when you no longer want them listed.

## What Each Tool Does

**1 — Convert to PDF** — Converts the Excel (`.xls .xlsx .xlsm .xlsb`) and Word (`.doc .docx`) files you pick to PDF. Excel: all visible sheets in one PDF per file. Word: the whole document. Output: `name.pdf`. Needs Microsoft Excel and Word.

**2 — Merge PDF** — Merges the PDFs you pick into one, in the order you type (`3, 1` merges file 3 first). Enter = all files in list order. Output: `firstfile_merged.pdf`.

**3 — Split PDF** — Splits each PDF you pick into single-page files. Output: `name p1.pdf`, `name p2.pdf`, ...

**4 — Rotate PDF** — Pick pages, then enter direction + angle: `R90`, `L90`, `R180`, `L180`. Picked pages rotate; the rest stay unchanged. Output: `name_rotated.pdf`.

**5 — Protect PDF** — Adds a password to the PDFs you pick. Output: `name_pw_PASSWORD.pdf` — remove the password from the filename before sending the file to anyone else. Files that already have `_pw_` in the name, or are already encrypted, are skipped.

**6 — Unprotect PDF** — Removes the password from the PDFs you pick. Output: `name.pdf` (the `_pw_PASSWORD` part is dropped). Files that were never password-protected are skipped, not treated as failures.

**7 — Convert to PDF Precise** — Like tool 1, but you pick which Excel **sheets** or which Word **pages** to include. Needs Microsoft Excel and Word.

**8 — Selective Combine** — Pick pages from one or more PDFs and combine them into **one** PDF, in the order you type. A page can be used more than once. Output: `Combined 12 pages.pdf`.

**9 — Selective Extract** — Pick pages; each comma-separated item becomes its own PDF. `a2-3` gives one PDF with pages 2-3; a whole file (`b`) gives one PDF per page. Output: `name p3.pdf`, `name p2-3.pdf`, ...

## How Selection Works

Tools 1, 2, 3, 5, 6 list **files** by number:

```
   [1] "Report.pdf"
   [2] "Invoice.pdf"
```

Tools 4, 7, 8, 9 list **pages** (Excel files in tool 7 list sheets):

```
   a Pages: [a1-a5] - "Invoice.pdf"
   b Pages: [b1-b20] - "Report.pdf"
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

A selection screen looks like this:

```
========================================================
Rotate PDF
========================================================

Select pages to rotate

e.g. a1, b2, c4-6

   a Pages: [a1-a5] - "Invoice.pdf"
   b Pages: [b1-b20] - "Report.pdf"

--------------------------------------------------------
[Enter] Select all
[#] Select
[0] Back to menu
[00] Cancel & exit

Pages:
```

If any part of your input isn't recognized, the whole entry is rejected (nothing is half-accepted) and the same screen comes back with the reason above the question:

```
========================================================
Rotate PDF
========================================================

Error: Invalid Input - not valid: "xyz"

Select pages to rotate
...
```

Rotate then asks for direction and angle (`R90`, `L90`, `R180`, `L180`). After that you get the confirm screen, which lists exactly what will be processed:

```
========================================================
Rotate PDF
========================================================

Confirm Inputs: 2
   1. "Invoice.pdf"   Page 1 of 5
   2. "Report.pdf"   Pages 2-3 of 20

Rotation : R90

--------------------------------------------------------
[Enter] Continue
[0] Back
[00] Cancel & exit

Ans:
```

The keys mean the same on every screen:

| Key | Meaning |
|---|---|
| `[Enter]` | the default action (select all, continue, back to menu) |
| `[#]` | pick items from the list |
| `[0]` | back one screen — `[0] Back to menu` on a tool's first screen, `[0] Exit` on the main menu |
| `[00]` | cancel and close the program, from any tool screen |

`[00]` shows `Cancelled.` and waits for `[Enter] to close...`, so a double-clicked window never vanishes unread. Nothing asks "are you sure": nothing has been processed until the run starts, so backing out costs nothing. Because `0` and `00` are navigation keys, neither can be used as a password.

## Unreadable Files

Tools 4, 7, 8, 9 need page or sheet counts, so unreadable files are listed once before the file list:

```
========================================================
Rotate PDF
========================================================

Unreadable: 1
   1. "Broken.pdf"   :   PdfStreamError - Stream has ended unexpectedly

--------------------------------------------------------
[Enter] Skip & continue
[0] Back to menu
[00] Cancel & exit

Ans:
```

For the other tools an unreadable file shows up under `Failed` after processing.

## Reading the Result

```
========================================================
Rotate PDF
========================================================

Inputs: 2
   1. "Invoice.pdf"   Page 1 of 5
   2. "Report.pdf"   Pages 2-3 of 20

Rotation : R90

Processing
Rotating ...

Outputs: 2
   1. "Invoice_rotated.pdf"
   2. "Report_rotated.pdf"

Elapsed: 0.1s

--------------------------------------------------------
[Enter] Back to menu
[0] Exit

Ans:
```

When something failed, a `Skipped` and/or `Failed` list appears after `Outputs`, and `[0]` changes to `[0] Save error log & back to menu`:

```
Skipped: 1
   1. "Plain.pdf"   :   no password

Failed: 1
   1. "Broken.pdf"   :   ValueError - wrong password, could not decrypt

Elapsed: 0.4s

--------------------------------------------------------
[Enter] Back to menu
[0] Save error log & back to menu

Ans:
```

- **Skipped** — nothing to do for that file; not a problem.
- **Failed** — something went wrong; the reason follows the colon.

`[0] Save error log & back to menu` writes `error log_YYYYMMDD.txt` next to the exe with the step, error, context and traceback for every failure, then returns to the menu, which shows `Saved: "error log_YYYYMMDD.txt"` at the top. The log starts with the input and output folder. To close the program afterwards, press `[0] Exit` on the menu.

Split PDF and Selective Extract add one line above `Inputs`, because their output goes to a different folder: `Output folder: "...\Split"`. The other tools print no folder line — input and output are the same folder as the exe.

## Requirements

- Tools 1 and 7: Microsoft Excel and Microsoft Word must be installed.
- Tools 2, 3, 4, 5, 6, 8, 9: no additional software required.

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| Console closes immediately | Input files missing from folder — place files in the same folder as the script/exe |
| Tool 1 or 7 fails | Microsoft Office not installed, or an Office process is stuck — close any lingering EXCEL.EXE / WINWORD.EXE |
| `OfficePrompt - Word showed a prompt: '... it caused a serious error ...'` under "Failed" | Word needs a click to open that file (an earlier run was force-closed while opening it). Open it once in Word, click Yes, close it, then run the tool again |
| File listed under "Failed" | Read the reason after the colon; save the error log with `[0]` on the result screen |
| Wrong password on Unprotect | Check the filename for the `_pw_PASSWORD` hint |
| File listed under "Skipped" | Nothing to do for that file; the reason is shown after the colon |
| Antivirus blocks the `.exe` | Unsigned binary — whitelist it with your antivirus/IT if needed |
| Windows says "Windows protected your PC" | SmartScreen flag on unsigned executables — click "More info" then "Run anyway" |

---

## For Developers

### Single-file layout (`PDF Tools.py`, top to bottom)

| Block | Contents |
|---|---|
| Constants | `VERSION`, `LINE_WIDTH`, `MAX_LISTED`, extensions, `PW_MARK`, `OUT_FOLDERS` (Split, Extract), `OFFICE_TIMEOUT` |
| Signals | `UserCancelled`, `NoInput`, `ExitSession`, `ExitProgram`, `InvalidInput`, `SkipItem`, `GoBack` |
| Paths / naming | `get_script_dir`, `get_unique_path`, `out_dir`, `list_files`, `prefix_generator`, `dedupe` |
| Screens | `clear_screen`, `open_screen`, `options`, `prompt_screen`, `nav_keys`, `check_nav` — one screen layout and one key vocabulary (script-conventions Console rules 10–11) |
| File index + selection | `Entry`, `Token`, `Index`, `scan_pdf_index`, `parse_tokens`, `prompt_selection`, `select_step`, `run_inputs`, `confirm`, `prompt_password`, `unit_input_lines`, `ordered_groups` |
| Processing + report | `run_items`, `announce`, `begin`, `finish`, `tail`, `write_error_log`, `make_failure` |
| Office helpers | `OfficePrompt`, `OfficeApps` (watchdog), `excel_export`, `word_export`, `excel_visible_sheets`, `word_page_count`, `scan_office_index` |
| Tools | nine thin `tool_xxx()` functions |
| Menu | `TOOLS`, `build_menu`, `choose_tool`, `run_tool`, `main` |

Logic used by more than one tool lives in the shared blocks. A `tool_xxx()` holds only its own prompts and its worker.

### The same flow in every tool

```
scan files once -> run_inputs([select, (tool step)] + confirm)
  -> announce(folders + Inputs) -> begin("Gerund") -> run_items -> finish
```

- Selection level is `"file"` (tools 1, 2, 3, 5, 6) or `"unit"` (4, 7, 8, 9). One parser, one set of prompts.
- Every screen goes through `prompt_screen()`: clear, banner, choices so far, the `Error: Invalid Input ...` / `Saved: ...` notice, body, `-` line, keys, prompt. Invalid input raises `InvalidInput` in the screen's handler and the same screen is redrawn with the reason — no pause.
- Input screens are steps in `run_inputs()`: each step stores its answer in a shared dict, the confirm screen is always last. `[0]` is one screen back (on the first step it returns to the menu), `[00]` ends the session from any tool screen; there is deliberately no extra "are you sure" prompt.
- The scan happens **once**, before the steps; going back never re-scans or re-shows `Unreadable`.
- The confirm screen and the run report's `Inputs` block print the same lines (`describe()` builds them once), so what you confirmed is what is reported.
- Workers return output names and never print. `run_items` is the only loop; it files exceptions into `Failed` / `Skipped`.
- `begin()` is the statement directly before `run_items`, so the last line on screen always names the running stage.

### Cancel / exit architecture (exceptions, not return values)

| Exception | Raised when | Caught in |
|---|---|---|
| `InvalidInput` | a screen's handler rejects the answer | `prompt_screen()` → same screen redrawn with the reason |
| `GoBack` | `[0]` on an input screen after the first | `run_inputs()` → previous screen |
| `UserCancelled` | `[0]` on a tool's first screen / `Unreadable` | `run_tool()` → back to menu |
| `NoInput` | no files found for the tool | `run_tool()` → normal tail |
| `ExitSession` | `[00]` on any tool screen | `main()` → `Cancelled.` + `[Enter] to close...` |
| `ExitProgram` | `[0] Exit` on the Tool finished screen (no failures) | `main()` → closes immediately |
| `SkipItem` | worker: nothing to do for the item | `run_items()` |

New mid-tool exit point: `raise UserCancelled()`, never `return`. An unexpected exception inside a tool is caught in `run_tool()`, shown as `FAILED`, and goes to the normal tail so the log can be saved.

### Tool finished screen

`[Enter] Back to menu` always. `[0]` is `Save error log & back to menu` when something failed, otherwise `Exit` — a lone `[Enter]` would give no choice, and closing would take `[Enter]` then `[0]` on the menu.

### Error log

Written only when `[0] Save error log & back to menu` is pressed at the tail. The header carries the resolved `Input folder:` and `Output folder:`. `make_failure()` calls `traceback.format_exc()` **inside** the `except` block — outside it the traceback is gone. Fields per entry: Step, Error, Context, Traceback. Filename `error log_YYYYMMDD.txt`, same-day re-runs get `_1`, `_2`.

### Output naming

Outputs go next to the inputs (`out_dir()`); only Split and Selective Extract get a subfolder. Folder lines print on the run report only when the output folder differs from the script folder. All outputs are deliverables: no run date. A processed copy keeps the source name plus a lowercase process suffix (`_merged`, `_rotated`); Protect uses `_pw_PASSWORD` so the password hint travels with the file, and Unprotect strips it. Split and Extract add a page label with a space (`name p3.pdf`); Convert and Unprotect only change the format or state, so they keep the plain name. Combine is not named after a source (`Combined 12 pages.pdf`). The collision counter `_1`, `_2` follows everything.

### Office watchdog (tools 1 and 7)

Every Excel/Word job goes through `OfficeApps.run()`. The instance is `DispatchEx` + hidden, and its PID is captured at start. A watchdog thread checks once a second: a visible window owned by that PID means Office is waiting on a prompt (e.g. Word's "it caused a serious error last time", which `DisplayAlerts` does not suppress); no answer after `OFFICE_TIMEOUT` seconds means it hung. Either way the watchdog reads the prompt text, kills **only that PID**, and the item lands under `Failed` as `OfficePrompt`; the next item gets a fresh instance. The user's own Excel/Word windows are never touched.

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

Format: `v{major}.{minor}_{YYYYMMDD}`. Current: `v2.2_20261113`. Update the `VERSION` constant in `PDF Tools.py`, `version_info.txt`, `PDF Tools Instruction.txt` and this file.

### Packaging

Build from a clean virtual environment so the exe only bundles what the script needs:

```bash
python -m venv venv
venv\Scripts\pip install pyinstaller pypdf pywin32
venv\Scripts\python -m PyInstaller --onefile --console --version-file=version_info.txt ^
  --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.server ^
  --hidden-import pywintypes --hidden-import win32gui --hidden-import win32process ^
  "PDF Tools.py"
```

`version_info.txt` embeds author/version info into the exe's Windows Properties > Details tab — this is what survives if the `.exe` ever gets forwarded on its own, with no source code or docs attached.

Output: `dist/PDF Tools.exe`. The `build/` folder and `.spec` file can be deleted after packaging.

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
- Screen clearing uses the Windows `cls` command: it works when double-clicked or run in cmd / PowerShell, not in an IDE output panel.

---
**Version:** v2.2_20261113

**Author:** Yeoh Huey Sun
**Email:** greyw901@gmail.com
**GitHub:** https://github.com/greyw901
