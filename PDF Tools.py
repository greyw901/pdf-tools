r"""
------------
Master launcher for all PDF tools.
Combines: Convert to PDF, Merge PDF, Split PDF, Rotate PDF,
          Protect PDF, Unprotect PDF, Convert to PDF_Precise,
          Combine PDF Pages, Extract PDF Pages

Author: Yeoh Huey Sun

Output folders (auto-created in same folder as script/exe):
  To PDF\     - Convert to PDF / Convert to PDF_Precise
  Merged\     - Merge PDF
  Split\      - Split PDF
  Extract\    - Extract PDF Pages
  Combined\   - Combine PDF Pages
  Protected\  - Protect PDF
  Unprotected\ - Unprotect PDF
  Rotated\    - Rotate PDF
"""

import os
import sys
import re
import logging
import datetime
from itertools import product

logging.getLogger("pypdf").setLevel(logging.ERROR)

# ── Path resolution (works for .py and PyInstaller .exe) ──
def get_script_dir():
    return os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, 'frozen', False) else __file__
    ))

# ── Shared utilities ───────────────────────────────────────
def get_unique_path(folder, name):
    candidate = os.path.join(folder, name)
    if not os.path.exists(candidate):
        return candidate
    base, ext = os.path.splitext(name)
    counter = 1
    while True:
        candidate = os.path.join(folder, f"{base}_{counter}{ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1

def get_pdf_files(folder):
    return sorted([
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(folder, f))
    ])

def prefix_generator():
    letters = "abcdefghijklmnopqrstuvwxyz"
    length = 1
    while True:
        for combo in product(letters, repeat=length):
            yield "".join(combo)
        length += 1

# ══════════════════════════════════════════════════════════
#  TOOL 1 — Convert to PDF
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf(script_dir):
    import win32com.client

    out_folder = os.path.join(script_dir, "To PDF")
    os.makedirs(out_folder, exist_ok=True)

    excel_exts = {".xls", ".xlsx", ".xlsm", ".xlsb"}
    word_exts  = {".doc", ".docx"}
    excel_files, word_files = [], []

    for fname in sorted(os.listdir(script_dir)):
        fpath = os.path.join(script_dir, fname)
        if not os.path.isfile(fpath):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext in excel_exts:
            excel_files.append(fpath)
        elif ext in word_exts:
            word_files.append(fpath)

    if not excel_files and not word_files:
        print("No Excel or Word files found.")
        return

    errors = []

    if excel_files:
        print("Starting Excel...\n")
        xl_app = win32com.client.Dispatch("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        try:
            for fpath in excel_files:
                fname     = os.path.basename(fpath)
                base_name = os.path.splitext(fname)[0]
                out_path  = get_unique_path(out_folder, base_name + ".pdf")
                print(f"  {fname}  ->  {os.path.basename(out_path)}")
                try:
                    wb = xl_app.Workbooks.Open(fpath, ReadOnly=True, UpdateLinks=False)
                    visible = [s.Name for s in wb.Sheets if s.Visible == -1]
                    if not visible:
                        print(f"ERROR: No visible sheets found. ({fname})")
                        wb.Close(SaveChanges=False)
                        continue
                    wb.ExportAsFixedFormat(Type=0, Filename=out_path, Quality=0,
                                           IncludeDocProperties=True,
                                           IgnorePrintAreas=False, OpenAfterPublish=False)
                    wb.Close(SaveChanges=False)
                except Exception as e:
                    print(f"ERROR: pywintypes.com_error - Could be: file is open in another program / corrupted / Excel frozen. ({fname})")
                    print(f"Error Details: {e}")
                    errors.append(fname)
        finally:
            xl_app.Quit()

    if word_files:
        print("\nStarting Word...\n")
        wd_app = win32com.client.Dispatch("Word.Application")
        wd_app.Visible = False
        wd_app.DisplayAlerts = False
        try:
            for fpath in word_files:
                fname     = os.path.basename(fpath)
                base_name = os.path.splitext(fname)[0]
                out_path  = get_unique_path(out_folder, base_name + ".pdf")
                print(f"  {fname}  ->  {os.path.basename(out_path)}")
                try:
                    doc = wd_app.Documents.Open(fpath, ReadOnly=True, AddToRecentFiles=False)
                    doc.ExportAsFixedFormat(OutputFileName=out_path, ExportFormat=17,
                                            OpenAfterExport=False, OptimizeFor=0, Range=0,
                                            IncludeDocProps=True, CreateBookmarks=0,
                                            DocStructureTags=True)
                    doc.Close(SaveChanges=False)
                except Exception as e:
                    print(f"ERROR: pywintypes.com_error - Could be: file is open in another program / corrupted / Word frozen. ({fname})")
                    print(f"Error Details: {e}")
                    errors.append(fname)
        finally:
            wd_app.Quit()

    if errors:
        print(f"\nFailed:")
        for f in errors:
            print(f'  "{f}"')

    print(f"\nOutput: To PDF\\")
    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 2 — Merge PDF
# ══════════════════════════════════════════════════════════
def tool_merge_pdf(script_dir):
    from pypdf import PdfWriter, PdfReader

    out_folder = os.path.join(script_dir, "Merged")
    os.makedirs(out_folder, exist_ok=True)

    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    first_stem = os.path.splitext(pdf_files[0])[0]
    out_name   = first_stem + "_merged.pdf"
    out_path   = get_unique_path(out_folder, out_name)

    print(f"Merging {len(pdf_files)} file(s):")
    for f in pdf_files:
        print(f"  {f}")

    writer = PdfWriter()
    errors = []
    for fname in pdf_files:
        fpath = os.path.join(script_dir, fname)
        try:
            reader = PdfReader(fpath)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            errors.append(fname)

    if len(errors) == len(pdf_files):
        print("\nAll files failed — output not written.")
        return

    with open(out_path, "wb") as f:
        writer.write(f)

    print(f"\nOutput: Merged\\")
    print(out_name)

    if errors:
        print(f"\nFailed:")
        for f in errors:
            print(f'  "{f}"')

    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 3 — Split PDF
# ══════════════════════════════════════════════════════════
def tool_split_pdf(script_dir):
    from pypdf import PdfWriter, PdfReader

    out_folder = os.path.join(script_dir, "Split")
    os.makedirs(out_folder, exist_ok=True)

    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    print(f"Splitting {len(pdf_files)} file(s):\n")
    errors = []
    for fname in pdf_files:
        fpath = os.path.join(script_dir, fname)
        stem  = os.path.splitext(fname)[0]
        try:
            reader = PdfReader(fpath)
            print(f"  (Pages: {len(reader.pages)}) {fname}")
            for i, page in enumerate(reader.pages, start=1):
                writer = PdfWriter()
                writer.add_page(page)
                out_path = get_unique_path(out_folder, f"{stem}_{i}.pdf")
                with open(out_path, "wb") as f:
                    writer.write(f)
        except Exception as e:
            errors.append(fname)

    if errors:
        print(f"\nFailed:")
        for f in errors:
            print(f'  "{f}"')

    print(f"\nOutput: Split\\")
    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 4 — Protect PDF
# ══════════════════════════════════════════════════════════
def tool_protect_pdf(script_dir):
    from pypdf import PdfReader, PdfWriter

    out_folder = os.path.join(script_dir, "Protected")
    os.makedirs(out_folder, exist_ok=True)

    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    print(f"Found {len(pdf_files)} PDF file(s):\n")
    for f in pdf_files:
        print(f"  {f}")

    print()
    password = input("Enter password to apply: ")
    if not password:
        print("No password entered. Aborted.")
        return

    print()
    errors  = []
    success = []
    for fname in pdf_files:
        fpath = os.path.join(script_dir, fname)
        stem  = os.path.splitext(fname)[0]
        if "_pw_" in stem:
            print(f"Skipped (already protected): {fname}")
            continue
        out_name = f"{stem}_pw_{password}.pdf"
        out_path = get_unique_path(out_folder, out_name)
        try:
            reader = PdfReader(fpath)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(password)
            with open(out_path, "wb") as f_out:
                writer.write(f_out)
            print(f"  {fname}")
            success.append(os.path.basename(out_path))
        except Exception as e:
            errors.append((fname, str(e)))

    if success:
        print(f"\nOutput: Protected\\")
        for n in success:
            print(n)

    if errors:
        print(f"\nFailed:")
        for fname, err in errors:
            print(f'  "{fname}"')
        choice = input("\nPress 1 to save error log, or Enter to continue: ").strip()
        if choice == "1":
            log_path = os.path.join(script_dir, "protect_pdf_errors.txt")
            with open(log_path, "w") as log:
                log.write("Protect PDF — Error Log\n" + "=" * 40 + "\n\n")
                for fname, err in errors:
                    log.write(f"{fname}\n  {err}\n\n")
            print(f"Log saved: protect_pdf_errors.txt")

    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 5 — Unprotect PDF
# ══════════════════════════════════════════════════════════
def tool_unprotect_pdf(script_dir):
    from pypdf import PdfReader, PdfWriter

    out_folder = os.path.join(script_dir, "Unprotected")
    os.makedirs(out_folder, exist_ok=True)

    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    print(f"Found {len(pdf_files)} PDF file(s):\n")
    for f in pdf_files:
        print(f"  {f}")

    print()
    password = input("Enter password to unlock: ")
    if not password:
        print("No password entered. Aborted.")
        return

    print()
    errors      = []
    success     = []
    already_ok  = []
    no_password = []

    for fname in pdf_files:
        fpath = os.path.join(script_dir, fname)
        stem  = os.path.splitext(fname)[0]

        if stem.endswith("_unprotected"):
            already_ok.append(fname)
            continue

        try:
            reader = PdfReader(fpath)

            if not reader.is_encrypted:
                no_password.append(fname)
                continue

            result = reader.decrypt(password)
            if result == 0:
                raise ValueError("Wrong password — could not decrypt.")

            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            out_name = f"{stem}_unprotected.pdf"
            out_path = get_unique_path(out_folder, out_name)
            with open(out_path, "wb") as f_out:
                writer.write(f_out)
            print(f"  {fname}")
            success.append(os.path.basename(out_path))

        except ValueError as e:
            errors.append((fname, str(e)))
        except Exception as e:
            errors.append((fname, str(e)))

    if already_ok:
        print(f"\nSkipped (already unprotected):")
        for f in already_ok:
            print(f'  "{f}"')

    if no_password:
        print(f"\nSkipped (no password):")
        for f in no_password:
            print(f'  "{f}"')

    if success:
        print(f"\nOutput: Unprotected\\")
        for n in success:
            print(n)

    if errors:
        print(f"\nFailed:")
        for fname, err in errors:
            print(f'  "{fname}"')
        choice = input("\nPress 1 to save error log, or Enter to continue: ").strip()
        if choice == "1":
            log_path = os.path.join(script_dir, "unprotect_pdf_errors.txt")
            with open(log_path, "w") as log:
                log.write("Unprotect PDF — Error Log\n" + "=" * 40 + "\n\n")
                for fname, err in errors:
                    log.write(f"{fname}\n  {err}\n\n")
            print(f"Log saved: unprotect_pdf_errors.txt")

    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 6 — Convert to PDF_Precise
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf_precise(script_dir):
    import win32com.client

    out_folder = os.path.join(script_dir, "To PDF")
    os.makedirs(out_folder, exist_ok=True)

    def parse_selection(raw, max_index):
        raw = raw.strip()
        if not raw:
            return None
        selected = []
        for token in raw.replace(",", " ").split():
            if "-" in token:
                parts = token.split("-")
                try:
                    start, end = int(parts[0]), int(parts[1])
                    for n in range(start, end + 1):
                        if 1 <= n <= max_index:
                            selected.append(n - 1)
                except ValueError:
                    pass
            else:
                try:
                    n = int(token)
                    if 1 <= n <= max_index:
                        selected.append(n - 1)
                except ValueError:
                    pass
        seen, result = set(), []
        for i in selected:
            if i not in seen:
                seen.add(i)
                result.append(i)
        return result if result else None

    excel_exts = {".xls", ".xlsx", ".xlsm", ".xlsb"}
    word_exts  = {".doc", ".docx"}
    excel_files, word_files = [], []

    for fname in sorted(os.listdir(script_dir)):
        fpath = os.path.join(script_dir, fname)
        if not os.path.isfile(fpath):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext in excel_exts:
            excel_files.append(fpath)
        elif ext in word_exts:
            word_files.append(fpath)

    if not excel_files and not word_files:
        print("No Excel or Word files found.")
        return

    errors = []

    if excel_files:
        print("Starting Excel...\n")
        xl_app = win32com.client.Dispatch("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        try:
            for fpath in excel_files:
                fname     = os.path.basename(fpath)
                base_name = os.path.splitext(fname)[0]
                out_path  = get_unique_path(out_folder, base_name + ".pdf")
                print(f"  {fname}  ->  {os.path.basename(out_path)}")
                try:
                    wb = xl_app.Workbooks.Open(fpath, ReadOnly=True, UpdateLinks=False)
                    visible = [s.Name for s in wb.Sheets if s.Visible == -1]
                    if not visible:
                        print(f"ERROR: No visible sheets found. ({fname})")
                        wb.Close(SaveChanges=False)
                        continue
                    print(f"  Sheets:")
                    for i, name in enumerate(visible, start=1):
                        print(f"    {i} - {name}")
                    print("  [Selections] - e.g. 1, 2, 4-6")
                    print("  [Enter] - Select all")
                    print("  [0] - Cancel & Return to Menu")
                    raw = input("  ")
                    if raw.strip() == "0":
                        wb.Close(SaveChanges=False)
                        return "CANCELLED"
                    indices = parse_selection(raw, len(visible))
                    selected_names = visible if indices is None else [visible[i] for i in indices]
                    if not selected_names:
                        print(f"Invalid Input: \"{raw}\"")
                        wb.Close(SaveChanges=False)
                        continue
                    hidden = []
                    for sheet in wb.Sheets:
                        if sheet.Visible == -1 and sheet.Name not in selected_names:
                            sheet.Visible = False
                            hidden.append(sheet.Name)
                    try:
                        wb.ExportAsFixedFormat(Type=0, Filename=out_path, Quality=0,
                                               IncludeDocProperties=True,
                                               IgnorePrintAreas=False, OpenAfterPublish=False)
                    finally:
                        for sheet in wb.Sheets:
                            if sheet.Name in hidden:
                                sheet.Visible = -1
                    wb.Close(SaveChanges=False)
                    print()
                except Exception as e:
                    print(f"ERROR: pywintypes.com_error - Could be: file is open in another program / corrupted / Excel frozen. ({fname})")
                    print(f"Error Details: {e}\n")
                    errors.append(fname)
        finally:
            xl_app.Quit()

    if word_files:
        print("\nStarting Word...\n")
        wd_app = win32com.client.Dispatch("Word.Application")
        wd_app.Visible = False
        wd_app.DisplayAlerts = False
        try:
            for fpath in word_files:
                fname     = os.path.basename(fpath)
                base_name = os.path.splitext(fname)[0]
                out_path  = get_unique_path(out_folder, base_name + ".pdf")
                print(f"  {fname}  ->  {os.path.basename(out_path)}")
                try:
                    doc = wd_app.Documents.Open(fpath, ReadOnly=True, AddToRecentFiles=False)
                    doc.ExportAsFixedFormat(OutputFileName=out_path, ExportFormat=17,
                                            OpenAfterExport=False, OptimizeFor=0, Range=0,
                                            IncludeDocProps=True, CreateBookmarks=0,
                                            DocStructureTags=True)
                    doc.Close(SaveChanges=False)
                except Exception as e:
                    print(f"ERROR: pywintypes.com_error - Could be: file is open in another program / corrupted / Word frozen. ({fname})")
                    print(f"Error Details: {e}")
                    errors.append(fname)
        finally:
            wd_app.Quit()

    if errors:
        print(f"Failed:")
        for f in errors:
            print(f'  "{f}"')

    print(f"\nOutput: To PDF\\")
    print("\nDone.")

# ══════════════════════════════════════════════════════════
#  TOOL 7 — Combine PDF Pages
# ══════════════════════════════════════════════════════════
def tool_combine_pdf_pages(script_dir):
    from pypdf import PdfReader, PdfWriter

    out_folder = os.path.join(script_dir, "Combined")
    os.makedirs(out_folder, exist_ok=True)

    def parse_selection(raw, code_map, prefix_page_map):
        tokens = re.split(r'[\s,]+', raw.strip().lower())
        selected = []
        invalid  = []
        for token in tokens:
            if not token:
                continue
            if re.match(r'^[a-z]+$', token):
                if token in prefix_page_map:
                    for code in prefix_page_map[token]:
                        selected.append(code)
                else:
                    invalid.append(token)
                continue
            range_match = re.match(r'^([a-z]+)(\d+)-([a-z]+)?(\d+)$', token)
            if range_match:
                p1 = range_match.group(1)
                n1 = int(range_match.group(2))
                p2 = range_match.group(3) or p1
                n2 = int(range_match.group(4))
                if p1 == p2:
                    range_valid = True
                    for n in range(n1, n2 + 1):
                        code = f"{p1}{n}"
                        if code in code_map:
                            selected.append(code)
                        else:
                            range_valid = False
                    if not range_valid:
                        invalid.append(token)
                else:
                    invalid.append(token)
                continue
            if token in code_map:
                selected.append(token)
            else:
                invalid.append(token)
        return selected, invalid

    # ── Scan files ONCE before entering the selection loop ──
    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    code_map, prefix_page_map, display_lines = {}, {}, []
    unreadable = []
    gen = prefix_generator()

    for fpath_name in pdf_files:
        fpath  = os.path.join(script_dir, fpath_name)
        prefix = next(gen)
        try:
            reader    = PdfReader(fpath)
            num_pages = len(reader.pages)
        except Exception as e:
            unreadable.append(fpath_name)
            continue
        page_codes = []
        for i in range(num_pages):
            code = f"{prefix}{i+1}"
            code_map[code] = (fpath, i)
            page_codes.append(code)
        prefix_page_map[prefix] = page_codes
        display_lines.append(f"  {prefix} (Pages: {prefix}1-{prefix}{num_pages}) - {fpath_name}")

    if unreadable:
        quoted = "\n".join(f'"{f}"' for f in unreadable)
        print(f"Unreadable PDF:\n{quoted}\n")
        print("[Enter] - Skip & Continue")
        print("[0] - Cancel & Return to Menu")
        choice = input().strip()
        if choice == "0":
            return "CANCELLED"
        print()

    if not code_map:
        print("No readable PDF files found.")
        return

    while True:
        print("\n" + "\n".join(display_lines) + "\n")
        print("[Selections] - e.g. a1, b, c4-6")
        print("[Enter] - Select all")
        print("[0] - Cancel & Return to Menu")
        raw = input().strip().lower()

        if raw == "0":
            return "CANCELLED"

        if not raw:
            all_prefixes = list(prefix_page_map.keys())
            raw = " ".join(all_prefixes)

        selected_codes, invalid_tokens = parse_selection(raw, code_map, prefix_page_map)

        if invalid_tokens:
            quoted = ", ".join(f'"{t}"' for t in invalid_tokens)
            print(f"\nInvalid Input: {quoted}")
            print("[Enter] - Skip Invalid Input & Continue")
            print("[N] - Re-enter")
            print("[0] - Cancel & Return to Menu")
            choice = input().strip().lower()
            if choice == "0":
                return "CANCELLED"
            if choice == "n":
                print()
                continue

        if not selected_codes:
            print("No valid pages selected. Aborted.")
            return

        print(f"\nSelection summary (Pages: {len(selected_codes)}):")
        prev_prefix = None
        for code in selected_codes:
            fpath, page_idx = code_map[code]
            prefix = re.match(r'^([a-z]+)', code).group(1)
            if prefix != prev_prefix:
                print(f"  {os.path.basename(fpath)}")
                prev_prefix = prefix
            print(f"    {code} - page {page_idx + 1}")

        print()
        print("[Enter] - Proceed")
        print("[N] - Re-enter")
        print("[0] - Cancel & Return to Menu")
        proceed = input().strip().lower()
        if proceed == "0":
            return "CANCELLED"
        if proceed == "n":
            print()
            continue

        readers, writer, errors = {}, PdfWriter(), []
        for code in selected_codes:
            fpath, page_idx = code_map[code]
            try:
                if fpath not in readers:
                    readers[fpath] = PdfReader(fpath)
                writer.add_page(readers[fpath].pages[page_idx])
            except Exception as e:
                errors.append(code)

        total_pages = len(selected_codes) - len(errors)
        date_str    = datetime.datetime.now().strftime("%Y%m%d")
        out_name    = f"{date_str} Combined_{total_pages} pages.pdf"
        out_path    = get_unique_path(out_folder, out_name)

        with open(out_path, "wb") as f:
            writer.write(f)

        print(f"\nOutput: Combined\\")
        print(f"{os.path.basename(out_path)}")
        if errors:
            print(f"\nFailed pages: {', '.join(errors)}")

        print("\nDone.")
        return


def tool_extract_pdf_pages(script_dir):
    from pypdf import PdfReader, PdfWriter

    out_folder = os.path.join(script_dir, "Extract")
    os.makedirs(out_folder, exist_ok=True)

    def parse_tokens(raw, code_map, prefix_page_map):
        tokens = re.split(r',', raw.strip().lower())
        result  = []
        invalid = []
        for token in tokens:
            token = token.strip()
            if not token:
                continue
            if re.match(r'^[a-z]+$', token):
                if token in prefix_page_map:
                    for code in prefix_page_map[token]:
                        n = int(re.search(r'\d+', code).group())
                        result.append((f"p{n}", [code]))
                else:
                    invalid.append(token)
                continue
            range_match = re.match(r'^([a-z]+)(\d+)-([a-z]+)?(\d+)$', token)
            if range_match:
                p1 = range_match.group(1)
                n1 = int(range_match.group(2))
                p2 = range_match.group(3) or p1
                n2 = int(range_match.group(4))
                if p1 == p2:
                    codes = []
                    range_valid = True
                    for n in range(n1, n2 + 1):
                        code = f"{p1}{n}"
                        if code in code_map:
                            codes.append(code)
                        else:
                            range_valid = False
                    if range_valid and codes:
                        result.append((f"p{n1}-{n2}", codes))
                    else:
                        invalid.append(token)
                else:
                    invalid.append(token)
                continue
            if token in code_map:
                n = int(re.search(r'\d+', token).group())
                result.append((f"p{n}", [token]))
            else:
                invalid.append(token)
        return result, invalid

    def write_pdf(pages_data, out_path):
        writer = PdfWriter()
        for reader, idx in pages_data:
            writer.add_page(reader.pages[idx])
        with open(out_path, "wb") as f:
            writer.write(f)

    # ── Scan files ONCE before entering the selection loop ──
    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    code_map        = {}
    prefix_page_map = {}
    prefix_file_map = {}
    display_lines   = []
    unreadable      = []
    gen             = prefix_generator()

    for fpath_name in pdf_files:
        fpath  = os.path.join(script_dir, fpath_name)
        prefix = next(gen)
        try:
            reader    = PdfReader(fpath)
            num_pages = len(reader.pages)
        except Exception as e:
            unreadable.append(fpath_name)
            continue
        page_codes = []
        for i in range(num_pages):
            code = f"{prefix}{i+1}"
            code_map[code] = (fpath, i)
            page_codes.append(code)
        prefix_page_map[prefix] = page_codes
        prefix_file_map[prefix] = fpath
        display_lines.append(f"  {prefix} (Pages: {prefix}1-{prefix}{num_pages}) - {fpath_name}")

    if unreadable:
        quoted = "\n".join(f'"{f}"' for f in unreadable)
        print(f"Unreadable PDF:\n{quoted}\n")
        print("[Enter] - Skip & Continue")
        print("[0] - Cancel & Return to Menu")
        choice = input().strip()
        if choice == "0":
            return "CANCELLED"
        print()

    if not code_map:
        print("No readable PDF files found.")
        return

    while True:
        print("\n" + "\n".join(display_lines) + "\n")
        print("[Selections] - e.g. a1, b2, c4-6")
        print("[Enter] - Split all")
        print("[0] - Cancel & Return to Menu")
        raw = input().strip().lower()

        if raw == "0":
            return "CANCELLED"

        if not raw:
            total = sum(len(v) for v in prefix_page_map.values())
            print(f"\nSplit all: Output files: {total}")
            print("[Enter] - Proceed")
            print("[N] - Re-enter")
            print("[0] - Cancel & Return to Menu")
            proceed = input().strip().lower()
            if proceed == "0":
                return "CANCELLED"
            if proceed == "n":
                print()
                continue

            print("\nSplitting...")
            readers, errors = {}, []
            for prefix, codes in prefix_page_map.items():
                fpath = prefix_file_map[prefix]
                stem  = os.path.splitext(os.path.basename(fpath))[0]
                if fpath not in readers:
                    try:
                        readers[fpath] = PdfReader(fpath)
                    except Exception as e:
                        errors.append(os.path.basename(fpath))
                        continue
                reader = readers[fpath]
                for code in codes:
                    _, page_idx = code_map[code]
                    out_name = f"{stem}_p{page_idx+1}.pdf"
                    out_path = get_unique_path(out_folder, out_name)
                    try:
                        write_pdf([(reader, page_idx)], out_path)
                    except Exception as e:
                        errors.append(out_name)
            if errors:
                print(f"\nFailed:")
                for f in errors:
                    print(f'  "{f}"')
            print(f"\nOutput: Extract\\")
            print("\nDone.")
            return

        token_list, invalid_tokens = parse_tokens(raw, code_map, prefix_page_map)

        if invalid_tokens:
            quoted = ", ".join(f'"{t}"' for t in invalid_tokens)
            print(f"\nInvalid Input: {quoted}")
            print("[Enter] - Skip Invalid Input & Continue")
            print("[N] - Re-enter")
            print("[0] - Cancel & Return to Menu")
            choice = input().strip().lower()
            if choice == "0":
                return "CANCELLED"
            if choice == "n":
                print()
                continue

        if not token_list:
            print("No valid selection. Aborted.")
            return

        print(f"\nExtract summary (Output files: {len(token_list)}):")
        for label, codes in token_list:
            fpath, _ = code_map[codes[0]]
            stem     = os.path.splitext(os.path.basename(fpath))[0]
            pages    = [code_map[c][1] + 1 for c in codes]
            print(f"  {stem}_{label}.pdf  (Pages: {', '.join(str(p) for p in pages)})")

        print()
        print("[Enter] - Proceed")
        print("[N] - Re-enter")
        print("[0] - Cancel & Return to Menu")
        proceed = input().strip().lower()
        if proceed == "0":
            return "CANCELLED"
        if proceed == "n":
            print()
            continue

        print("\nExtracting...")
        readers, errors = {}, []
        for label, codes in token_list:
            fpath, _ = code_map[codes[0]]
            stem     = os.path.splitext(os.path.basename(fpath))[0]
            out_name = f"{stem}_{label}.pdf"
            out_path = get_unique_path(out_folder, out_name)
            try:
                if fpath not in readers:
                    readers[fpath] = PdfReader(fpath)
                pages_data = [(readers[fpath], code_map[c][1]) for c in codes]
                write_pdf(pages_data, out_path)
            except Exception as e:
                errors.append(out_name)

        if errors:
            print(f"\nFailed:")
            for f in errors:
                print(f'  "{f}"')

        print(f"\nOutput: Extract\\")
        print("\nDone.")
        return


def tool_rotate_pdf(script_dir):
    from pypdf import PdfReader, PdfWriter

    out_folder = os.path.join(script_dir, "Rotated")
    os.makedirs(out_folder, exist_ok=True)

    def parse_selection_rotate(raw, code_map, prefix_page_map):
        tokens = re.split(r'[\s,]+', raw.strip().lower())
        selected = []
        invalid  = []
        for token in tokens:
            if not token:
                continue
            if re.match(r'^[a-z]+$', token):
                if token in prefix_page_map:
                    for code in prefix_page_map[token]:
                        selected.append(code)
                else:
                    invalid.append(token)
                continue
            range_match = re.match(r'^([a-z]+)(\d+)-([a-z]+)?(\d+)$', token)
            if range_match:
                p1 = range_match.group(1)
                n1 = int(range_match.group(2))
                p2 = range_match.group(3) or p1
                n2 = int(range_match.group(4))
                if p1 == p2:
                    range_valid = True
                    for n in range(n1, n2 + 1):
                        code = f"{p1}{n}"
                        if code in code_map:
                            selected.append(code)
                        else:
                            range_valid = False
                    if not range_valid:
                        invalid.append(token)
                else:
                    invalid.append(token)
                continue
            if token in code_map:
                selected.append(token)
            else:
                invalid.append(token)
        return selected, invalid

    # ── Scan files ONCE before entering the selection loop ──
    pdf_files = get_pdf_files(script_dir)
    if not pdf_files:
        print("No PDF files found.")
        return

    code_map        = {}
    prefix_page_map = {}
    prefix_file_map = {}
    display_lines   = []
    unreadable      = []
    gen             = prefix_generator()

    for fname in pdf_files:
        fpath  = os.path.join(script_dir, fname)
        prefix = next(gen)
        try:
            reader    = PdfReader(fpath)
            num_pages = len(reader.pages)
        except Exception as e:
            unreadable.append(fname)
            continue
        page_codes = []
        for i in range(num_pages):
            code = f"{prefix}{i+1}"
            code_map[code] = (fpath, i)
            page_codes.append(code)
        prefix_page_map[prefix] = page_codes
        prefix_file_map[prefix] = fpath
        display_lines.append(f"  {prefix} (Pages: {prefix}1-{prefix}{num_pages}) - {fname}")

    if unreadable:
        quoted = "\n".join(f'"{f}"' for f in unreadable)
        print(f"Unreadable PDF:\n{quoted}\n")
        print("[Enter] - Skip & Continue")
        print("[0] - Cancel & Return to Menu")
        choice = input().strip()
        if choice == "0":
            return "CANCELLED"
        print()

    if not code_map:
        print("No readable PDF files found.")
        return

    while True:
        print("\n" + "\n".join(display_lines) + "\n")

        print("[Selections] - e.g. a1, b2, c4-6")
        print("[Enter] - Select all")
        print("[0] - Cancel & Return to Menu")
        raw = input().strip().lower()

        if raw == "0":
            return "CANCELLED"

        if not raw:
            selected_codes = []
            for prefix in prefix_page_map:
                selected_codes.extend(prefix_page_map[prefix])
            invalid_tokens = []
        else:
            selected_codes, invalid_tokens = parse_selection_rotate(raw, code_map, prefix_page_map)

        if invalid_tokens:
            quoted = ", ".join(f'"{t}"' for t in invalid_tokens)
            print(f"\nInvalid Input: {quoted}")
            print("[Enter] - Skip Invalid Input & Continue")
            print("[N] - Re-enter")
            print("[0] - Cancel & Return to Menu")
            choice = input().strip().lower()
            if choice == "0":
                return "CANCELLED"
            if choice == "n":
                print()
                continue

        if not selected_codes:
            print("No valid pages selected. Aborted.")
            return

        print("Rotation? 90 / 180 / 270: ", end="")
        angle_raw = input().strip()
        if angle_raw not in ("90", "180", "270"):
            print(f"Invalid Input: \"{angle_raw}\"")
            return
        angle = int(angle_raw)

        print(f"\nRotate summary (Pages: {len(selected_codes)}; Rotation: {angle}):")
        prev_prefix = None
        for code in selected_codes:
            fpath, page_idx = code_map[code]
            prefix = re.match(r'^([a-z]+)', code).group(1)
            if prefix != prev_prefix:
                print(f"  {os.path.basename(fpath)}")
                prev_prefix = prefix
            print(f"    {code} - page {page_idx + 1}")

        print()
        print("[Enter] - Proceed")
        print("[N] - Re-enter")
        print("[0] - Cancel & Return to Menu")
        proceed = input().strip().lower()
        if proceed == "0":
            return "CANCELLED"
        if proceed == "n":
            print()
            continue

        pages_to_rotate_by_file = {}
        for code in selected_codes:
            fpath, page_idx = code_map[code]
            pages_to_rotate_by_file.setdefault(fpath, set()).add(page_idx)

        print("\nRotating...")
        errors = []
        rotated_names = []
        for fpath, rotate_indices in pages_to_rotate_by_file.items():
            fname = os.path.basename(fpath)
            stem  = os.path.splitext(fname)[0]
            try:
                reader = PdfReader(fpath)
                writer = PdfWriter()
                for i, page in enumerate(reader.pages):
                    if i in rotate_indices:
                        page.rotate(angle)
                    writer.add_page(page)
                out_name = f"{stem}_rotated.pdf"
                out_path = get_unique_path(out_folder, out_name)
                with open(out_path, "wb") as f:
                    writer.write(f)
                rotated_names.append(os.path.basename(out_path))
            except Exception as e:
                errors.append(fname)

        if rotated_names:
            print(f"\nOutput: Rotated\\")
            for n in rotated_names:
                print(n)
        if errors:
            print(f"\nFailed:")
            for f in errors:
                print(f'  "{f}"')

        print("\nDone.")
        return

# ══════════════════════════════════════════════════════════
#  MENU
# ══════════════════════════════════════════════════════════
MENU = """
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
"""

TOOLS = {
    "1": ("Convert to PDF",
          "Converts all Excel and Word files in this folder to PDF (all visible sheets).",
          tool_convert_to_pdf),
    "2": ("Merge PDF",
          "Combines all PDFs in this folder into one, in alphabetical order.",
          tool_merge_pdf),
    "3": ("Split PDF",
          "Splits every PDF in this folder into individual single-page files.",
          tool_split_pdf),
    "4": ("Rotate PDF",
          "Rotates selected pages in PDFs by 90/180/270 degrees.",
          tool_rotate_pdf),
    "5": ("Protect PDF",
          "Adds password protection to all PDFs in this folder.",
          tool_protect_pdf),
    "6": ("Unprotect PDF",
          "Removes password protection from all PDFs in this folder.",
          tool_unprotect_pdf),
    "7": ("Convert to PDF_Precise",
          "Converts Excel and Word files to PDF, letting you choose which sheets to include.",
          tool_convert_to_pdf_precise),
    "8": ("Combine PDF Pages",
          "Combines selected pages from multiple PDFs into one output PDF.",
          tool_combine_pdf_pages),
    "9": ("Extract PDF Pages",
          "Extracts selected pages from PDFs as separate files.",
          tool_extract_pdf_pages),
}

def main():
    script_dir = get_script_dir()
    os.chdir(script_dir)

    while True:
        print(MENU, end="")
        choice = input().strip().lower()

        if choice == "0":
            print("Goodbye.")
            break

        if choice not in TOOLS:
            print(f"\nInvalid Input: \"{choice}\"\n")
            continue

        name, desc, fn = TOOLS[choice]
        print(f"\n{'─' * 56}")
        print(f"  {name}")
        print(f"{'─' * 56}")
        print(desc)
        print()

        try:
            result = fn(script_dir)
        except Exception as e:
            print("ERROR: An unexpected error occurred.")
            print(f"Error Details: {e}")
            result = None

        if result == "CANCELLED":
            continue

        print("\n[Enter] - Continue to Menu")
        print("[0] - Exit Session")
        choice = input().strip()
        if choice == "0":
            sys.exit(0)

try:
    main()
except Exception as e:
    print("ERROR: An unexpected error occurred.")
    print(f"Error Details: {e}")
    input("\nPress Enter to close...")
