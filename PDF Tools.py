r"""
Master launcher for all PDF tools.
Tools: Convert to PDF, Merge PDF, Split PDF, Rotate PDF, Protect PDF,
       Unprotect PDF, Convert to PDF Precise, Selective Combine,
       Selective Extract

Every tool follows the same flow: pick files (or pages) -> confirm summary
-> process -> run report. Nothing is processed before the summary is confirmed.

Author: Yeoh Huey Sun
Email:  greyw901@gmail.com
GitHub: https://github.com/greyw901

Outputs are written next to the input files (same folder as the script/exe).
Exceptions: Split PDF -> Split\ , Selective Extract -> Extract\ (both can write many files).
"""

from __future__ import annotations

import datetime
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import traceback
import warnings
from dataclasses import dataclass, field
from itertools import count, product
from typing import Callable, Optional

# ══════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════
VERSION       = "v2.1_20261111"
LINE_WIDTH    = 56                       # banner / divider width
MAX_LISTED    = 20                       # longest numbered list printed before "... and n more"
EXCEL_EXTS    = {".xls", ".xlsx", ".xlsm", ".xlsb"}
WORD_EXTS     = {".doc", ".docx"}
PW_MARK       = "_pw_"                   # Protect PDF writes the password into the filename after this
BAD_NAME_CHARS = '\\/:*?"<>|'            # cannot appear in a Windows filename (password goes in the name)
SELECT_SPLIT  = re.compile(r"[\s,]+")
OUT_FOLDERS   = {"split": "Split", "extract": "Extract"}   # only tools listed here get an output subfolder


logging.getLogger("pypdf").setLevel(logging.ERROR)
warnings.filterwarnings("ignore")


# ══════════════════════════════════════════════════════════
#  SIGNALS (exceptions) - single catch site: run_tool() / main()
# ══════════════════════════════════════════════════════════
class UserCancelled(Exception):
    """User pressed [0] at a prompt: abandon the tool, back to the menu (no tail)."""


class NoInput(Exception):
    """Nothing to process (no files found). Message is printed, then the normal tail."""


class ExitSession(Exception):
    """[00] on any screen: end the session. Caught in main()."""


class SkipItem(Exception):
    """Raised by a worker: item needs no action. Reported under Skipped, not Failed."""


class GoBack(Exception):
    """[0] on a tool's input screen (not the first): one screen back.
    Single catch site: run_inputs()."""


# ══════════════════════════════════════════════════════════
#  SHARED: paths, naming
# ══════════════════════════════════════════════════════════
def get_script_dir() -> str:
    return os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, "frozen", False) else __file__
    ))


def get_unique_path(folder: str, name: str) -> str:
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


def out_dir(script_dir: str, tool_key: str) -> str:
    """Outputs go next to the inputs; only tools listed in OUT_FOLDERS (Split) get a subfolder."""
    return os.path.join(script_dir, OUT_FOLDERS[tool_key]) if tool_key in OUT_FOLDERS else script_dir


def list_files(folder: str, exts: set) -> list:
    return sorted(
        f for f in os.listdir(folder)
        if os.path.splitext(f)[1].lower() in exts
        and not f.startswith("~$")                       # Office lock files
        and os.path.isfile(os.path.join(folder, f))
    )


def prefix_generator():
    letters = "abcdefghijklmnopqrstuvwxyz"
    length = 1
    while True:
        for combo in product(letters, repeat=length):
            yield "".join(combo)
        length += 1


def number_generator():
    """File-level tools: files are picked by number (1, 2, 3 ...)."""
    return (str(n) for n in count(1))


def stem_of(name: str) -> str:
    return os.path.splitext(name)[0]


def dedupe(seq: list) -> list:
    seen, out = set(), []
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


# ══════════════════════════════════════════════════════════
#  SHARED: file index (codes a, a1, a2 ...) and selection
# ══════════════════════════════════════════════════════════
@dataclass(eq=False)                             # identity hash: entries go into sets
class Entry:
    prefix: str
    path: str
    name: str
    kind: str                                   # "pdf" | "excel" | "word"
    count: int = 0                              # pages / sheets (0 when not scanned)
    unit_label: str = ""                        # "Pages" / "Sheets"
    unit_word: str = ""                         # "page" / "sheet"
    unit_names: list = field(default_factory=list)


@dataclass
class Token:
    text: str
    prefix: Optional[str] = None                # set when the token is a whole file
    codes: list = field(default_factory=list)   # unit codes (unit-level selection only)


class Index:
    def __init__(self, entries: list) -> None:
        self.entries = entries
        self.by_prefix = {e.prefix: e for e in entries}
        self.code_map = {}
        for e in entries:
            for i in range(e.count):
                self.code_map[f"{e.prefix}{i + 1}"] = (e, i)

    def codes_of(self, prefix: str) -> list:
        e = self.by_prefix[prefix]
        return [f"{prefix}{i + 1}" for i in range(e.count)]

    def display_lines(self, level: str) -> list:
        lines = []
        for e in self.entries:
            if level == "file":
                lines.append(f'   [{e.prefix}] "{e.name}"')
            else:
                span = f"{e.prefix}1" if e.count == 1 else f"{e.prefix}1-{e.prefix}{e.count}"
                lines.append(f'   {e.prefix} {e.unit_label}: [{span}] - "{e.name}"')
                for i, n in enumerate(e.unit_names, start=1):
                    lines.append(f"       {e.prefix}{i} - {n}")
        return lines


def numbered(items: list, limit: Optional[int] = MAX_LISTED) -> list:
    shown = items if limit is None else items[:limit]
    lines = [f"   {i}. {s}" for i, s in enumerate(shown, start=1)]
    if limit is not None and len(items) > limit:
        lines.append(f"   ... and {len(items) - limit} more")
    return lines


def failure_lines(pairs: list) -> list:
    return [f'   {i}. "{n}"   :   {r}' for i, (n, r) in enumerate(pairs, start=1)]


# Screen grammar (script-conventions, Console rules 10-11), the same on every screen:
#   ====  / blank / heading + body / blank / ---- / [key] Label lines / blank / "> "
def open_screen() -> None:
    """Top of a new screen: the = line, then a blank line."""
    print("=" * LINE_WIDTH)
    print()


def divider(keys: list) -> None:
    """Blank, - line, key legend, blank: the input line never sits flush under the options."""
    print()
    print("-" * LINE_WIDTH)
    for k in keys:
        print(k)
    print()


def ask(valid: set, keys: list) -> str:
    """Legend once, then read until the answer is in `valid` (lower-cased). Never raises."""
    divider(keys)
    while True:
        raw = input("> ").strip().lower()
        if raw in valid:
            return raw
        print(f'Invalid Input: "{raw}"')


def nav_keys(first: bool) -> list:
    """Fixed vocabulary: [0] back one level (to the menu on a tool's first screen), [00] exit."""
    return ["[0] Back to menu" if first else "[0] Back", "[00] Exit"]


def check_nav(raw: str, first: bool) -> None:
    """[00] ends the session; [0] goes back (to the menu on the first screen). Checked before parsing."""
    if raw == "00":
        raise ExitSession()
    if raw == "0":
        if first:
            raise UserCancelled()
        raise GoBack()


def show_unreadable(unreadable: list) -> None:
    open_screen()
    print(f"Unreadable: {len(unreadable)}")
    print()
    for line in failure_lines(unreadable):
        print(line)
    choice = ask({"", "0", "00"}, ["[Enter] Skip & continue", "[0] Back to menu", "[00] Exit"])
    if choice:
        check_nav(choice, True)


def finalize_index(entries: list, unreadable: list, noun: str) -> Index:
    if not entries and not unreadable:
        raise NoInput(f"No {noun} files found.")
    if unreadable:
        show_unreadable(unreadable)
    if not entries:
        raise NoInput(f"No readable {noun} files found.")
    return Index(entries)


def scan_pdf_index(script_dir: str, count_pages: bool) -> Index:
    """Index the PDFs in the folder once. count_pages=False lists names only (file-level tools)."""
    from pypdf import PdfReader
    gen = prefix_generator() if count_pages else number_generator()
    entries, unreadable = [], []
    for name in list_files(script_dir, {".pdf"}):
        path = os.path.join(script_dir, name)
        if not count_pages:
            entries.append(Entry(next(gen), path, name, "pdf"))
            continue
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                raise ValueError("file is password-protected")
            n = len(reader.pages)
            if n == 0:
                raise ValueError("file has no pages")
        except Exception as e:
            unreadable.append((name, f"{type(e).__name__} - {e}"))
            continue
        entries.append(Entry(next(gen), path, name, "pdf", n, "Pages", "page"))
    return finalize_index(entries, unreadable, "PDF")


def parse_tokens(raw: str, index: Index, level: str) -> tuple:
    """level 'file': 1, 3, 4-6 (files by number).  level 'unit': a, a1, a1-a9, a1-9 (pages/sheets)."""
    tokens, invalid = [], []
    order = [e.prefix for e in index.entries]
    for text in filter(None, SELECT_SPLIT.split(raw.strip().lower())):
        if text in index.by_prefix:
            codes = index.codes_of(text) if level == "unit" else []
            tokens.append(Token(text, prefix=text, codes=codes))
            continue
        if level == "file":
            m = re.fullmatch(r"(\d+)-(\d+)", text)
            if m and m.group(1) in order and m.group(2) in order:
                i, j = order.index(m.group(1)), order.index(m.group(2))
                if i <= j:
                    tokens.extend(Token(p, prefix=p) for p in order[i:j + 1])
                    continue
            invalid.append(text)
            continue
        m = re.fullmatch(r"([a-z]+)(\d+)-(?:([a-z]+))?(\d+)", text)
        if m:
            p1, n1, p2, n2 = m.group(1), int(m.group(2)), m.group(3) or m.group(1), int(m.group(4))
            if p1 == p2 and p1 in index.by_prefix and 1 <= n1 <= n2 <= index.by_prefix[p1].count:
                tokens.append(Token(text, codes=[f"{p1}{n}" for n in range(n1, n2 + 1)]))
                continue
            invalid.append(text)
            continue
        if text in index.code_map:
            tokens.append(Token(text, codes=[text]))
            continue
        invalid.append(text)
    return tokens, invalid


def prompt_selection(index: Index, level: str, heading: str, example: str, first: bool = True) -> list:
    """Selection screen. Returns tokens. Invalid input is re-asked, never half-accepted."""
    open_screen()
    print(heading)
    print()
    print(f"Eg.: {example}")
    print()
    for line in index.display_lines(level):
        print(line)
    divider(["[#] Select", "[Enter] Select all"] + nav_keys(first))
    while True:
        raw = input("> ").strip()
        check_nav(raw, first)
        if not raw:
            return [Token(e.prefix, prefix=e.prefix,
                          codes=index.codes_of(e.prefix) if level == "unit" else [])
                    for e in index.entries]
        tokens, invalid = parse_tokens(raw, index, level)
        if invalid or not tokens:
            print("Invalid Input: " + ", ".join(f'"{t}"' for t in (invalid or [raw])))
            continue
        return tokens


def files_of(tokens: list, index: Index) -> list:
    return dedupe([index.by_prefix[t.prefix] for t in tokens if t.prefix])


def flatten_codes(tokens: list) -> list:
    return [c for t in tokens for c in t.codes]


def group_by_entry(codes: list, index: Index) -> list:
    """[(entry, sorted unique 0-based unit indexes)] in order of first appearance."""
    groups = {}
    for c in codes:
        e, i = index.code_map[c]
        groups.setdefault(e.prefix, (e, set()))[1].add(i)
    return [(e, sorted(idxs)) for e, idxs in groups.values()]


def ranges_text(idxs: list) -> str:
    """Sorted 0-based indexes -> '1-3, 5' (1-based, consecutive runs collapsed)."""
    runs, start, prev = [], None, None
    for i in idxs:
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            runs.append((start, prev))
            start = prev = i
    if start is not None:
        runs.append((start, prev))
    return ", ".join(f"{a + 1}" if a == b else f"{a + 1}-{b + 1}" for a, b in runs)


def unit_label(entry: Entry, idxs: list) -> str:
    return entry.unit_word.capitalize() if len(idxs) == 1 else entry.unit_label


def file_input_lines(entries: list) -> list:
    return [f'"{e.name}"' for e in entries]


def unit_input_lines(groups: list) -> list:
    """[(entry, idxs)] -> 'name   Pages 1-3, 5 of 12'. Same lines on the confirm screen and the run report."""
    return [f'"{e.name}"   {unit_label(e, idxs)} {ranges_text(idxs)} of {e.count}' for e, idxs in groups]


def ordered_groups(codes: list, index: Index) -> list:
    """Like group_by_entry but keeps the typed order: each consecutive run of one file is one group."""
    groups = []
    for c in codes:
        e, i = index.code_map[c]
        if groups and groups[-1][0] is e and groups[-1][1][-1] == i - 1:
            groups[-1][1].append(i)
        else:
            groups.append((e, [i]))
    return groups


def confirm(inputs: list, extra: Optional[list] = None) -> None:
    """Confirm screen: Inputs block (+ extra lines) then [Enter] / [0] / [00]."""
    open_screen()
    print(f"Confirm Inputs: {len(inputs)}")
    for line in numbered(inputs):
        print(line)
    for line in extra or []:
        print(line)
    choice = ask({"", "0", "00"}, ["[Enter] Confirm"] + nav_keys(False))
    if choice:
        check_nav(choice, False)


def prompt_password(heading: str, forbid_name_chars: bool) -> str:
    """Visible password screen. "0" and "00" are navigation keys, so neither can be a password."""
    open_screen()
    print(heading)
    divider(nav_keys(False))
    while True:
        pw = input("> ")
        check_nav(pw, False)
        if not pw:
            print('Invalid Input: ""')
        elif forbid_name_chars and any(c in BAD_NAME_CHARS for c in pw):
            print(f'Invalid Input: "{pw}"   :   password goes into the filename, cannot contain {BAD_NAME_CHARS}')
        else:
            return pw


def select_step(index: Index, level: str, heading: str, example: str) -> Callable:
    """Input step: pick files / pages. Stores the tokens in state['tokens']."""
    def step(state: dict) -> None:
        state["tokens"] = prompt_selection(index, level, heading, example, state["first"])
    return step


def run_inputs(steps: list, describe: Callable) -> dict:
    """Input screens in order, then the confirm screen. describe(state) -> (inputs, extra).
    [0] is one screen back (on the first screen it cancels), [00] cancels from anywhere.
    Single catch site of GoBack."""
    def confirm_step(state: dict) -> None:
        inputs, extra = describe(state)
        confirm(inputs, extra)

    steps = steps + [confirm_step]
    state, i = {}, 0
    while i < len(steps):
        state["first"] = i == 0
        try:
            steps[i](state)
            i += 1
        except GoBack:
            i = max(i - 1, 0)
    return state


# ══════════════════════════════════════════════════════════
#  SHARED: processing loop, run report, error log
# ══════════════════════════════════════════════════════════
def make_failure(item: str, step: str, context: str, exc: Exception) -> dict:
    # Must be called inside the except block: format_exc() is empty outside it.
    return {
        "item": item,
        "step": step,
        "error": f"{type(exc).__name__} - {exc}",
        "context": context,
        "trace": traceback.format_exc(),
    }


def run_items(items: list, step: str, worker: Callable, name_of: Callable,
              context_of: Optional[Callable] = None) -> tuple:
    """The loop prints nothing; the worker returns output names. -> (outputs, skipped, failures)"""
    outputs, skipped, failures = [], [], []
    for item in items:
        name = name_of(item)
        try:
            outputs.extend(worker(item))
        except SkipItem as s:
            skipped.append((name, str(s)))
        except Exception as e:
            ctx = context_of(item) if context_of else ""
            path = getattr(item, "path", None) or (item[0].path if isinstance(item, tuple) and hasattr(item[0], "path") else None)
            if path:                                    # file facts the log needs to diagnose without the input
                try:
                    size = f"{os.path.getsize(path)} bytes"
                except OSError:
                    size = "size unknown"
                ctx = "; ".join(x for x in (ctx, f'file = "{path}"', f"size = {size}") if x)
            failures.append(make_failure(name, step, ctx, e))
    return outputs, skipped, failures


def announce(script_dir: str, out_folder: str, inputs: list) -> None:
    """Run-report head. A folder line prints only when it carries information (the output is
    not in the script folder); then the Inputs block - printed before any work starts."""
    os.makedirs(out_folder, exist_ok=True)
    open_screen()
    if out_folder != script_dir:
        print(f'Output folder: "{out_folder}"')
        print()
    print(f"Inputs: {len(inputs)}")
    for line in numbered(inputs):
        print(line)


def begin(stage: str) -> float:
    """Stage line - call it as the statement right before the work it names."""
    print("\nProcessing")
    print(f"{stage} ...")
    return time.time()


def write_error_log(tool_name: str, failures: list, folder: str, out_folder: str) -> str:
    stamp = datetime.datetime.now().strftime("%Y%m%d")
    path = get_unique_path(folder, f"error log_{stamp}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"Tool: {tool_name}\n")
        f.write(f'Input folder: "{folder}"\n')
        f.write(f'Output folder: "{out_folder}"\n\n')
        for r in failures:
            f.write(f'"{r["item"]}"\n')
            f.write(f'  Step: {r["step"]}\n')
            f.write(f'  Error: {r["error"]}\n')
            f.write(f'  Context: {r["context"]}\n')
            f.write("  Traceback:\n")
            for line in r["trace"].rstrip().splitlines():
                f.write(f"    {line}\n")
            f.write("\n")
    return os.path.basename(path)


def tail(tool_name: str, script_dir: str, out_folder: str, failures: list) -> None:
    """Tool finished (suite): [Enter] back to menu, [0] save log (failures only), [00] exit."""
    keys, valid = ["[Enter] Back to menu"], {"", "00"}
    if failures:
        keys.append("[0] Save error log & back to menu")
        valid.add("0")
    keys.append("[00] Exit")
    choice = ask(valid, keys)
    if choice == "00":
        raise ExitSession()
    if choice == "0":
        print(f'Saved: "{write_error_log(tool_name, failures, script_dir, out_folder)}"')


def finish(tool_name: str, script_dir: str, out_folder: str, t0: float, outputs: list,
           skipped: list, failures: list) -> None:
    print(f"\nOutputs: {len(outputs)}")
    for line in numbered([f'"{o}"' for o in outputs]):
        print(line)
    if skipped:
        print(f"\nSkipped: {len(skipped)}")
        for line in failure_lines(skipped):
            print(line)
    if failures:
        print(f"\nFailed: {len(failures)}")
        for line in failure_lines([(r["item"], r["error"]) for r in failures]):
            print(line)
    print(f"\nElapsed: {time.time() - t0:.1f}s")
    tail(tool_name, script_dir, out_folder, failures)


def write_pdf(writer, path: str) -> None:
    with open(path, "wb") as f:
        writer.write(f)


# ══════════════════════════════════════════════════════════
#  SHARED: Office (COM) helpers - used by Convert to PDF / Precise
# ══════════════════════════════════════════════════════════
class OfficeApps:
    """Starts Excel / Word on first use, quits them in close(). One COM instance per app."""

    def __init__(self) -> None:
        self._apps = {}

    def get(self, kind: str):
        if kind not in self._apps:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application" if kind == "excel" else "Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            self._apps[kind] = app
        return self._apps[kind]

    def close(self) -> None:
        for app in self._apps.values():
            try:
                app.Quit()
            except Exception:
                pass
        self._apps = {}


def excel_export(app, src: str, out_path: str, keep_sheets: Optional[list]) -> None:
    wb = app.Workbooks.Open(src, ReadOnly=True, UpdateLinks=False)
    try:
        visible = [s.Name for s in wb.Sheets if s.Visible == -1]
        if not visible:
            raise ValueError("no visible sheets")
        if keep_sheets is not None:
            for s in wb.Sheets:                       # hide the rest; workbook closes unsaved
                if s.Visible == -1 and s.Name not in keep_sheets:
                    s.Visible = False
        wb.ExportAsFixedFormat(Type=0, Filename=out_path, Quality=0,
                               IncludeDocProperties=True,
                               IgnorePrintAreas=False, OpenAfterPublish=False)
    finally:
        wb.Close(SaveChanges=False)


def word_export(app, src: str, out_path: str, keep_pages: Optional[list]) -> None:
    """keep_pages: 0-based page indexes, or None for the whole document."""
    from pypdf import PdfReader, PdfWriter
    doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
    tmp_dir = tempfile.mkdtemp() if keep_pages is not None else None
    try:
        target = out_path if keep_pages is None else os.path.join(tmp_dir, "full.pdf")
        doc.ExportAsFixedFormat(OutputFileName=target, ExportFormat=17,
                                OpenAfterExport=False, OptimizeFor=0, Range=0,
                                IncludeDocProps=True, CreateBookmarks=0,
                                DocStructureTags=True)
        if keep_pages is not None:
            reader = PdfReader(target)
            writer = PdfWriter()
            for i in keep_pages:
                writer.add_page(reader.pages[i])
            write_pdf(writer, out_path)
    finally:
        doc.Close(SaveChanges=False)
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def scan_office_index(script_dir: str, with_units: bool) -> Index:
    """with_units=False: names only (Convert). True: opens each file for sheet / page counts (Precise)."""
    gen = prefix_generator() if with_units else number_generator()
    entries, unreadable = [], []
    apps = OfficeApps()
    try:
        for name in list_files(script_dir, EXCEL_EXTS | WORD_EXTS):
            path = os.path.join(script_dir, name)
            kind = "excel" if os.path.splitext(name)[1].lower() in EXCEL_EXTS else "word"
            if not with_units:
                entries.append(Entry(next(gen), path, name, kind))
                continue
            try:
                if kind == "excel":
                    wb = apps.get("excel").Workbooks.Open(path, ReadOnly=True, UpdateLinks=False)
                    try:
                        visible = [s.Name for s in wb.Sheets if s.Visible == -1]
                    finally:
                        wb.Close(SaveChanges=False)
                    if not visible:
                        raise ValueError("no visible sheets")
                    entries.append(Entry(next(gen), path, name, kind, len(visible),
                                         "Sheets", "sheet", visible))
                else:
                    doc = apps.get("word").Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
                    try:
                        pages = int(doc.ComputeStatistics(2))          # wdStatisticPages
                    finally:
                        doc.Close(SaveChanges=False)
                    entries.append(Entry(next(gen), path, name, kind, pages, "Pages", "page"))
            except Exception as e:
                unreadable.append((name, f"{type(e).__name__} - {e}"))
    finally:
        apps.close()
    return finalize_index(entries, unreadable, "Excel or Word")


# ══════════════════════════════════════════════════════════
#  TOOL 1 - Convert to PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf(script_dir: str) -> None:
    index = scan_office_index(script_dir, with_units=False)
    state = run_inputs([select_step(index, "file", "Select files to convert", "1, 3, 4-6")],
                       lambda st: (file_input_lines(files_of(st["tokens"], index)), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "convert")
    announce(script_dir, out_folder, file_input_lines(entries))
    apps = OfficeApps()

    def worker(entry: Entry) -> list:
        out_path = get_unique_path(out_folder, stem_of(entry.name) + ".pdf")
        if entry.kind == "excel":
            excel_export(apps.get("excel"), entry.path, out_path, None)
        else:
            word_export(apps.get("word"), entry.path, out_path, None)
        return [os.path.basename(out_path)]

    try:
        t0 = begin("Converting")
        outputs, skipped, failures = run_items(entries, "converting to PDF", worker, lambda e: e.name)
    finally:
        apps.close()
    finish("Convert to PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 2 - Merge PDF   (file-level; merge order = order selected)
# ══════════════════════════════════════════════════════════
def tool_merge_pdf(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=False)
    state = run_inputs(
        [select_step(index, "file", "Select files to merge", "3, 1, 4-6        (merged in the order typed)")],
        lambda st: (file_input_lines(files_of(st["tokens"], index)), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "merge")
    announce(script_dir, out_folder, file_input_lines(entries))
    writer = PdfWriter()
    merged = []

    def worker(entry: Entry) -> list:
        reader = PdfReader(entry.path)
        if reader.is_encrypted:
            raise ValueError("file is password-protected")
        pages = list(reader.pages)                  # parse everything before touching the writer
        for p in pages:
            writer.add_page(p)
        merged.append(entry)
        return []

    t0 = begin("Merging")
    _, skipped, failures = run_items(entries, "reading pages", worker, lambda e: e.name)
    outputs = []
    if merged:
        out_name = f"{stem_of(merged[0].name)} Merged.pdf"
        try:
            out_path = get_unique_path(out_folder, out_name)
            write_pdf(writer, out_path)
            outputs.append(os.path.basename(out_path))
        except Exception as e:
            failures.append(make_failure(out_name, "writing merged file", f"files merged = {len(merged)}", e))
    finish("Merge PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 3 - Split PDF   (file-level; one file per page)
# ══════════════════════════════════════════════════════════
def tool_split_pdf(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=False)
    state = run_inputs([select_step(index, "file", "Select files to split", "1, 3, 4-6")],
                       lambda st: (file_input_lines(files_of(st["tokens"], index)), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "split")
    announce(script_dir, out_folder, file_input_lines(entries))

    def worker(entry: Entry) -> list:
        reader = PdfReader(entry.path)
        if reader.is_encrypted:
            raise ValueError("file is password-protected")
        names = []
        for i, page in enumerate(reader.pages, start=1):
            w = PdfWriter()
            w.add_page(page)
            out_path = get_unique_path(out_folder, f"{stem_of(entry.name)} p{i}.pdf")
            write_pdf(w, out_path)
            names.append(os.path.basename(out_path))
        return names

    t0 = begin("Splitting")
    outputs, skipped, failures = run_items(entries, "reading and splitting pages", worker, lambda e: e.name)
    finish("Split PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 4 - Rotate PDF   (page-level)
# ══════════════════════════════════════════════════════════
def prompt_rotation() -> tuple:
    """Rotation screen -> (clockwise degrees for pypdf, label such as 'R90'). Left N == right (360 - N)."""
    open_screen()
    print("Rotation by: [L / R] [90 / 180]")
    print()
    print("Eg.: R90")
    divider(["[L/R][90/180] Select"] + nav_keys(False))
    while True:
        raw = input("> ").strip()
        check_nav(raw, False)
        m = re.fullmatch(r"([LR])\s*(90|180)", raw.upper())
        if m:
            direction, degrees = m.group(1), int(m.group(2))
            return (degrees if direction == "R" else 360 - degrees), f"{direction}{degrees}"
        print(f'Invalid Input: "{raw}"')


def tool_rotate_pdf(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=True)

    def rotation_step(st: dict) -> None:
        st["angle"], st["label"] = prompt_rotation()

    def describe(st: dict) -> tuple:
        st["groups"] = group_by_entry(dedupe(flatten_codes(st["tokens"])), index)
        return unit_input_lines(st["groups"]), [f"\nRotation: {st['label']}"]

    state = run_inputs([select_step(index, "unit", "Select pages to rotate", "a1, b2, c4-6"),
                        rotation_step], describe)
    angle, label, groups = state["angle"], state["label"], state["groups"]

    out_folder = out_dir(script_dir, "rotate")
    announce(script_dir, out_folder, unit_input_lines(groups))
    print(f"\nRotation: {label}")

    def worker(item: tuple) -> list:
        entry, idxs = item
        reader = PdfReader(entry.path)
        writer = PdfWriter()
        for i, page in enumerate(reader.pages):
            if i in idxs:
                page.rotate(angle)
            writer.add_page(page)
        out_path = get_unique_path(out_folder, f"{stem_of(entry.name)} Rotated.pdf")
        write_pdf(writer, out_path)
        return [os.path.basename(out_path)]

    t0 = begin("Rotating")
    outputs, skipped, failures = run_items(
        groups, "rotating pages", worker, lambda it: it[0].name,
        lambda it: f"pages selected = {[i + 1 for i in it[1]]}; rotation = {label}")
    finish("Rotate PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 5 - Protect PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_protect_pdf(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=False)

    def password_step(st: dict) -> None:
        st["password"] = prompt_password("Password to apply", forbid_name_chars=True)

    state = run_inputs(
        [select_step(index, "file", "Select files to protect", "1, 3, 4-6"), password_step],
        lambda st: (file_input_lines(files_of(st["tokens"], index)), [f"\nPassword: {st['password']}"]))
    entries, password = files_of(state["tokens"], index), state["password"]

    out_folder = out_dir(script_dir, "protect")
    announce(script_dir, out_folder, file_input_lines(entries))

    def worker(entry: Entry) -> list:
        stem = stem_of(entry.name)
        if PW_MARK in stem:
            raise SkipItem(f"already protected (filename has {PW_MARK})")
        reader = PdfReader(entry.path)
        if reader.is_encrypted:
            raise SkipItem("already protected (file is encrypted)")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)
        # Password in the filename is a deliberate choice (ADR), despite the underscore rule.
        out_path = get_unique_path(out_folder, f"{stem}{PW_MARK}{password}.pdf")
        write_pdf(writer, out_path)
        return [os.path.basename(out_path)]

    t0 = begin("Protecting")
    outputs, skipped, failures = run_items(entries, "reading and encrypting", worker, lambda e: e.name)
    finish("Protect PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 6 - Unprotect PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_unprotect_pdf(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=False)

    def password_step(st: dict) -> None:
        st["password"] = prompt_password("Password to unlock", forbid_name_chars=False)

    state = run_inputs(
        [select_step(index, "file", "Select files to unprotect", "1, 3, 4-6"), password_step],
        lambda st: (file_input_lines(files_of(st["tokens"], index)), [f"\nPassword: {st['password']}"]))
    entries, password = files_of(state["tokens"], index), state["password"]

    out_folder = out_dir(script_dir, "unprotect")
    announce(script_dir, out_folder, file_input_lines(entries))

    def worker(entry: Entry) -> list:
        reader = PdfReader(entry.path)
        if not reader.is_encrypted:
            raise SkipItem("no password")
        if reader.decrypt(password) == 0:
            raise ValueError("wrong password, could not decrypt")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        stem = stem_of(entry.name)
        base = re.sub(re.escape(PW_MARK) + r".*$", "", stem) or stem      # drop the password hint
        out_path = get_unique_path(out_folder, base + ".pdf")
        write_pdf(writer, out_path)
        return [os.path.basename(out_path)]

    t0 = begin("Unprotecting")
    outputs, skipped, failures = run_items(entries, "opening and decrypting", worker, lambda e: e.name)
    finish("Unprotect PDF", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 7 - Convert to PDF Precise   (unit-level: Excel sheets, Word pages)
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf_precise(script_dir: str) -> None:
    index = scan_office_index(script_dir, with_units=True)

    def describe(st: dict) -> tuple:
        st["groups"] = group_by_entry(dedupe(flatten_codes(st["tokens"])), index)
        return unit_input_lines(st["groups"]), []

    state = run_inputs([select_step(index, "unit", "Select sheets or pages to convert", "a1, b, c2-3")], describe)
    groups = state["groups"]

    out_folder = out_dir(script_dir, "convert")
    announce(script_dir, out_folder, unit_input_lines(groups))
    apps = OfficeApps()

    def worker(item: tuple) -> list:
        entry, idxs = item
        whole = len(idxs) == entry.count                      # everything selected: no filtering
        out_path = get_unique_path(out_folder, stem_of(entry.name) + ".pdf")
        if entry.kind == "excel":
            keep = None if whole else [entry.unit_names[i] for i in idxs]
            excel_export(apps.get("excel"), entry.path, out_path, keep)
        else:
            word_export(apps.get("word"), entry.path, out_path, None if whole else idxs)
        return [os.path.basename(out_path)]

    try:
        t0 = begin("Converting")
        outputs, skipped, failures = run_items(
            groups, "converting selection to PDF", worker, lambda it: it[0].name,
            lambda it: f"{it[0].unit_word}s selected = {[i + 1 for i in it[1]]} of {it[0].count}")
    finally:
        apps.close()
    finish("Convert to PDF Precise", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 8 - Selective Combine   (unit-level; one output, order as selected)
# ══════════════════════════════════════════════════════════
def tool_selective_combine(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=True)

    def describe(st: dict) -> tuple:
        st["codes"] = flatten_codes(st["tokens"])             # order and repeats kept on purpose
        return unit_input_lines(ordered_groups(st["codes"], index)), [f"\nOutput pages: {len(st['codes'])}"]

    state = run_inputs(
        [select_step(index, "unit", "Select pages to combine", "a1, b, c4-6        (combined in the order typed)")],
        describe)
    codes = state["codes"]

    out_folder = out_dir(script_dir, "combine")
    announce(script_dir, out_folder, unit_input_lines(ordered_groups(codes, index)))
    writer = PdfWriter()
    readers = {}
    added = []

    def worker(code: str) -> list:
        entry, idx = index.code_map[code]
        if entry.path not in readers:
            readers[entry.path] = PdfReader(entry.path)
        writer.add_page(readers[entry.path].pages[idx])
        added.append(code)
        return []

    t0 = begin("Combining")
    _, skipped, failures = run_items(
        codes, "adding pages", worker,
        lambda c: f'{index.code_map[c][0].name} {c}',
        lambda c: f"page = {index.code_map[c][1] + 1} of {index.code_map[c][0].count}")
    outputs = []
    if added:
        n = len(added)
        out_name = f"Combined {n} page.pdf" if n == 1 else f"Combined {n} pages.pdf"
        try:
            out_path = get_unique_path(out_folder, out_name)
            write_pdf(writer, out_path)
            outputs.append(os.path.basename(out_path))
        except Exception as e:
            failures.append(make_failure(out_name, "writing combined file", f"pages added = {n}", e))
    finish("Selective Combine", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  TOOL 9 - Selective Extract   (unit-level; each token = one output file)
# ══════════════════════════════════════════════════════════
def tool_selective_extract(script_dir: str) -> None:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, count_pages=True)

    def make_job(codes: list) -> tuple:
        entry = index.code_map[codes[0]][0]
        idxs = [index.code_map[c][1] for c in codes]
        label = f"p{idxs[0] + 1}" if len(idxs) == 1 else f"p{idxs[0] + 1}-{idxs[-1] + 1}"
        return entry, idxs, f"{stem_of(entry.name)} {label}.pdf"

    def jobs_of(tokens: list) -> list:
        jobs = []
        for t in tokens:
            if t.prefix:                                       # whole file -> one output per page
                jobs.extend(make_job([c]) for c in t.codes)
            else:
                jobs.append(make_job(t.codes))
        return jobs

    def describe(st: dict) -> tuple:
        st["jobs"] = jobs_of(st["tokens"])
        groups = group_by_entry(flatten_codes(st["tokens"]), index)
        extra = [f"\nOutput files: {len(st['jobs'])}"] + numbered([f'"{name}"' for _, _, name in st["jobs"]])
        return unit_input_lines(groups), extra

    state = run_inputs(
        [select_step(index, "unit", "Select pages to extract", "a1, b2, c4-6        (each entry = one file)")], describe)
    jobs = state["jobs"]

    out_folder = out_dir(script_dir, "extract")
    announce(script_dir, out_folder, unit_input_lines(group_by_entry(flatten_codes(state["tokens"]), index)))
    readers = {}

    def worker(job: tuple) -> list:
        entry, idxs, name = job
        if entry.path not in readers:
            readers[entry.path] = PdfReader(entry.path)
        w = PdfWriter()
        for i in idxs:
            w.add_page(readers[entry.path].pages[i])
        out_path = get_unique_path(out_folder, name)
        write_pdf(w, out_path)
        return [os.path.basename(out_path)]

    t0 = begin("Extracting")
    outputs, skipped, failures = run_items(
        jobs, "extracting pages", worker, lambda j: j[2],
        lambda j: f'source = "{j[0].name}"; pages = {[i + 1 for i in j[1]]}')
    finish("Selective Extract", script_dir, out_folder, t0, outputs, skipped, failures)


# ══════════════════════════════════════════════════════════
#  MENU
# ══════════════════════════════════════════════════════════
TOOLS = {
    "1": ("Convert to PDF", tool_convert_to_pdf),
    "2": ("Merge PDF", tool_merge_pdf),
    "3": ("Split PDF", tool_split_pdf),
    "4": ("Rotate PDF", tool_rotate_pdf),
    "5": ("Protect PDF", tool_protect_pdf),
    "6": ("Unprotect PDF", tool_unprotect_pdf),
    "7": ("Convert to PDF Precise", tool_convert_to_pdf_precise),
    "8": ("Selective Combine", tool_selective_combine),
    "9": ("Selective Extract", tool_selective_extract),
}


def build_menu() -> str:
    """The boxed two-column menu: BASIC (1-6) on the left, ADVANCE (7-9) on the right."""
    w = 31                                       # inner width of each column
    bar = "+" + "-" * w + "+" + "-" * w + "+"
    full = "+" + "-" * (2 * w + 1) + "+"

    def row(left: str, right: str) -> str:
        return f"| {left:<{w - 1}}| {right:<{w - 1}}|"

    cell = lambda k: f"[{k}] - {TOOLS[k][0]}"
    lefts = [cell(k) for k in "123456"]
    rights = [cell(k) for k in "789"]
    lines = [bar, row(f"  PDF Tools {VERSION}", ""), bar, row("BASIC", "ADVANCE"), bar]
    for i, left in enumerate(lefts):
        lines.append(row(left, rights[i] if i < len(rights) else ""))
    lines += [bar, f"| {'[00] - Exit':<{2 * w}}|", full]
    return chr(10).join(lines)


def choose_tool() -> str:
    """Menu screen. Returns '00' (exit) or a key of TOOLS. No [0]: there is no previous level."""
    open_screen()
    print(build_menu())
    return ask({"00"} | set(TOOLS), ["[#] Select", "[00] Exit"])


def run_tool(name: str, fn: Callable, script_dir: str) -> None:
    """One tool run. UserCancelled ends quietly (back to the menu); ExitSession propagates to main()."""
    try:
        fn(script_dir)
    except UserCancelled:
        return
    except ExitSession:
        raise
    except NoInput as e:
        open_screen()
        print(e)
        tail(name, script_dir, script_dir, [])
    except Exception as e:
        failures = [make_failure(name, "running tool", "", e)]
        open_screen()
        print("FAILED: tool stopped unexpectedly")
        for line in failure_lines([(failures[0]["item"], failures[0]["error"])]):
            print(line)
        tail(name, script_dir, script_dir, failures)


def main() -> None:
    script_dir = get_script_dir()
    os.chdir(script_dir)
    for stream in (sys.stdout, sys.stderr):      # a filename the console cannot encode must not crash a tool
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    while True:
        choice = choose_tool()
        if choice == "00":
            break
        name, fn = TOOLS[choice]
        try:
            run_tool(name, fn, script_dir)
        except ExitSession:
            break

    # Every exit ends here, so a double-clicked window never vanishes unread
    # (which is why no extra "are you sure" is needed before it).
    ask({""}, ["[Enter] Close"])


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FAILED: {type(e).__name__} - {e}")
        try:
            input("\n[Enter] Close\n\n> ")
        except Exception:
            pass
