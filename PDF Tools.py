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
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import warnings
from dataclasses import dataclass, field
from itertools import count, product
from typing import Callable, Optional

# ══════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════
SUITE_NAME    = "PDF Tools"
VERSION       = "v2.2_20261113"          # not shown on screen; goes into the exe version file
LINE_WIDTH    = 56                       # banner / divider width
MAX_LISTED    = 20                       # longest numbered list printed before "... and n more"
EXCEL_EXTS    = {".xls", ".xlsx", ".xlsm", ".xlsb"}
WORD_EXTS     = {".doc", ".docx"}
PW_MARK       = "_pw_"                   # Protect PDF writes the password into the filename after this
BAD_NAME_CHARS = '\\/:*?"<>|'            # cannot appear in a Windows filename (password goes in the name)
SELECT_SPLIT  = re.compile(r"[\s,]+")
OUT_FOLDERS   = {"split": "Split", "extract": "Extract"}   # only tools listed here get an output subfolder
OFFICE_TIMEOUT = 300                     # s; one Excel/Word open+export with no response -> force-closed
NO_WINDOW     = 0x08000000               # CREATE_NO_WINDOW: helper commands don't flash a console


logging.getLogger("pypdf").setLevel(logging.ERROR)
warnings.filterwarnings("ignore")


# ══════════════════════════════════════════════════════════
#  SIGNALS (exceptions) - each has one catch site, named below
# ══════════════════════════════════════════════════════════
class UserCancelled(Exception):
    """[0] on a tool's first screen: abandon the tool, back to the menu (no tail).
    Catch site: run_tool()."""


class NoInput(Exception):
    """Nothing to process (no files found). Message is printed, then the normal tail.
    Catch site: run_tool()."""


class ExitSession(Exception):
    """[00] on a tool screen: Cancelled. screen, then the program closes. Catch site: main()."""


class ExitProgram(Exception):
    """[0] Exit on the Tool finished screen (no failures): close immediately, no screen.
    Catch site: main()."""


class InvalidInput(Exception):
    """Raised by a screen's input handler; message = reason ("" for a bare notice).
    Catch site: prompt_screen(), which redraws the same screen with the notice."""


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
#   clear / banner / choices so far / notice / heading + body / ---- / [key] Label lines / blank / prompt
def clear_screen() -> None:
    os.system("cls")


def open_screen(title: str, choices: Optional[list] = None, notice: Optional[str] = None) -> None:
    """New screen: clear, banner, the choices made so far, then the notice left by the
    previous input (Error: Invalid Input ... / Saved: ...)."""
    clear_screen()
    print("=" * LINE_WIDTH)
    print(title)
    print("=" * LINE_WIDTH)
    print()
    if choices:
        for line in choices:
            print(line)
        print()
    if notice:
        print(notice)
        print()


def options(keys: list) -> None:
    """Blank, - line, key legend, blank: the input line never sits flush under the options."""
    print()
    print("-" * LINE_WIDTH)
    for k in keys:
        print(k)
    print()


def prompt_screen(title: str, body: Callable, keys: list, prompt: str, handle: Callable,
                  choices: Optional[list] = None, notice: Optional[str] = None,
                  strip: bool = True):
    """Draw the screen, read one answer, return handle(answer). handle raises InvalidInput
    to have the same screen redrawn with the notice."""
    while True:
        open_screen(title, choices, notice)
        body()
        options(keys)
        raw = input(prompt)
        try:
            return handle(raw.strip() if strip else raw)
        except InvalidInput as e:
            notice = f"Error: Invalid Input - {e}" if str(e) else "Error: Invalid Input"


def nav_keys(first: bool) -> list:
    """Fixed vocabulary: [0] back one level (to the menu on a tool's first screen),
    [00] cancel & exit."""
    return ["[0] Back to menu" if first else "[0] Back", "[00] Cancel & exit"]


def check_nav(raw: str, first: bool) -> None:
    """[00] ends the session; [0] goes back (to the menu on the first screen). Checked before parsing."""
    if raw == "00":
        raise ExitSession()
    if raw == "0":
        if first:
            raise UserCancelled()
        raise GoBack()


def show_unreadable(title: str, unreadable: list) -> None:
    def body() -> None:
        print(f"Unreadable: {len(unreadable)}")
        for line in failure_lines(unreadable):
            print(line)

    def handle(raw: str) -> None:
        if raw:
            check_nav(raw, True)
            raise InvalidInput("")

    prompt_screen(title, body, ["[Enter] Skip & continue"] + nav_keys(True), "Ans: ", handle)


def finalize_index(entries: list, unreadable: list, noun: str, title: str) -> Index:
    if not entries and not unreadable:
        raise NoInput(f"No {noun} files found.")
    if unreadable:
        show_unreadable(title, unreadable)
    if not entries:
        raise NoInput(f"No readable {noun} files found.")
    return Index(entries)


def scan_pdf_index(script_dir: str, count_pages: bool, title: str) -> Index:
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
    return finalize_index(entries, unreadable, "PDF", title)


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


def prompt_selection(title: str, index: Index, level: str, heading: str, example: str,
                     noun: str, first: bool, choices: list) -> list:
    """Selection screen. Returns tokens. Invalid input redraws the screen, never half-accepted."""
    def body() -> None:
        print(heading)
        print()
        print(f"e.g. {example}")
        print()
        for line in index.display_lines(level):
            print(line)

    def handle(raw: str) -> list:
        check_nav(raw, first)
        if not raw:
            return [Token(e.prefix, prefix=e.prefix,
                          codes=index.codes_of(e.prefix) if level == "unit" else [])
                    for e in index.entries]
        tokens, invalid = parse_tokens(raw, index, level)
        if invalid or not tokens:
            raise InvalidInput("not valid: " + ", ".join(f'"{t}"' for t in (invalid or [raw])))
        return tokens

    return prompt_screen(title, body, ["[Enter] Select all", "[#] Select"] + nav_keys(first),
                         f"{noun}: ", handle, choices)


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


def inputs_block(lines: list) -> list:
    return [f"Inputs: {len(lines)}"] + numbered(lines)


def confirm(title: str, inputs: list, extra: Optional[list] = None) -> None:
    """Confirm screen: Inputs block (+ extra lines) then [Enter] / [0] / [00]."""
    def body() -> None:
        print(f"Confirm Inputs: {len(inputs)}")
        for line in numbered(inputs):
            print(line)
        for line in extra or []:
            print(line)

    def handle(raw: str) -> None:
        if raw:
            check_nav(raw, False)
            raise InvalidInput("")

    prompt_screen(title, body, ["[Enter] Continue"] + nav_keys(False), "Ans: ", handle)


def prompt_password(title: str, heading: str, forbid_name_chars: bool, choices: list) -> str:
    """Visible password screen. "0" and "00" are navigation keys, so neither can be a password."""
    def handle(pw: str) -> str:
        check_nav(pw, False)
        if not pw:
            raise InvalidInput("")
        if forbid_name_chars and any(c in BAD_NAME_CHARS for c in pw):
            raise InvalidInput(f"password goes into the filename, cannot contain {BAD_NAME_CHARS}")
        return pw

    return prompt_screen(title, lambda: print(heading), nav_keys(False), "Password: ", handle,
                         choices, strip=False)


def select_step(title: str, index: Index, level: str, heading: str, example: str,
                noun: str, lines_of: Callable) -> tuple:
    """Input step: pick files / pages into state['tokens']. lines_of(tokens) -> the Inputs
    lines this choice shows on later screens."""
    def step(state: dict, first: bool, choices: list) -> None:
        state["tokens"] = prompt_selection(title, index, level, heading, example, noun, first, choices)

    def shown(state: dict) -> list:
        return inputs_block(lines_of(state["tokens"]))

    return step, shown


def run_inputs(title: str, steps: list, describe: Callable) -> dict:
    """steps: [(screen(state, first, choices), shown(state) -> lines)], then the confirm
    screen; describe(state) -> (inputs, extra). Each screen shows the choices of the steps
    before it; the confirm screen shows none - everything chosen is its body.
    [0] is one screen back (on the first screen: to the menu). Single catch site of GoBack."""
    state, i = {}, 0
    while True:
        try:
            if i == len(steps):
                inputs, extra = describe(state)
                confirm(title, inputs, extra)
                return state
            choices = []
            for _, shown in steps[:i]:
                if choices:
                    choices.append("")
                choices.extend(shown(state))
            steps[i][0](state, i == 0, choices)
            i += 1
        except GoBack:
            i -= 1


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


class Transcript:
    """Run-report lines: printed live, and kept so the Tool finished screen can be
    redrawn after invalid input without losing the report."""

    def __init__(self) -> None:
        self.lines = []

    def say(self, text: str = "") -> None:
        print(text, flush=True)
        self.lines.append(text)


def announce(title: str, script_dir: str, out_folder: str, inputs: list) -> Transcript:
    """Run-report head. A folder line prints only when it carries information (the output is
    not in the script folder); then the Inputs block - printed before any work starts."""
    os.makedirs(out_folder, exist_ok=True)
    open_screen(title)
    tr = Transcript()
    if out_folder != script_dir:
        tr.say(f'Output folder: "{out_folder}"')
        tr.say()
    tr.say(f"Inputs: {len(inputs)}")
    for line in numbered(inputs):
        tr.say(line)
    return tr


def begin(tr: Transcript, stage: str) -> float:
    """Stage line - call it as the statement right before the work it names."""
    tr.say()
    tr.say("Processing")
    tr.say(f"{stage} ...")
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


def tail(title: str, script_dir: str, out_folder: str, failures: list,
         tr: Transcript) -> Optional[str]:
    """Tool finished (suite): [Enter] back to menu; [0] save log & back to menu when there
    are failures, otherwise [0] Exit (raises ExitProgram). Returns the notice for the menu
    screen (Saved: ...) or None."""
    keys = ["[Enter] Back to menu"]
    keys.append("[0] Save error log & back to menu" if failures else "[0] Exit")

    def body() -> None:
        for line in tr.lines:
            print(line)

    def handle(raw: str) -> Optional[str]:
        if raw == "":
            return None
        if raw == "0" and failures:
            return f'Saved: "{write_error_log(title, failures, script_dir, out_folder)}"'
        if raw == "0":
            raise ExitProgram()
        raise InvalidInput("")

    return prompt_screen(title, body, keys, "Ans: ", handle)


def finish(title: str, script_dir: str, out_folder: str, t0: float, outputs: list,
           skipped: list, failures: list, tr: Transcript) -> Optional[str]:
    tr.say()
    tr.say(f"Outputs: {len(outputs)}")
    for line in numbered([f'"{o}"' for o in outputs]):
        tr.say(line)
    if skipped:
        tr.say()
        tr.say(f"Skipped: {len(skipped)}")
        for line in failure_lines(skipped):
            tr.say(line)
    if failures:
        tr.say()
        tr.say(f"Failed: {len(failures)}")
        for line in failure_lines([(r["item"], r["error"]) for r in failures]):
            tr.say(line)
    tr.say()
    tr.say(f"Elapsed: {time.time() - t0:.1f}s")
    return tail(title, script_dir, out_folder, failures, tr)


def write_pdf(writer, path: str) -> None:
    with open(path, "wb") as f:
        writer.write(f)


# ══════════════════════════════════════════════════════════
#  SHARED: Office (COM) helpers - used by Convert to PDF / Precise
# ══════════════════════════════════════════════════════════
class OfficePrompt(Exception):
    """Our hidden Excel/Word showed a prompt (or stopped responding) and was force-closed.
    Raised by OfficeApps.run(); lands under Failed / Unreadable like any per-item error."""


def _office_pids(exe: str) -> set:
    out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {exe}", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, creationflags=NO_WINDOW).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.strip().strip('"').split('","')
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _visible_windows(pid: int) -> list:
    """Visible, titled top-level windows of pid. Our app runs with Visible=False, so any
    such window is a prompt waiting for a click (DisplayAlerts does not cover all of them,
    e.g. Word's "it caused a serious error last time")."""
    import win32gui
    import win32process
    found = []

    def cb(h, _) -> None:
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            if win32process.GetWindowThreadProcessId(h)[1] == pid:
                found.append(win32gui.GetWindowText(h))
    win32gui.EnumWindows(cb, None)
    return found


# Office draws prompt text itself (no Win32 text); UI Automation via PowerShell can read it.
_PS_PROMPT_TEXT = (
    "Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes;"
    "$A=[System.Windows.Automation.AutomationElement];"
    "$c=New-Object System.Windows.Automation.PropertyCondition($A::ProcessIdProperty,{pid});"
    "$t=New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty,"
    "[System.Windows.Automation.ControlType]::Text);"
    "foreach($w in $A::RootElement.FindAll('Children',$c)){{"
    "foreach($e in $w.FindAll('Descendants',$t)){{$e.Current.Name}}}}"
)


def _prompt_text(pid: int, fallback: str) -> str:
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", _PS_PROMPT_TEXT.format(pid=pid)],
                           capture_output=True, text=True, timeout=20, creationflags=NO_WINDOW)
        text = " ".join(line.strip() for line in r.stdout.splitlines() if line.strip())
        return text or fallback
    except Exception:
        return fallback


class OfficeApps:
    """Starts Excel / Word on first use, quits them in close(). One COM instance per app.
    Every COM job goes through run(), which a watchdog guards: a prompt or a hang
    force-closes OUR instance (never the user's) and raises OfficePrompt."""

    EXE = {"excel": "EXCEL.EXE", "word": "WINWORD.EXE"}

    def __init__(self) -> None:
        self._apps = {}
        self._pids = {}

    def get(self, kind: str):
        if kind not in self._apps:
            import win32com.client
            before = _office_pids(self.EXE[kind])
            # DispatchEx: always a fresh instance - Dispatch can attach to the user's open
            # Excel/Word, which close() would then Quit (DisplayAlerts off = unsaved work lost).
            app = win32com.client.DispatchEx("Excel.Application" if kind == "excel" else "Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            new = _office_pids(self.EXE[kind]) - before
            self._pids[kind] = next(iter(new)) if len(new) == 1 else None   # None: unguarded
            self._apps[kind] = app
        return self._apps[kind]

    def run(self, kind: str, job: Callable):
        """job(app) under the watchdog. On a prompt / hang: kill our instance, raise OfficePrompt;
        the next run() starts a fresh instance."""
        app = self.get(kind)
        pid = self._pids[kind]
        done, hit = threading.Event(), {}
        name = "Word" if kind == "word" else "Excel"

        def watch() -> None:
            start = time.time()
            while not done.wait(1):
                windows = _visible_windows(pid)
                if windows:
                    hit["why"] = f"{name} showed a prompt: '{_prompt_text(pid, windows[0])}'"
                elif time.time() - start > OFFICE_TIMEOUT:
                    hit["why"] = f"{name} did not respond for {OFFICE_TIMEOUT}s"
                else:
                    continue
                subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True,
                               creationflags=NO_WINDOW)
                return

        watcher = threading.Thread(target=watch, daemon=True) if pid else None
        if watcher:
            watcher.start()
        try:
            return job(app)
        except Exception:
            if hit:
                raise OfficePrompt(hit["why"] + f" - {name} was closed") from None
            raise
        finally:
            done.set()
            if watcher:
                watcher.join()
            if hit:
                self._apps.pop(kind, None)
                self._pids.pop(kind, None)

    def close(self) -> None:
        for app in self._apps.values():
            try:
                app.Quit()
            except Exception:
                pass
        self._apps = {}
        self._pids = {}


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


def excel_visible_sheets(app, path: str) -> list:
    wb = app.Workbooks.Open(path, ReadOnly=True, UpdateLinks=False)
    try:
        return [s.Name for s in wb.Sheets if s.Visible == -1]
    finally:
        wb.Close(SaveChanges=False)


def word_page_count(app, path: str) -> int:
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        return int(doc.ComputeStatistics(2))                    # wdStatisticPages
    finally:
        doc.Close(SaveChanges=False)


def scan_office_index(script_dir: str, with_units: bool, title: str) -> Index:
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
                    visible = apps.run("excel", lambda app: excel_visible_sheets(app, path))
                    if not visible:
                        raise ValueError("no visible sheets")
                    entries.append(Entry(next(gen), path, name, kind, len(visible),
                                         "Sheets", "sheet", visible))
                else:
                    pages = apps.run("word", lambda app: word_page_count(app, path))
                    entries.append(Entry(next(gen), path, name, kind, pages, "Pages", "page"))
            except Exception as e:
                unreadable.append((name, f"{type(e).__name__} - {e}"))
    finally:
        apps.close()
    return finalize_index(entries, unreadable, "Excel or Word", title)


# ══════════════════════════════════════════════════════════
#  TOOL 1 - Convert to PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf(script_dir: str, title: str) -> Optional[str]:
    index = scan_office_index(script_dir, False, title)
    lines_of = lambda toks: file_input_lines(files_of(toks, index))
    state = run_inputs(title, [select_step(title, index, "file", "Select files to convert",
                                           "1, 3, 4-6", "Files", lines_of)],
                       lambda st: (lines_of(st["tokens"]), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "convert")
    tr = announce(title, script_dir, out_folder, file_input_lines(entries))
    apps = OfficeApps()

    def worker(entry: Entry) -> list:
        out_path = get_unique_path(out_folder, stem_of(entry.name) + ".pdf")
        export = excel_export if entry.kind == "excel" else word_export
        apps.run(entry.kind, lambda app: export(app, entry.path, out_path, None))
        return [os.path.basename(out_path)]

    try:
        t0 = begin(tr, "Converting")
        outputs, skipped, failures = run_items(entries, "converting to PDF", worker, lambda e: e.name)
    finally:
        apps.close()
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 2 - Merge PDF   (file-level; merge order = order selected)
# ══════════════════════════════════════════════════════════
def tool_merge_pdf(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, False, title)
    lines_of = lambda toks: file_input_lines(files_of(toks, index))
    state = run_inputs(title, [select_step(
        title, index, "file", "Select files to merge",
        "3, 1, 4-6        (merged in the order typed)", "Files", lines_of)],
        lambda st: (lines_of(st["tokens"]), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "merge")
    tr = announce(title, script_dir, out_folder, file_input_lines(entries))
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

    t0 = begin(tr, "Merging")
    _, skipped, failures = run_items(entries, "reading pages", worker, lambda e: e.name)
    outputs = []
    if merged:
        out_name = f"{stem_of(merged[0].name)}_merged.pdf"
        try:
            out_path = get_unique_path(out_folder, out_name)
            write_pdf(writer, out_path)
            outputs.append(os.path.basename(out_path))
        except Exception as e:
            failures.append(make_failure(out_name, "writing merged file", f"files merged = {len(merged)}", e))
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 3 - Split PDF   (file-level; one file per page)
# ══════════════════════════════════════════════════════════
def tool_split_pdf(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, False, title)
    lines_of = lambda toks: file_input_lines(files_of(toks, index))
    state = run_inputs(title, [select_step(title, index, "file", "Select files to split",
                                           "1, 3, 4-6", "Files", lines_of)],
                       lambda st: (lines_of(st["tokens"]), []))
    entries = files_of(state["tokens"], index)

    out_folder = out_dir(script_dir, "split")
    tr = announce(title, script_dir, out_folder, file_input_lines(entries))

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

    t0 = begin(tr, "Splitting")
    outputs, skipped, failures = run_items(entries, "reading and splitting pages", worker, lambda e: e.name)
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 4 - Rotate PDF   (page-level)
# ══════════════════════════════════════════════════════════
def prompt_rotation(title: str, choices: list) -> tuple:
    """Rotation screen -> (clockwise degrees for pypdf, label such as 'R90'). Left N == right (360 - N)."""
    def body() -> None:
        print("Rotation by: [L / R] [90 / 180]")
        print()
        print("e.g. R90")

    def handle(raw: str) -> tuple:
        check_nav(raw, False)
        if not raw:
            raise InvalidInput("")
        m = re.fullmatch(r"([LR])\s*(90|180)", raw.upper())
        if not m:
            raise InvalidInput(f'"{raw}" does not match the pattern R90')
        direction, degrees = m.group(1), int(m.group(2))
        return (degrees if direction == "R" else 360 - degrees), f"{direction}{degrees}"

    return prompt_screen(title, body, ["[L/R][90/180] Select"] + nav_keys(False), "Rotation: ",
                         handle, choices)


def tool_rotate_pdf(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, True, title)
    lines_of = lambda toks: unit_input_lines(group_by_entry(dedupe(flatten_codes(toks)), index))

    def rotation_step(st: dict, first: bool, choices: list) -> None:
        st["angle"], st["label"] = prompt_rotation(title, choices)

    def describe(st: dict) -> tuple:
        st["groups"] = group_by_entry(dedupe(flatten_codes(st["tokens"])), index)
        return unit_input_lines(st["groups"]), ["", f"Rotation : {st['label']}"]

    state = run_inputs(title, [
        select_step(title, index, "unit", "Select pages to rotate", "a1, b2, c4-6", "Pages", lines_of),
        (rotation_step, lambda st: [f"Rotation : {st['label']}"]),
    ], describe)
    angle, label, groups = state["angle"], state["label"], state["groups"]

    out_folder = out_dir(script_dir, "rotate")
    tr = announce(title, script_dir, out_folder, unit_input_lines(groups))
    tr.say()
    tr.say(f"Rotation : {label}")

    def worker(item: tuple) -> list:
        entry, idxs = item
        reader = PdfReader(entry.path)
        writer = PdfWriter()
        for i, page in enumerate(reader.pages):
            if i in idxs:
                page.rotate(angle)
            writer.add_page(page)
        out_path = get_unique_path(out_folder, f"{stem_of(entry.name)}_rotated.pdf")
        write_pdf(writer, out_path)
        return [os.path.basename(out_path)]

    t0 = begin(tr, "Rotating")
    outputs, skipped, failures = run_items(
        groups, "rotating pages", worker, lambda it: it[0].name,
        lambda it: f"pages selected = {[i + 1 for i in it[1]]}; rotation = {label}")
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 5 - Protect PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_protect_pdf(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, False, title)
    lines_of = lambda toks: file_input_lines(files_of(toks, index))

    def password_step(st: dict, first: bool, choices: list) -> None:
        st["password"] = prompt_password(title, "Password to apply", True, choices)

    state = run_inputs(title, [
        select_step(title, index, "file", "Select files to protect", "1, 3, 4-6", "Files", lines_of),
        (password_step, lambda st: [f"Password : {st['password']}"]),
    ], lambda st: (lines_of(st["tokens"]), ["", f"Password : {st['password']}"]))
    entries, password = files_of(state["tokens"], index), state["password"]

    out_folder = out_dir(script_dir, "protect")
    tr = announce(title, script_dir, out_folder, file_input_lines(entries))

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

    t0 = begin(tr, "Protecting")
    outputs, skipped, failures = run_items(entries, "reading and encrypting", worker, lambda e: e.name)
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 6 - Unprotect PDF   (file-level)
# ══════════════════════════════════════════════════════════
def tool_unprotect_pdf(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, False, title)
    lines_of = lambda toks: file_input_lines(files_of(toks, index))

    def password_step(st: dict, first: bool, choices: list) -> None:
        st["password"] = prompt_password(title, "Password to unlock", False, choices)

    state = run_inputs(title, [
        select_step(title, index, "file", "Select files to unprotect", "1, 3, 4-6", "Files", lines_of),
        (password_step, lambda st: [f"Password : {st['password']}"]),
    ], lambda st: (lines_of(st["tokens"]), ["", f"Password : {st['password']}"]))
    entries, password = files_of(state["tokens"], index), state["password"]

    out_folder = out_dir(script_dir, "unprotect")
    tr = announce(title, script_dir, out_folder, file_input_lines(entries))

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

    t0 = begin(tr, "Unprotecting")
    outputs, skipped, failures = run_items(entries, "opening and decrypting", worker, lambda e: e.name)
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 7 - Convert to PDF Precise   (unit-level: Excel sheets, Word pages)
# ══════════════════════════════════════════════════════════
def tool_convert_to_pdf_precise(script_dir: str, title: str) -> Optional[str]:
    index = scan_office_index(script_dir, True, title)
    lines_of = lambda toks: unit_input_lines(group_by_entry(dedupe(flatten_codes(toks)), index))

    def describe(st: dict) -> tuple:
        st["groups"] = group_by_entry(dedupe(flatten_codes(st["tokens"])), index)
        return unit_input_lines(st["groups"]), []

    state = run_inputs(title, [select_step(title, index, "unit", "Select sheets or pages to convert",
                                           "a1, b, c2-3", "Sheets or pages", lines_of)], describe)
    groups = state["groups"]

    out_folder = out_dir(script_dir, "convert")
    tr = announce(title, script_dir, out_folder, unit_input_lines(groups))
    apps = OfficeApps()

    def worker(item: tuple) -> list:
        entry, idxs = item
        whole = len(idxs) == entry.count                      # everything selected: no filtering
        out_path = get_unique_path(out_folder, stem_of(entry.name) + ".pdf")
        if entry.kind == "excel":
            keep = None if whole else [entry.unit_names[i] for i in idxs]
            apps.run("excel", lambda app: excel_export(app, entry.path, out_path, keep))
        else:
            apps.run("word", lambda app: word_export(app, entry.path, out_path, None if whole else idxs))
        return [os.path.basename(out_path)]

    try:
        t0 = begin(tr, "Converting")
        outputs, skipped, failures = run_items(
            groups, "converting selection to PDF", worker, lambda it: it[0].name,
            lambda it: f"{it[0].unit_word}s selected = {[i + 1 for i in it[1]]} of {it[0].count}")
    finally:
        apps.close()
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 8 - Selective Combine   (unit-level; one output, order as selected)
# ══════════════════════════════════════════════════════════
def tool_selective_combine(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, True, title)
    lines_of = lambda toks: unit_input_lines(ordered_groups(flatten_codes(toks), index))

    def describe(st: dict) -> tuple:
        st["codes"] = flatten_codes(st["tokens"])             # order and repeats kept on purpose
        return (unit_input_lines(ordered_groups(st["codes"], index)),
                ["", f"Output pages : {len(st['codes'])}"])

    state = run_inputs(title, [select_step(
        title, index, "unit", "Select pages to combine",
        "a1, b, c4-6        (combined in the order typed)", "Pages", lines_of)], describe)
    codes = state["codes"]

    out_folder = out_dir(script_dir, "combine")
    tr = announce(title, script_dir, out_folder, unit_input_lines(ordered_groups(codes, index)))
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

    t0 = begin(tr, "Combining")
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
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


# ══════════════════════════════════════════════════════════
#  TOOL 9 - Selective Extract   (unit-level; each token = one output file)
# ══════════════════════════════════════════════════════════
def tool_selective_extract(script_dir: str, title: str) -> Optional[str]:
    from pypdf import PdfReader, PdfWriter
    index = scan_pdf_index(script_dir, True, title)
    lines_of = lambda toks: unit_input_lines(group_by_entry(flatten_codes(toks), index))

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
        extra = ["", f"Output files: {len(st['jobs'])}"] + numbered([f'"{name}"' for _, _, name in st["jobs"]])
        return unit_input_lines(groups), extra

    state = run_inputs(title, [select_step(
        title, index, "unit", "Select pages to extract",
        "a1, b2, c4-6        (each entry = one file)", "Pages", lines_of)], describe)
    jobs = state["jobs"]

    out_folder = out_dir(script_dir, "extract")
    tr = announce(title, script_dir, out_folder, lines_of(state["tokens"]))
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

    t0 = begin(tr, "Extracting")
    outputs, skipped, failures = run_items(
        jobs, "extracting pages", worker, lambda j: j[2],
        lambda j: f'source = "{j[0].name}"; pages = {[i + 1 for i in j[1]]}')
    return finish(title, script_dir, out_folder, t0, outputs, skipped, failures, tr)


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

    def row(left: str, right: str) -> str:
        return f"| {left:<{w - 1}}| {right:<{w - 1}}|"

    cell = lambda k: f"[{k}] {TOOLS[k][0]}"
    lefts = [cell(k) for k in "123456"]
    rights = [cell(k) for k in "789"]
    lines = [bar, row("BASIC", "ADVANCE"), bar]
    for i, left in enumerate(lefts):
        lines.append(row(left, rights[i] if i < len(rights) else ""))
    lines.append(bar)
    return chr(10).join(lines)


def choose_tool(notice: Optional[str]) -> str:
    """Menu screen. Returns '0' (exit) or a key of TOOLS. notice: left by the last tool (Saved: ...)."""
    def handle(raw: str) -> str:
        if raw == "0" or raw in TOOLS:
            return raw
        raise InvalidInput("")

    return prompt_screen(SUITE_NAME, lambda: print(build_menu()), ["[#] Select", "[0] Exit"],
                         "Ans: ", handle, notice=notice)


def run_tool(name: str, fn: Callable, script_dir: str) -> Optional[str]:
    """One tool run -> notice for the menu screen. UserCancelled ends quietly (back to the
    menu); ExitSession propagates to main()."""
    try:
        return fn(script_dir, name)
    except UserCancelled:
        return None
    except (ExitSession, ExitProgram):
        raise
    except NoInput as e:
        open_screen(name)
        tr = Transcript()
        tr.say(str(e))
        return tail(name, script_dir, script_dir, [], tr)
    except Exception as e:
        failures = [make_failure(name, "running tool", "", e)]
        open_screen(name)
        tr = Transcript()
        tr.say("FAILED: tool stopped unexpectedly")
        for line in failure_lines([(failures[0]["item"], failures[0]["error"])]):
            tr.say(line)
        return tail(name, script_dir, script_dir, failures, tr)


def main() -> None:
    script_dir = get_script_dir()
    os.chdir(script_dir)
    for stream in (sys.stdout, sys.stderr):      # a filename the console cannot encode must not crash a tool
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    notice = None
    while True:
        choice = choose_tool(notice)
        if choice == "0":
            return                                  # [0] Exit: close immediately, no screen
        name, fn = TOOLS[choice]
        try:
            notice = run_tool(name, fn, script_dir)
        except ExitProgram:
            return                                  # [0] Exit after a tool: close immediately
        except ExitSession:
            open_screen(name)
            print("Cancelled.")
            print()
            print("-" * LINE_WIDTH)
            input("[Enter] to close...")
            return


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FAILED: {type(e).__name__} - {e}")
        try:
            print()
            print("-" * LINE_WIDTH)
            input("[Enter] to close...")
        except Exception:
            pass
