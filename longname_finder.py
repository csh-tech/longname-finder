# -*- coding: utf-8 -*-
"""
LongName Finder
Find files with very long names (the kind that make Windows Explorer hang or
crash) and export them to an Excel report with clickable links.

Only the file name itself is checked (including the extension), not the folder path.
Windows allows at most 255 characters per file name.

Usage:
  longname_finder.py                       -> a folder picker opens
  longname_finder.py "D:\\Data"             -> scan a specific folder
  longname_finder.py "D:\\Data" --min 200   -> custom threshold
  longname_finder.py --lang en|he           -> force interface language

Interface text lives in locales/*.json (one file per language).

License: MIT
"""
import argparse
import locale
import os
import sys
from datetime import datetime
from pathlib import Path

__version__ = "1.0.0"
APP_NAME = "LongName Finder"

IS_WINDOWS = os.name == "nt"
EXCEL_MAX_ROWS = 1_048_575
NAME_MAX = 255          # Windows (NTFS) limit for a single file name
DEFAULT_MIN = 240       # default threshold: names close to the limit

# --------------------------------------------------------------------------
# Translations - loaded from locales/<code>.json
# A "locales" folder next to the exe/script overrides the built-in one,
# so new languages can be added without rebuilding.
# --------------------------------------------------------------------------
import json


def _locale_dirs():
    base = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__)))
    dirs = [os.path.join(base, "locales")]
    if hasattr(sys, "_MEIPASS"):                       # inside the PyInstaller exe
        dirs.append(os.path.join(sys._MEIPASS, "locales"))
    return dirs


def load_locales() -> dict:
    found = {}
    for d in reversed(_locale_dirs()):                 # later dirs override earlier
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.endswith(".json"):
                try:
                    with open(os.path.join(d, fn), encoding="utf-8") as f:
                        found[fn[:-5]] = json.load(f)
                except (OSError, ValueError) as e:
                    print(f"Could not load language file {fn}: {e}")
    if "en" not in found:
        print("Language files not found (locales/en.json is required).")
        sys.exit(1)
    return found


STRINGS = load_locales()


def detect_language() -> str:
    """Pick the interface language from the operating system (falls back to English)."""
    codes = []
    if IS_WINDOWS:
        try:
            import ctypes
            langid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            codes.append(locale.windows_locale.get(langid, ""))
        except Exception:
            pass
    codes += [os.environ.get("LANG", ""), locale.getlocale()[0] or ""]
    for code in codes:
        short = code[:2].lower()
        if short == "iw":
            short = "he"
        if short in STRINGS:
            return short
    return "en"


def get_lang_from_argv() -> str:
    """Read --lang early, so even argparse help text is translated."""
    for i, arg in enumerate(sys.argv):
        if arg.startswith("--lang="):
            val = arg.split("=", 1)[1]
        elif arg == "--lang" and i + 1 < len(sys.argv):
            val = sys.argv[i + 1]
        else:
            continue
        if val in STRINGS:
            return val
    return detect_language()


LANG = get_lang_from_argv()
IS_RTL = bool(STRINGS[LANG].get("_rtl", False))
_FMT = {"max": NAME_MAX, "default": DEFAULT_MIN}


def t(key: str, **kwargs):
    value = STRINGS[LANG].get(key, STRINGS["en"].get(key, key))
    args = {**_FMT, **kwargs}
    if isinstance(value, list):
        return [v.format(**args) for v in value]
    return value.format(**args)


try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print(t("missing_openpyxl"))
    sys.exit(1)


# --------------------------------------------------------------------------
# File system helpers
# --------------------------------------------------------------------------
def name_len(name: str) -> int:
    """Length as Windows counts it (UTF-16 code units), incl. extension."""
    return len(name.encode("utf-16-le")) // 2


def to_long(path: str) -> str:
    """Add the \\\\?\\ prefix on Windows so paths over 260 characters can be accessed."""
    if not IS_WINDOWS:
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):  # network path \\server\share
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def from_long(path: str) -> str:
    """Remove the \\\\?\\ prefix for display."""
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def walk_files(root: str, errors: list):
    """Walk all files (long paths included) without stopping on unreadable folders."""
    stack = [to_long(root)]
    while stack:
        current = stack.pop()
        try:
            with os.scandir(current) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            yield from_long(entry.path)
                    except OSError as e:
                        errors.append((from_long(entry.path), str(e)))
        except OSError as e:
            errors.append((from_long(current), str(e)))


def file_url(path: str) -> str:
    return Path(path).as_uri() if os.path.isabs(path) else path


def pick_folder() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        tk_root = tk.Tk()
        tk_root.withdraw()
        folder = filedialog.askdirectory(title=t("pick_title"))
        tk_root.destroy()
        return folder
    except Exception:
        return input(t("pick_prompt")).strip().strip('"')


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------
def ask_save_path(default_name: str):
    """Ask where to save the report. Returns None if the dialog can't be shown or is cancelled."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        tk_root = tk.Tk()
        tk_root.withdraw()
        path = filedialog.asksaveasfilename(
            title=t("save_title"), initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        tk_root.destroy()
        return path or None
    except Exception:
        return None


def build_report(root: str, min_len: int, out_path: str, ask_where: bool = False):
    errors, rows = [], []
    print(t("scanning", root=root))
    for i, full in enumerate(walk_files(root, errors), 1):
        if i % 5000 == 0:
            print(t("progress", scanned=i, found=len(rows)))
        folder, name = os.path.split(full)
        n = name_len(name)
        if n >= min_len:
            rows.append((name, folder, full, n, len(full)))
            if len(rows) >= EXCEL_MAX_ROWS:
                print(t("excel_limit"))
                break

    rows.sort(key=lambda r: r[3], reverse=True)  # longest names first
    rtl = IS_RTL

    wb = Workbook()
    ws = wb.active
    ws.title = t("sheet_files")
    ws.sheet_view.rightToLeft = rtl
    ws.append(t("headers"))
    head_fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    link_font = Font(color="0563C1", underline="single")
    warn_fill = PatternFill("solid", fgColor="FCE4D6")
    for r, (name, folder, full, nlen, flen) in enumerate(rows, start=2):
        at_limit = nlen >= NAME_MAX
        ws.cell(r, 1, name)
        ws.cell(r, 2, folder)
        c = ws.cell(r, 3, t("open_file"))
        c.hyperlink = file_url(full)
        c.font = link_font
        c = ws.cell(r, 4, t("open_folder"))
        c.hyperlink = file_url(folder)
        c.font = link_font
        ws.cell(r, 5, nlen)
        ws.cell(r, 6, flen)
        ws.cell(r, 7, t("yes") if at_limit else t("no"))
        if at_limit:
            for col in (1, 2, 5, 6, 7):
                ws.cell(r, col).fill = warn_fill

    for col, width in zip(range(1, 8), [60, 70, 12, 12, 14, 14, 12]):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:G{len(rows) + 1}"

    if errors:
        es = wb.create_sheet(t("sheet_errors"))
        es.sheet_view.rightToLeft = rtl
        es.append(t("err_headers"))
        for c in es[1]:
            c.font = Font(bold=True)
        for e in errors:
            es.append(list(e))
        es.column_dimensions["A"].width = 90
        es.column_dimensions["B"].width = 60

    if ask_where:
        chosen = ask_save_path(os.path.basename(out_path))
        if chosen:
            out_path = chosen
        else:
            print(t("cancelled_save"))
    wb.save(out_path)
    at_limit = sum(1 for r in rows if r[3] >= NAME_MAX)
    print()
    print(t("done", count=len(rows), min=min_len, at_limit=at_limit, max=NAME_MAX))
    if errors:
        print(t("errors_note", count=len(errors)))
    print(t("saved", path=os.path.abspath(out_path)))
    return os.path.abspath(out_path)


def main():
    p = argparse.ArgumentParser(prog="LongNameFinder", description=t("arg_desc"))
    p.add_argument("folder", nargs="?", help=t("arg_folder"))
    p.add_argument("--min", type=int, default=DEFAULT_MIN, help=t("arg_min"))
    p.add_argument("-o", "--output", help=t("arg_out"))
    p.add_argument("--lang", choices=sorted(STRINGS), help=t("arg_lang"))
    p.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    a = p.parse_args()

    interactive = not a.folder
    folder = a.folder or pick_folder()
    if not folder or not os.path.isdir(to_long(folder)):
        print(t("no_folder"))
        if interactive and IS_WINDOWS:
            input(t("press_enter"))
        sys.exit(1)
    folder = os.path.abspath(folder)
    out = a.output or f"long_names_{datetime.now():%Y%m%d_%H%M}.xlsx"
    saved = build_report(folder, a.min, out, ask_where=interactive and not a.output)

    if interactive and IS_WINDOWS:
        try:
            os.startfile(saved)                  # open the report in Excel
        except OSError:
            pass
        input("\n" + t("press_enter"))


if __name__ == "__main__":
    main()
