"""Faithful Excel (.xlsx / .xlsm) to Markdown converter for MarkItDown.

MarkItDown's built-in converter reads workbooks through pandas, which treats
the first row as the header, prints numbers in its own float format, and drops
everything that isn't a cell value. This converter streams the workbook XML
directly (fast and memory-friendly even for 200 MB+ workbooks) and keeps:

* the real header row: taken from the autofilter range, pivot table layout,
  freeze panes or a manual override, with a label row above it (e.g. week
  names over dates) merged in;
* numbers as Excel displays them, using each cell's number format
  (no scientific notation, no lost digits);
* comments, autofilter criteria, merged cells, data validation, hidden
  rows / columns / sheets, and pivot table layout and source.

Formula results are the values Excel cached when the file was last saved.
"""

from __future__ import annotations

import math
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, BinaryIO, Dict, List, Optional, Tuple

from lxml import etree
from markitdown import DocumentConverter, DocumentConverterResult, StreamInfo

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = "{%s}" % M
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
TC = "{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}"
X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
XM = "{http://schemas.microsoft.com/office/excel/2006/main}"

ACCEPTED_EXTENSIONS = (".xlsx", ".xlsm")
ACCEPTED_MIME_PREFIXES = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml",
    "application/vnd.ms-excel.sheet.macroenabled",
)

# Built-in number formats (ECMA-376 18.8.30), used when styles.xml doesn't define the id.
BUILTIN_FORMATS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
    11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm",
    17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}
DATE_FORMAT_IDS = set(range(14, 23)) | {45, 46, 47}


@dataclass
class XlsxOptions:
    """Conversion options; also accepted as MarkItDown convert() keyword arguments."""

    header_rows: Dict[str, int] = field(default_factory=dict)  # sheet name -> header row (0 = none)
    visible_only: bool = False  # skip hidden sheets, rows and columns
    raw_values: bool = False  # full-precision numbers instead of Excel's display format
    formulas: str = "summary"  # "none", "summary" (per-sheet pattern list) or "cells" (also under each value)
    recalc_note: str = ""  # set when Excel recalculated the workbook before conversion

    @classmethod
    def from_kwargs(cls, base: "XlsxOptions", kwargs: Dict[str, Any]) -> "XlsxOptions":
        formulas = kwargs.get("xlsx_formulas") or base.formulas
        if formulas not in FORMULA_MODES:
            raise ValueError(f"xlsx_formulas must be one of {', '.join(FORMULA_MODES)} (got {formulas!r})")
        return cls(
            header_rows=dict(kwargs.get("xlsx_header_rows") or base.header_rows),
            visible_only=bool(kwargs.get("xlsx_visible_only", base.visible_only)),
            raw_values=bool(kwargs.get("xlsx_raw_values", base.raw_values)),
            formulas=formulas,
            recalc_note=kwargs.get("xlsx_recalc_note") or base.recalc_note,
        )


FORMULA_MODES = ("none", "summary", "cells")


# --------------------------------------------------------------------------- #
# Cell references
# --------------------------------------------------------------------------- #
_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def parse_range(ref: str) -> Optional[Tuple[int, int, int, int]]:
    """'A14:DZ6758' -> (min_row, min_col, max_row, max_col)."""
    parts = ref.split(":")
    m1 = _REF.fullmatch(parts[0].strip())
    m2 = _REF.fullmatch(parts[-1].strip())
    if not m1 or not m2:
        return None
    r1, c1 = int(m1.group(2)), col_index(m1.group(1))
    r2, c2 = int(m2.group(2)), col_index(m2.group(1))
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def compress_numbers(numbers: List[int], limit: int = 12) -> str:
    """[3,4,5,9] -> '3–5, 9' (at most `limit` ranges)."""
    ranges: List[Tuple[int, int]] = []
    for n in sorted(numbers):
        if ranges and n == ranges[-1][1] + 1:
            ranges[-1] = (ranges[-1][0], n)
        else:
            ranges.append((n, n))
    text = ", ".join(f"{a}" if a == b else f"{a}–{b}" for a, b in ranges[:limit])
    if len(ranges) > limit:
        text += f" …（共 {len(ranges)} 段）"
    return text


def compress_columns(cols: List[int]) -> str:
    ranges: List[Tuple[int, int]] = []
    for n in sorted(cols):
        if ranges and n == ranges[-1][1] + 1:
            ranges[-1] = (ranges[-1][0], n)
        else:
            ranges.append((n, n))
    return ", ".join(col_letter(a) if a == b else f"{col_letter(a)}:{col_letter(b)}" for a, b in ranges)


# --------------------------------------------------------------------------- #
# Number formatting (Excel display)
# --------------------------------------------------------------------------- #
def _split_sections(code: str) -> List[str]:
    sections, current, quoted, escaped = [], "", False, False
    for ch in code:
        if escaped:
            current += ch
            escaped = False
        elif ch == "\\":
            current += ch
            escaped = True
        elif ch == '"':
            quoted = not quoted
            current += ch
        elif ch == ";" and not quoted:
            sections.append(current)
            current = ""
        else:
            current += ch
    sections.append(current)
    return sections


def _strip_literals(section: str) -> str:
    """Remove quoted text, escapes, brackets, padding and fill so only codes remain."""
    s = re.sub(r'"[^"]*"', "", section)
    s = re.sub(r"\\.", "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    s = re.sub(r"[_*].", "", s)
    return s


def is_date_code(code: str) -> bool:
    section = _strip_literals(_split_sections(code)[0]).lower()
    if "general" in section:
        return False
    return bool(re.search(r"[dmyhs]", section)) and not re.search(r"[0#?]", section)


def general_text(value: float, significant: int = 10) -> str:
    """Excel 'General'-like text without scientific notation or float noise."""
    if value == 0:
        return "0"
    if float(value).is_integer() and abs(value) < 1e21:
        return str(int(value))
    exponent = math.floor(math.log10(abs(value)))
    places = max(0, significant - 1 - exponent)
    d = Decimal(repr(float(value))).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    text = format(d, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def _date_text(value: float, code: str, date1904: bool) -> str:
    epoch = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
    if not date1904 and value < 60:  # Excel's fictitious 1900-02-29
        value += 1
    try:
        moment = epoch + timedelta(days=float(value))
    except OverflowError:
        return general_text(value)
    s = _strip_literals(_split_sections(code)[0]).lower()
    s_no_ampm = s.replace("am/pm", "").replace("a/p", "")
    has_time = "h" in s_no_ampm or "s" in s_no_ampm
    has_day = "d" in s_no_ampm
    has_year = "y" in s_no_ampm
    has_month = "m" in s_no_ampm.replace("h:mm", "").replace("mm:ss", "")
    if has_time and not (has_day or has_year):
        return moment.strftime("%H:%M:%S" if "s" in s_no_ampm else "%H:%M")
    if has_time:
        return moment.strftime("%Y-%m-%d %H:%M:%S" if "s" in s_no_ampm else "%Y-%m-%d %H:%M")
    if has_year and has_month and not has_day:
        return moment.strftime("%Y-%m")
    return moment.strftime("%Y-%m-%d")


def _numeric_section(section: str, value: float) -> Optional[str]:
    """Render a non-negative value with one format section. None = unsupported."""
    if re.search(r"[eE][+-]|/", _strip_literals(section)):
        return None  # scientific / fraction formats: fall back to General
    prefix, suffix, pattern = "", "", ""
    started = False
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            end = section.find('"', i + 1)
            end = len(section) if end < 0 else end
            literal = section[i + 1:end]
            i = end + 1
            if started:
                suffix += literal
            else:
                prefix += literal
            continue
        if ch == "\\" and i + 1 < len(section):
            if started:
                suffix += section[i + 1]
            else:
                prefix += section[i + 1]
            i += 2
            continue
        if ch == "[":
            end = section.find("]", i)
            end = len(section) if end < 0 else end
            token = section[i + 1:end]
            currency = re.match(r"\$([^-\]]*)", token)
            if currency and currency.group(1):
                if started:
                    suffix += currency.group(1)
                else:
                    prefix += currency.group(1)
            i = end + 1
            continue
        if ch in "_*":
            i += 2
            continue
        if ch in "0#?.,":  # literal text between placeholders ends up in the suffix
            started = True
            pattern += ch
            i += 1
            continue
        if ch == "%":
            pattern += "%"
            started = True
            i += 1
            continue
        if ch == "@":
            i += 1
            continue
        if started:
            suffix += ch
        else:
            prefix += ch
        i += 1

    percent = pattern.count("%")
    digits = pattern.replace("%", "")
    if not re.search(r"[0#?]", digits):
        return (prefix + suffix).strip() or None
    int_part, _, frac_part = digits.partition(".")
    scale = len(int_part) - len(int_part.rstrip(","))
    int_core = int_part.rstrip(",")
    grouping = "," in int_core
    min_int = int_core.count("0")
    min_dec = frac_part.count("0")
    max_dec = len(re.sub(r"[^0#?]", "", frac_part))

    v = Decimal(repr(float(value))) * (Decimal(100) ** percent) / (Decimal(1000) ** scale)
    v = v.quantize(Decimal(1).scaleb(-max_dec), rounding=ROUND_HALF_UP)
    text = format(v, "f")
    whole, _, frac = text.partition(".")
    frac = frac.rstrip("0")
    if len(frac) < min_dec:
        frac = frac + "0" * (min_dec - len(frac))
    whole_digits = whole.lstrip("-")
    if whole_digits == "0" and min_int == 0:
        whole_digits = "" if frac else ("0" if not prefix.strip() and not suffix.strip() else "")
    whole_digits = whole_digits.rjust(min_int, "0") if whole_digits or min_int else whole_digits
    if grouping and whole_digits:
        whole_digits = f"{int(whole_digits):,}"
    body = whole_digits + ("." + frac if frac else "")
    if not body and not (prefix.strip() or suffix.strip()):
        body = "0"
    return (prefix + body + "%" * percent + suffix).strip()


def format_number(value: float, code: Optional[str], date1904: bool = False, raw: bool = False) -> str:
    if raw:
        return general_text(value, significant=15)
    if not code or code.strip().lower() == "general":
        return general_text(value)
    sections = _split_sections(code)
    if any(re.match(r"\s*\[[<>=]", s) for s in sections):
        return general_text(value)  # conditional sections: rare, keep it exact instead
    if value < 0:
        if len(sections) >= 2 and sections[1].strip():
            section, v, sign = sections[1], -value, ""
        else:
            section, v, sign = sections[0], -value, "-"
    elif value == 0 and len(sections) >= 3 and sections[2].strip():
        section, v, sign = sections[2], 0.0, ""
    else:
        section, v, sign = sections[0], value, ""
    if section.strip().lower() == "general":
        return sign + general_text(v)
    if is_date_code(section):
        return _date_text(value, section, date1904)
    rendered = _numeric_section(section, v)
    if rendered is None:
        return general_text(value)
    return sign + rendered if sign and rendered else rendered


# --------------------------------------------------------------------------- #
# Formulas
# --------------------------------------------------------------------------- #
# Parts of a formula that must never be rewritten: "string literals",
# 'quoted sheet names' and [structured / external references].
_FORMULA_LITERAL = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|\[[^\]]*\]')
_WORD = r"A-Za-z0-9_.\u00C0-\uFFFF"
_A1_TOKEN = re.compile(
    rf"(?<![{_WORD}$])(?:"
    r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})"  # A1, $A$1
    r"|(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})"  # whole columns  A:B
    r"|(\$?)(\d{1,7}):(\$?)(\d{1,7})"  # whole rows     1:2
    rf")(?![{_WORD}(])"
)
_SHEET_REF = re.compile(rf"'((?:[^']|'')+)'!|(?<![{_WORD}'\]])([A-Za-z_\u00C0-\uFFFF][{_WORD}]*)!")
_FUTURE_PREFIX = re.compile(r"_xl(?:fn|ws|pm)\.")
MAX_ROW, MAX_COL = 1_048_576, 16_384


def _outside_literals(formula: str, transform) -> str:
    out, pos = [], 0
    for m in _FORMULA_LITERAL.finditer(formula):
        out.append(transform(formula[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(transform(formula[pos:]))
    return "".join(out)


def shift_formula(formula: str, d_row: int, d_col: int) -> str:
    """Move a formula by (d_row, d_col) the way Excel fills shared formulas:
    relative references shift, $-anchored ones stay, off-sheet ones become #REF!."""

    def shift(m: "re.Match") -> str:
        g = m.groups()
        if g[1] is not None:
            col = col_index(g[1]) + (0 if g[0] else d_col)
            row = int(g[3]) + (0 if g[2] else d_row)
            if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
                return "#REF!"
            return f"{g[0]}{col_letter(col)}{g[2]}{row}"
        if g[5] is not None:
            a = col_index(g[5]) + (0 if g[4] else d_col)
            b = col_index(g[7]) + (0 if g[6] else d_col)
            if not (1 <= a <= MAX_COL and 1 <= b <= MAX_COL):
                return "#REF!"
            return f"{g[4]}{col_letter(a)}:{g[6]}{col_letter(b)}"
        a = int(g[9]) + (0 if g[8] else d_row)
        b = int(g[11]) + (0 if g[10] else d_row)
        if not (1 <= a <= MAX_ROW and 1 <= b <= MAX_ROW):
            return "#REF!"
        return f"{g[8]}{a}:{g[10]}{b}"

    if d_row == 0 and d_col == 0:
        return formula
    return _outside_literals(formula, lambda code: _A1_TOKEN.sub(shift, code))


def formula_pattern(formula: str, row: int, col: int) -> str:
    """R1C1-style key: identical for every copy of a formula filled down or across."""

    def rel(value: int, base: int, absolute: str, axis: str) -> str:
        return f"{axis}{value}" if absolute else f"{axis}[{value - base}]"

    def convert(m: "re.Match") -> str:
        g = m.groups()
        if g[1] is not None:
            return rel(int(g[3]), row, g[2], "R") + rel(col_index(g[1]), col, g[0], "C")
        if g[5] is not None:
            return rel(col_index(g[5]), col, g[4], "C") + ":" + rel(col_index(g[7]), col, g[6], "C")
        return rel(int(g[9]), row, g[8], "R") + ":" + rel(int(g[11]), row, g[10], "R")

    return _outside_literals(formula, lambda code: _A1_TOKEN.sub(convert, code))


def display_formula(formula: str, array: bool = False) -> str:
    """'=XLOOKUP(...)' as Excel shows it (without the _xlfn. storage prefixes)."""
    text = "=" + _FUTURE_PREFIX.sub("", formula)
    return "{" + text + "}" if array else text


def referenced_sheets(formula: str) -> set:
    names = set()
    for part in re.split(r'"(?:[^"]|"")*"', formula):  # ignore string literals
        for m in _SHEET_REF.finditer(part):
            names.add(m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2))
    return names


@dataclass
class FormulaPattern:
    first_ref: Tuple[int, int]  # (row, col) of the first cell using it
    example: str  # that cell's formula, as displayed
    count: int = 0
    cells: Dict[int, List[int]] = field(default_factory=dict)  # col -> rows


def _row_runs(rows: List[int]) -> List[Tuple[int, int]]:
    runs, start, prev = [], rows[0], rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
            continue
        runs.append((start, prev))
        start = prev = r
    runs.append((start, prev))
    return runs


def describe_cells(cells: Dict[int, List[int]], limit: int = 4) -> str:
    """Compact description of where a formula pattern is used.

    Columns with identical row runs merge into rectangles; rectangles that repeat
    at a fixed row interval (e.g. one row in every 6-row block) are summarised."""
    runs = [(r1, r2, c) for c, rows in cells.items() for r1, r2 in _row_runs(sorted(rows))]
    runs.sort()
    rects: List[List[int]] = []
    for r1, r2, c in runs:
        last = rects[-1] if rects else None
        if last and last[0] == r1 and last[1] == r2 and last[3] == c - 1:
            last[3] = c
        else:
            rects.append([r1, r2, c, c])

    def ref(r1: int, r2: int, c1: int, c2: int) -> str:
        start = f"{col_letter(c1)}{r1}"
        return start if (r1, c1) == (r2, c2) else f"{start}:{col_letter(c2)}{r2}"

    # Group rectangles of the same shape and columns; spot a fixed repeat interval.
    groups: Dict[Tuple[int, int, int], List[List[int]]] = {}
    for rect in rects:
        groups.setdefault((rect[2], rect[3], rect[1] - rect[0]), []).append(rect)
    parts: List[Tuple[Tuple[int, int], str]] = []
    for (c1, c2, height), members in groups.items():
        starts = [m[0] for m in members]
        steps = {b - a for a, b in zip(starts, starts[1:])}
        if len(members) >= 3 and len(steps) == 1:
            step = steps.pop()
            first, last = members[0], members[-1]
            parts.append(((c1, first[0]), f"{ref(*first[:2], c1, c2)} 起每 {step} 列一次，共 {len(members):,} 次（到第 {last[1]:,} 列）"))
        else:
            parts.extend(((c1, m[0]), ref(m[0], m[1], c1, c2)) for m in members)
    parts.sort()
    text = "、".join(p for _, p in parts[:limit])
    return text + (f" 等 {len(parts):,} 處" if len(parts) > limit else "")


# --------------------------------------------------------------------------- #
# Workbook model
# --------------------------------------------------------------------------- #
@dataclass
class Cell:
    text: str
    kind: str  # "s" text, "n" number, "d" date, "b" bool, "e" error, "f" formula without a cached value
    formula: Optional[str] = None  # displayed formula, kept only with formulas="cells"


@dataclass
class Pivot:
    name: str
    ref: Tuple[int, int, int, int]
    header_row: int
    page_rows: Tuple[int, int] = (0, 0)
    rows: List[str] = field(default_factory=list)
    columns: List[str] = field(default_factory=list)
    values: List[str] = field(default_factory=list)
    filters: List[str] = field(default_factory=list)
    source: str = ""
    refreshed: str = ""


@dataclass
class Sheet:
    name: str
    state: str
    path: str
    cells: Dict[int, Dict[int, Cell]] = field(default_factory=dict)
    hidden_rows: set = field(default_factory=set)
    hidden_cols: set = field(default_factory=set)
    merged: List[str] = field(default_factory=list)
    autofilter: Optional[str] = None
    filter_columns: List[str] = field(default_factory=list)
    freeze: Tuple[int, int] = (0, 0)  # frozen rows, frozen columns
    validations: List[str] = field(default_factory=list)
    comments: Dict[str, List[Tuple[str, str]]] = field(default_factory=dict)
    pivots: List[Pivot] = field(default_factory=list)
    tables: List[Tuple[str, Tuple[int, int, int, int], int]] = field(default_factory=list)
    dimension: str = ""
    formula_count: int = 0
    formula_patterns: Dict[str, FormulaPattern] = field(default_factory=dict)
    shared_formulas: Dict[str, Tuple[str, int, int, str]] = field(default_factory=dict)  # si -> master
    references: Dict[str, int] = field(default_factory=dict)  # other sheet -> formula cells reading it


class WorkbookReader:
    def __init__(self, archive: zipfile.ZipFile, options: XlsxOptions):
        self.zip = archive
        self.names = set(archive.namelist())
        self.options = options
        self.shared: List[str] = []
        self.formats: List[Optional[str]] = []
        self.date1904 = False
        self.persons: Dict[str, str] = {}

    # ----- package helpers ------------------------------------------------- #
    def _rels(self, part: str) -> Dict[str, Tuple[str, str]]:
        rels_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rels_path not in self.names:
            return {}
        out = {}
        for rel in etree.parse(self.zip.open(rels_path)).getroot():
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                continue
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target)
            )
            out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], resolved)
        return out

    def _xml(self, part: str):
        return etree.parse(self.zip.open(part)).getroot()

    # ----- workbook-level parts --------------------------------------------- #
    def read(self) -> List[Sheet]:
        wb_part = "xl/workbook.xml"
        if wb_part not in self.names:
            raise ValueError("Not an Excel workbook (xl/workbook.xml missing)")
        rels = self._rels(wb_part)
        workbook = self._xml(wb_part)
        pr = workbook.find(NS + "workbookPr")
        self.date1904 = pr is not None and pr.get("date1904") in ("1", "true")
        for kind, target in rels.values():
            if kind == "sharedStrings":
                self._read_shared_strings(target)
            elif kind == "styles":
                self._read_styles(target)
        if "xl/persons/person.xml" in self.names:
            for person in self._xml("xl/persons/person.xml"):
                self.persons[person.get("id")] = person.get("displayName", "")
        sheets = []
        for el in workbook.iter(NS + "sheet"):
            kind, target = rels.get(el.get(R_ID), ("", ""))
            if kind != "worksheet" or target not in self.names:
                continue  # chart sheets, dialog sheets
            sheets.append(Sheet(el.get("name", ""), el.get("state", "visible"), target))
        return sheets

    def _read_shared_strings(self, part: str) -> None:
        for _, si in etree.iterparse(self.zip.open(part), tag=NS + "si"):
            texts = [t.text or "" for t in si.iter(NS + "t") if t.getparent().tag != NS + "rPh"]
            self.shared.append("".join(texts))
            si.clear()

    def _read_styles(self, part: str) -> None:
        root = self._xml(part)
        custom = {int(n.get("numFmtId")): n.get("formatCode", "") for n in root.iter(NS + "numFmt")}
        xfs = root.find(NS + "cellXfs")
        for xf in (xfs if xfs is not None else []):
            fmt_id = int(xf.get("numFmtId", "0"))
            code = custom.get(fmt_id) or BUILTIN_FORMATS.get(fmt_id)
            if code is None and fmt_id in DATE_FORMAT_IDS:
                code = "yyyy-mm-dd"
            self.formats.append(code)

    # ----- worksheet -------------------------------------------------------- #
    def load_sheet(self, sheet: Sheet) -> None:
        raw = self.options.raw_values
        for event, el in etree.iterparse(self.zip.open(sheet.path), events=("end",)):
            tag = el.tag[len(NS):] if el.tag.startswith(NS) else el.tag
            if tag == "row":
                self._read_row(sheet, el, raw)
                el.clear()
                while el.getprevious() is not None:
                    del el.getparent()[0]
            elif tag == "dimension":
                sheet.dimension = el.get("ref", "")
            elif tag == "pane" and el.get("state") in ("frozen", "frozenSplit"):
                sheet.freeze = (int(float(el.get("ySplit", "0"))), int(float(el.get("xSplit", "0"))))
            elif tag == "col" and el.get("hidden") in ("1", "true"):
                sheet.hidden_cols.update(range(int(el.get("min")), int(el.get("max")) + 1))
            elif tag == "mergeCell":
                sheet.merged.append(el.get("ref", ""))
            elif tag == "autoFilter" and el.getparent() is not None and el.getparent().tag == NS + "worksheet":
                sheet.autofilter = el.get("ref")
                sheet.filter_columns = self._describe_filter(el)
            elif tag in ("dataValidation", X14 + "dataValidation"):
                # Lists sourced from another sheet are stored in the x14 extension.
                sheet.validations.append(self._describe_validation(el))
        self._read_sheet_parts(sheet)
        for pattern in sheet.formula_patterns.values():
            for name in referenced_sheets(pattern.example):
                if name != sheet.name:
                    sheet.references[name] = sheet.references.get(name, 0) + pattern.count
        sheet.shared_formulas.clear()

    def _read_row(self, sheet: Sheet, row, raw: bool) -> None:
        r = int(row.get("r", "0")) or (max(sheet.cells) + 1 if sheet.cells else 1)
        if row.get("hidden") in ("1", "true"):
            sheet.hidden_rows.add(r)
        cells: Dict[int, Cell] = {}
        next_col = 1
        for c in row.iter(NS + "c"):
            ref = c.get("r")
            col = col_index(_REF.fullmatch(ref).group(1)) if ref else next_col
            next_col = col + 1
            cell = self._cell_value(c, raw)
            f = c.find(NS + "f")
            if f is not None and self.options.formulas != "none":
                formula = self._read_formula(sheet, f, r, col)
                if formula and self.options.formulas == "cells":
                    if cell is None:
                        cell = Cell("", "f")
                    cell.formula = formula
            if cell is not None:
                cells[col] = cell
        if cells:
            sheet.cells[r] = cells

    def _read_formula(self, sheet: Sheet, f, r: int, col: int) -> Optional[str]:
        """Record the cell's formula pattern; return the displayed formula when per-cell output needs it."""
        kind = f.get("t", "normal")
        text = f.text or ""
        if kind == "dataTable":
            return None
        if kind == "shared":
            si = f.get("si", "")
            if text:  # the master cell carries the formula for the whole block
                key = formula_pattern(text, r, col)
                sheet.shared_formulas[si] = (text, r, col, key)
            else:
                master = sheet.shared_formulas.get(si)
                if master is None:
                    return None
                m_text, m_row, m_col, key = master
                text = shift_formula(m_text, r - m_row, col - m_col) if self.options.formulas == "cells" else ""
        elif text:
            key = formula_pattern(text, r, col)
        else:
            return None
        sheet.formula_count += 1
        pattern = sheet.formula_patterns.get(key)
        displayed = display_formula(text, kind == "array") if text else None
        if pattern is None:
            pattern = FormulaPattern((r, col), displayed or "")
            sheet.formula_patterns[key] = pattern
        pattern.count += 1
        pattern.cells.setdefault(col, []).append(r)
        return displayed

    def _cell_value(self, c, raw: bool) -> Optional[Cell]:
        t = c.get("t", "n")
        if t == "inlineStr":
            text = "".join(x.text or "" for x in c.iter(NS + "t"))
            return Cell(text, "s") if text != "" else None
        v = c.findtext(NS + "v")
        if v is None or v == "":
            return None
        if t == "s":
            text = self.shared[int(v)] if int(v) < len(self.shared) else ""
            return Cell(text, "s") if text != "" else None
        if t == "str":
            return Cell(v, "s")
        if t == "b":
            return Cell("TRUE" if v in ("1", "true") else "FALSE", "b")
        if t == "e":
            return Cell(v, "e")
        if t == "d":
            return Cell(v.replace("T00:00:00", "").replace("T", " "), "d")
        try:
            number = float(v)
        except ValueError:
            return Cell(v, "s")
        style = int(c.get("s", "0"))
        code = self.formats[style] if style < len(self.formats) else None
        if code and is_date_code(code):
            return Cell(_date_text(number, code, self.date1904), "d")
        return Cell(format_number(number, code, self.date1904, raw), "n")

    def _describe_filter(self, af) -> List[str]:
        bounds = parse_range(af.get("ref", "") or "")
        first_col = bounds[1] if bounds else 1
        out = []
        for fc in af.iter(NS + "filterColumn"):
            col = col_letter(first_col + int(fc.get("colId", "0")))
            values = [f.get("val") for f in fc.iter(NS + "filter") if f.get("val") is not None]
            parts = []
            if values:
                shown = "、".join(values[:15]) + (f" …（共 {len(values)} 個值）" if len(values) > 15 else "")
                parts.append(f"只顯示 {shown}")
            filters = fc.find(NS + "filters")
            if filters is not None and filters.get("blank") in ("1", "true"):
                parts.append("包含空白")
            ops = {"equal": "=", "notEqual": "≠", "greaterThan": ">", "greaterThanOrEqual": "≥",
                   "lessThan": "<", "lessThanOrEqual": "≤"}
            customs = [f"{ops.get(cf.get('operator', 'equal'), cf.get('operator'))} {cf.get('val')}"
                       for cf in fc.iter(NS + "customFilter")]
            if customs:
                cfs = fc.find(NS + "customFilters")
                joiner = " 且 " if cfs is not None and cfs.get("and") in ("1", "true") else " 或 "
                parts.append(joiner.join(customs))
            if fc.find(NS + "top10") is not None:
                t10 = fc.find(NS + "top10")
                parts.append(f"前 {t10.get('val')}{'%' if t10.get('percent') in ('1', 'true') else ' 項'}")
            if fc.find(NS + "colorFilter") is not None:
                parts.append("依顏色篩選")
            if fc.find(NS + "dynamicFilter") is not None:
                parts.append(f"動態篩選（{fc.find(NS + 'dynamicFilter').get('type')}）")
            out.append(f"{col} 欄：" + ("；".join(parts) if parts else "有篩選條件"))
        return out

    @staticmethod
    def _describe_validation(dv) -> str:
        kind = dv.get("type", "any")
        names = {"list": "清單", "whole": "整數", "decimal": "小數", "date": "日期", "time": "時間",
                 "textLength": "文字長度", "custom": "自訂公式", "any": "任何值"}
        ops = {"between": "介於", "notBetween": "不介於", "equal": "等於", "notEqual": "不等於",
               "greaterThan": "大於", "lessThan": "小於", "greaterThanOrEqual": "大於等於",
               "lessThanOrEqual": "小於等於"}
        if dv.tag == X14 + "dataValidation":  # x14: range and formulas are child elements
            sqref = dv.findtext(XM + "sqref") or ""
            f1 = dv.findtext(f"{X14}formula1/{XM}f") or ""
            f2 = dv.findtext(f"{X14}formula2/{XM}f") or ""
        else:
            sqref = dv.get("sqref", "")
            f1 = dv.findtext(NS + "formula1") or ""
            f2 = dv.findtext(NS + "formula2") or ""
        ranges = sqref.split()
        where = "、".join(ranges[:4]) + (f" 等 {len(ranges)} 段" if len(ranges) > 4 else "")
        if kind == "list":
            rule = f"選項：{f1.strip(chr(34)).replace(',', '、')}" if f1.startswith('"') else f"選項來自 {f1}"
        elif kind == "custom":
            rule = f"公式 ={f1}"
        elif f1:
            op = ops.get(dv.get("operator", "between"), dv.get("operator", ""))
            rule = f"{op} {f1}" + (f" 和 {f2}" if f2 else "")
        else:
            rule = ""
        text = f"{where}：{names.get(kind, kind)}" + (f"，{rule}" if rule else "")
        prompt = dv.get("prompt")
        if prompt:
            text += f"（提示：{prompt}）"
        return text

    def _read_sheet_parts(self, sheet: Sheet) -> None:
        legacy: Dict[str, List[Tuple[str, str]]] = {}
        threaded: Dict[str, List[Tuple[str, str]]] = {}
        for kind, target in self._rels(sheet.path).values():
            if target not in self.names:
                continue
            if kind == "comments":
                root = self._xml(target)
                authors = [a.text or "" for a in root.iter(NS + "author")]
                for cm in root.iter(NS + "comment"):
                    text = "".join(t.text or "" for t in cm.iter(NS + "t") if t.getparent().tag != NS + "rPh")
                    author_id = int(cm.get("authorId", "0"))
                    author = authors[author_id] if author_id < len(authors) else ""
                    legacy.setdefault(cm.get("ref"), []).append((author, text.strip()))
            elif kind == "threadedComment":
                for tc in self._xml(target).iter(TC + "threadedComment"):
                    text = "".join(t.text or "" for t in tc.iter(TC + "text"))
                    author = self.persons.get(tc.get("personId", ""), "")
                    threaded.setdefault(tc.get("ref"), []).append((author, text.strip()))
            elif kind == "pivotTable":
                pivot = self._read_pivot(target)
                if pivot:
                    sheet.pivots.append(pivot)
            elif kind == "table":
                root = self._xml(target)
                bounds = parse_range(root.get("ref", ""))
                if bounds:
                    name = root.get("displayName") or root.get("name") or "表格"
                    sheet.tables.append((name, bounds, int(root.get("headerRowCount", "1"))))
        # Threaded comments replace their legacy placeholder ("[Threaded comment] …").
        sheet.comments = {**legacy, **threaded}

    def _read_pivot(self, part: str) -> Optional[Pivot]:
        root = self._xml(part)
        location = root.find(NS + "location")
        bounds = parse_range(location.get("ref", "")) if location is not None else None
        if not bounds:
            return None
        # The header is the row right above the first data row (firstDataRow is a 0-based offset).
        first_data = int(location.get("firstDataRow", "1"))
        pivot = Pivot(name=root.get("name", "樞紐分析表"), ref=bounds,
                      header_row=bounds[0] + max(first_data - 1, 0))
        page_count = int(location.get("rowPageCount", "0"))
        if page_count:
            pivot.page_rows = (max(1, bounds[0] - page_count - 1), bounds[0] - 1)
        fields: List[str] = []
        for kind, target in self._rels(part).values():
            if kind == "pivotCacheDefinition" and target in self.names:
                cache = self._xml(target)
                fields = [cf.get("name", "") for cf in cache.iter(NS + "cacheField")]
                src = cache.find(f"{NS}cacheSource/{NS}worksheetSource")
                if src is not None:
                    where = src.get("name") or src.get("ref", "")
                    pivot.source = f"'{src.get('sheet')}'!{where}" if src.get("sheet") else where
                refreshed = cache.get("refreshedDate")
                if refreshed:
                    try:
                        pivot.refreshed = _date_text(float(refreshed), "yyyy-mm-dd hh:mm", False)
                    except ValueError:
                        pass

        def field_name(index: str) -> str:
            i = int(index)
            if i == -2:
                return "Σ 值"
            return fields[i] if 0 <= i < len(fields) else f"欄位 {i}"

        pivot.rows = [field_name(f.get("x", "0")) for f in root.iterfind(f"{NS}rowFields/{NS}field")]
        pivot.columns = [field_name(f.get("x", "0")) for f in root.iterfind(f"{NS}colFields/{NS}field")]
        pivot.filters = [field_name(f.get("fld", "0")) for f in root.iterfind(f"{NS}pageFields/{NS}pageField")]
        pivot.values = [df.get("name") or field_name(df.get("fld", "0")) for df in root.iterfind(f"{NS}dataFields/{NS}dataField")]
        return pivot


# --------------------------------------------------------------------------- #
# Markdown rendering
# --------------------------------------------------------------------------- #
def _escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\r\n", "\n").replace("\n", "<br>").strip()


def _code(formula: str) -> str:
    """A formula as inline code inside a table cell."""
    body = formula.replace("|", "\\|").replace("\n", " ")
    fence = "``" if "`" in body else "`"
    pad = " " if fence == "``" else ""
    return f"{fence}{pad}{body}{pad}{fence}"


def _looks_like_header(cells: Dict[int, Cell]) -> bool:
    texts = sum(1 for c in cells.values() if c.kind == "s")
    return texts >= 2 and texts >= 0.5 * len(cells)


class SheetRenderer:
    def __init__(self, sheet: Sheet, options: XlsxOptions):
        self.sheet = sheet
        self.options = options
        self.merged_values: Dict[Tuple[int, int], Cell] = {}
        for ref in sheet.merged:
            b = parse_range(ref)
            if not b:
                continue
            top = sheet.cells.get(b[0], {}).get(b[1])
            if top is None:
                continue
            for r in range(b[0], b[2] + 1):
                for c in range(b[1], b[3] + 1):
                    if (r, c) != (b[0], b[1]):
                        self.merged_values[(r, c)] = top
        self.notes: Dict[Tuple[int, int], int] = {}
        for i, ref in enumerate(sorted(sheet.comments, key=self._ref_key), start=1):
            m = _REF.fullmatch(ref or "")
            if m:
                self.notes[(int(m.group(2)), col_index(m.group(1)))] = i

    @staticmethod
    def _ref_key(ref: str):
        m = _REF.fullmatch(ref or "")
        return (int(m.group(2)), col_index(m.group(1))) if m else (0, 0)

    def _visible_row(self, r: int) -> bool:
        return not (self.options.visible_only and r in self.sheet.hidden_rows)

    def _visible_col(self, c: int) -> bool:
        return not (self.options.visible_only and c in self.sheet.hidden_cols)

    # ----- regions ------------------------------------------------------------ #
    def regions(self) -> List[dict]:
        """Pivot tables and Excel tables get their own region; the rest is the main region."""
        claimed: List[Tuple[int, int, int, int]] = []
        regions = []
        for pivot in self.sheet.pivots:
            top = pivot.page_rows[0] or pivot.ref[0]
            bounds = (top, pivot.ref[1], pivot.ref[2], pivot.ref[3])
            claimed.append(bounds)
            regions.append({"kind": "pivot", "pivot": pivot, "bounds": bounds})
        for name, bounds, header_count in self.sheet.tables:
            claimed.append(bounds)
            regions.append({"kind": "table", "name": name, "bounds": bounds,
                            "header": bounds[0] if header_count else 0})
        regions.sort(key=lambda reg: (reg["bounds"][1], reg["bounds"][0]))

        def free(r: int, c: int) -> bool:
            return not any(b[0] <= r <= b[2] and b[1] <= c <= b[3] for b in claimed)

        main_rows = {
            r: {c: cell for c, cell in cols.items() if free(r, c)}
            for r, cols in self.sheet.cells.items()
        }
        main_rows = {r: cols for r, cols in main_rows.items() if cols}
        if main_rows:
            regions.insert(0, {"kind": "main", "rows": main_rows})
        return regions

    def _main_header(self, rows: Dict[int, Dict[int, Cell]]) -> Tuple[int, str]:
        override = self.options.header_rows.get(self.sheet.name)
        if override is not None:
            return override, "手動指定"
        if self.sheet.autofilter:
            b = parse_range(self.sheet.autofilter)
            if b and b[0] in rows:
                return b[0], "依自動篩選範圍"
        frozen_rows = self.sheet.freeze[0]
        if frozen_rows and frozen_rows in rows and _looks_like_header(rows[frozen_rows]):
            return frozen_rows, "依凍結窗格"
        candidates = sorted(rows)[:30]
        if candidates:
            best = max(sum(1 for c in rows[r].values() if c.kind == "s") for r in candidates)
            for r in candidates:
                texts = sum(1 for c in rows[r].values() if c.kind == "s")
                if texts >= 2 and texts >= 0.6 * best and _looks_like_header(rows[r]):
                    return r, "自動判斷"
        return 0, "無表頭"

    # ----- tables ------------------------------------------------------------- #
    def _cell(self, rows, r: int, c: int, header: bool = False) -> Optional[Cell]:
        cell = rows.get(r, {}).get(c)
        if cell is None and header:
            cell = self.merged_values.get((r, c))
        return cell

    def _label_above(self, rows, header_row: int, c: int) -> Optional[str]:
        """Text label above a date/number header cell ("W1(W38)" over "2026-09-21")."""
        if header_row <= 1:
            return None
        cell = self._cell(rows, header_row, c, header=True)
        above = self._cell(rows, header_row - 1, c, header=True)
        if cell is None or cell.kind not in ("n", "d") or above is None or above.kind != "s":
            return None
        return above.text if above.text != cell.text else None

    def _column_names(self, rows, header_row: int, columns: List[int]) -> List[str]:
        names = []
        seen: Dict[str, int] = {}
        for c in columns:
            label = ""
            if header_row:
                cell = self._cell(rows, header_row, c, header=True)
                label = cell.text if cell else ""
                above = self._label_above(rows, header_row, c)
                if above:
                    label = f"{above} {label}".strip()
            label = _escape(label)
            tag = col_letter(c) + ("·隱" if c in self.sheet.hidden_cols else "")
            name = f"{label} ({tag})" if label else tag
            if name in seen:
                seen[name] += 1
                name = f"{name} #{seen[name]}"
            else:
                seen[name] = 1
            own = rows.get(header_row, {}).get(c) if header_row else None
            if own is not None and own.formula:  # e.g. a date header computed with EOMONTH
                name = f"{name}<br>{_code(own.formula)}"
            names.append(name)
        return names

    def _table(self, rows, row_numbers: List[int], columns: List[int], names: List[str]) -> str:
        lines = ["| 列 | " + " | ".join(names) + " |", "|---|" + "---|" * len(names)]
        for r in row_numbers:
            values = []
            for c in columns:
                cell = rows.get(r, {}).get(c)
                text = _escape(cell.text) if cell else ""
                if cell is not None and cell.formula:
                    text = f"{text}<br>{_code(cell.formula)}" if text else _code(cell.formula)
                note = self.notes.get((r, c))
                if note:
                    text = f"{text} [註{note}]".strip()
                values.append(text)
            label = f"{r} (隱)" if r in self.sheet.hidden_rows else str(r)
            lines.append(f"| {label} | " + " | ".join(values) + " |")
        return "\n".join(lines)

    def _used_columns(self, rows, row_numbers: List[int], extra_rows: List[int] = ()) -> List[int]:
        cols = set()
        for r in list(row_numbers) + list(extra_rows):
            cols.update(rows.get(r, {}))
        return sorted(c for c in cols if self._visible_col(c))

    def render(self) -> Tuple[str, str]:
        """(markdown, header description for the overview)."""
        sheet = self.sheet
        parts: List[str] = []
        header_desc = []
        info = []
        if sheet.dimension:
            info.append(f"使用範圍 {sheet.dimension}")
        if sheet.freeze != (0, 0):
            fr, fc = sheet.freeze
            info.append("凍結窗格：" + "、".join(x for x in (f"前 {fr} 列" if fr else "", f"前 {fc} 欄" if fc else "") if x))
        hidden_rows = sorted(sheet.hidden_rows)
        if hidden_rows:
            how = "被篩選或手動隱藏" if sheet.autofilter else "被隱藏"
            action = "已略過" if self.options.visible_only else "列號標示「(隱)」"
            info.append(f"{len(hidden_rows):,} 列{how}（{action}）")
        if sheet.hidden_cols:
            action = "已略過" if self.options.visible_only else "欄名標示「·隱」"
            info.append(f"隱藏欄 {compress_columns(sorted(sheet.hidden_cols))}（{action}）")

        for region in self.regions():
            if region["kind"] == "main":
                text, desc = self._render_main(region["rows"])
                header_desc.append(desc)
            elif region["kind"] == "pivot":
                text, desc = self._render_pivot(region)
                header_desc.append(desc)
            else:
                text, desc = self._render_table(region)
                header_desc.append(desc)
            if text:
                parts.append(text)

        details = self._details()
        out = []
        if info:
            out.append("> " + " · ".join(info))
        out.extend(parts)
        if details:
            out.append(details)
        descs = [d for d in header_desc if d]
        pivot_descs = [d for d in descs if d.startswith("樞紐")]
        if len(pivot_descs) > 3:  # keep the overview row readable
            descs = [d for d in descs if not d.startswith("樞紐")] + [f"{len(pivot_descs)} 個樞紐分析表各自的表頭"]
        return "\n\n".join(out), "；".join(descs)

    def _render_main(self, rows) -> Tuple[str, str]:
        header_row, reason = self._main_header(rows)
        out = []
        visible = [r for r in sorted(rows) if self._visible_row(r)]
        above = [r for r in visible if header_row and r < header_row]
        below = [r for r in visible if r > header_row] if header_row else visible
        header_rows = [header_row - 1, header_row] if header_row else []
        columns = self._used_columns(rows, above + below, header_rows)
        if not columns:
            return "", ""
        names = self._column_names(rows, header_row, columns)
        desc = f"第 {header_row} 列（{reason}）" if header_row else "無表頭（以欄字母命名）"
        if header_row and self._cell_merge_label_used(rows, header_row, columns):
            desc = f"第 {header_row - 1}–{header_row} 列（{reason}，合併上方標籤列）"
        if above:
            span = f"{above[0]}" if above[0] == above[-1] else f"{above[0]}–{above[-1]}"
            out.append(f"### 表頭上方的列（第 {span} 列：加總、參數或說明）\n\n"
                       + self._table(rows, above, columns, names))
        if below:
            title = f"### 資料（表頭：{desc}）" if above else f"表頭：{desc}"
            out.append(title + "\n\n" + self._table(rows, below, columns, names))
        return "\n\n".join(out), desc

    def _cell_merge_label_used(self, rows, header_row: int, columns: List[int]) -> bool:
        return any(self._label_above(rows, header_row, c) for c in columns)

    def _region_rows(self, bounds) -> Dict[int, Dict[int, Cell]]:
        r1, c1, r2, c2 = bounds
        return {
            r: {c: cell for c, cell in cols.items() if c1 <= c <= c2}
            for r, cols in self.sheet.cells.items() if r1 <= r <= r2
        }

    def _render_pivot(self, region) -> Tuple[str, str]:
        pivot: Pivot = region["pivot"]
        rows = self._region_rows(region["bounds"])
        r1, c1, r2, c2 = pivot.ref
        lines = [f"### 樞紐分析表「{pivot.name}」（{col_letter(c1)}{r1}:{col_letter(c2)}{r2}）"]
        facts = []
        if pivot.rows:
            facts.append("列：" + "、".join(pivot.rows))
        if pivot.columns:
            facts.append("欄：" + "、".join(pivot.columns))
        if pivot.values:
            facts.append("值：" + "、".join(pivot.values))
        if pivot.filters:
            facts.append("篩選：" + "、".join(pivot.filters))
        if pivot.source:
            facts.append(f"資料來源：{pivot.source}")
        snapshot = "上次重新整理" + (f"於 {pivot.refreshed}" if pivot.refreshed else "") + "時的結果（快照，無法在 Markdown 中篩選或重新整理）"
        facts.append(snapshot)
        lines.append("\n".join(f"- {f}" for f in facts))
        if pivot.page_rows != (0, 0):
            page = []
            for r in range(pivot.page_rows[0], pivot.page_rows[1] + 1):
                cells = rows.get(r, {})
                if cells:
                    values = [cells[c].text for c in sorted(cells)]
                    page.append(" = ".join(values[:2]) if len(values) >= 2 else values[0])
            if page:
                lines.append("目前的篩選：" + "；".join(page))
        captions = []
        for r in range(r1, pivot.header_row):  # e.g. the "加總 - Qty" caption above the header
            cells = rows.get(r, {})
            if cells:
                captions.append(" ｜ ".join(_escape(cells[c].text) for c in sorted(cells)))
        if captions:
            lines.append("樞紐標題列：" + "；".join(captions))
        body_rows = [r for r in sorted(rows) if r > pivot.header_row and r <= r2 and self._visible_row(r)]
        columns = self._used_columns(rows, body_rows, [pivot.header_row])
        if columns:
            names = self._column_names(rows, pivot.header_row, columns)
            lines.append(self._table(rows, body_rows, columns, names))
        return "\n\n".join(lines), f"樞紐「{pivot.name}」第 {pivot.header_row} 列"

    def _render_table(self, region) -> Tuple[str, str]:
        rows = self._region_rows(region["bounds"])
        header = region["header"]
        body = [r for r in sorted(rows) if r > header and self._visible_row(r)] if header else sorted(rows)
        columns = self._used_columns(rows, body, [header] if header else [])
        if not columns:
            return "", ""
        b = region["bounds"]
        names = self._column_names(rows, header, columns)
        title = f"### 表格「{region['name']}」（{col_letter(b[1])}{b[0]}:{col_letter(b[3])}{b[2]}）"
        return title + "\n\n" + self._table(rows, body, columns, names), f"表格「{region['name']}」第 {header} 列"

    def _details(self) -> str:
        sheet = self.sheet
        blocks = []
        if sheet.comments:
            items = []
            for ref in sorted(sheet.comments, key=self._ref_key):
                m = _REF.fullmatch(ref or "")
                n = self.notes.get((int(m.group(2)), col_index(m.group(1)))) if m else None
                for i, (author, text) in enumerate(sheet.comments[ref]):
                    who = f"（{author}）" if author else ""
                    prefix = f"[註{n}] {ref}{who}" if i == 0 else f"    ↳ 回覆{who}"
                    items.append(f"- {prefix}：{_escape(text)}")
            blocks.append("#### 註解\n\n" + "\n".join(items))
        if sheet.autofilter:
            text = f"#### 自動篩選\n\n- 範圍：{sheet.autofilter}"
            for fc in sheet.filter_columns:
                text += f"\n- {fc}"
            blocks.append(text)
        if sheet.validations:
            blocks.append("#### 資料驗證\n\n" + "\n".join(f"- {v}" for v in sheet.validations))
        if sheet.merged:
            shown = "、".join(sheet.merged[:40]) + (f" …（共 {len(sheet.merged)} 個）" if len(sheet.merged) > 40 else "")
            blocks.append(f"#### 合併儲存格\n\n{shown}（表格中只在左上角儲存格顯示內容）")
        if sheet.formula_patterns:
            blocks.append(self._formula_section())
        return "\n\n".join(blocks)

    def _formula_section(self) -> str:
        sheet = self.sheet
        patterns = sorted(sheet.formula_patterns.values(), key=lambda p: (p.first_ref[1], p.first_ref[0]))
        lines = [
            f"#### 公式（共 {sheet.formula_count:,} 個，{len(patterns):,} 種寫法）",
            "",
            "往下或往右複製的同一條公式算一種寫法；「公式」是該寫法第一格的內容，其他格的參照依相對位置移動。",
            "",
            "| 範圍 | 格數 | 公式（第一格） |",
            "|---|---|---|",
        ]
        for p in patterns:
            first = f"{col_letter(p.first_ref[1])}{p.first_ref[0]}"
            where = describe_cells(p.cells)
            lines.append(f"| {_escape(where)} | {p.count:,} | {first}：{_code(p.example)} |")
        return "\n".join(lines)


def render_workbook(archive: zipfile.ZipFile, options: XlsxOptions, title: str = "") -> str:
    reader = WorkbookReader(archive, options)
    sheets = reader.read()
    visible = [s for s in sheets if s.state == "visible"]
    hidden = [s for s in sheets if s.state != "visible"]
    if options.visible_only:
        hidden = []

    sections, overview, loaded = [], [], []
    for group, sheet_list in (("visible", visible), ("hidden", hidden)):
        if group == "hidden" and sheet_list:
            sections.append("# 隱藏的工作表\n\n以下工作表在 Excel 中是隱藏的。")
        for sheet in sheet_list:
            reader.load_sheet(sheet)
            body, header_desc = SheetRenderer(sheet, options).render()
            state = {"visible": "顯示", "hidden": "隱藏", "veryHidden": "深度隱藏"}.get(sheet.state, sheet.state)
            suffix = "" if sheet.state == "visible" else f"（{state}工作表）"
            sections.append(f"## {sheet.name}{suffix}\n\n{body}".rstrip())
            formula_desc = (f"{sheet.formula_count:,}（{len(sheet.formula_patterns):,} 種）"
                            if sheet.formula_count else "")
            overview.append(
                f"| {_escape(sheet.name)} | {state} | {sheet.dimension or '-'} | {header_desc or '-'} | "
                f"{formula_desc} | {len(sheet.pivots) or ''} | {sum(len(v) for v in sheet.comments.values()) or ''} |"
            )
            loaded.append(sheet)
            sheet.cells.clear()  # free memory before the next sheet
            for pattern in sheet.formula_patterns.values():
                pattern.cells.clear()  # positions were only needed for this sheet's section

    head = [f"# {title}" if title else "# Excel 活頁簿"]
    if options.recalc_note:
        values = f"數值：{options.recalc_note}。"
    else:
        values = "數值是 Excel 上次存檔時算好的結果（轉換時沒有重新計算，樞紐分析表也是上次重新整理的內容）。"
    formulas = {
        "none": "",
        "summary": "每張工作表最後的「公式」段落列出所有公式的寫法與範圍。",
        "cells": "公式附在每個儲存格的值下方（`=…`），每張工作表最後也有公式寫法的整理。",
    }[options.formulas]
    head.append(
        f"> {values}{formulas}數字依 Excel 儲存格格式顯示（raw 模式則為完整精度）。"
        "欄名後的括號是 Excel 欄字母，「列」是 Excel 列號；「(隱)」/「·隱」表示 Excel 中被隱藏的列或欄。"
    )
    head.append("| 工作表 | 狀態 | 範圍 | 表頭 | 公式 | 樞紐分析表 | 註解 |\n|---|---|---|---|---|---|---|\n" + "\n".join(overview))
    dependencies = _dependency_table(loaded)
    if dependencies:
        head.append(dependencies)
    return "\n\n".join(head + sections)


def _dependency_table(sheets: List[Sheet]) -> str:
    """Which sheets each sheet's formulas and pivot tables read from."""
    rows = []
    for sheet in sheets:
        reads = sorted(sheet.references.items(), key=lambda kv: -kv[1])
        formula_text = "、".join(f"{_escape(name)}（{count:,} 格）" for name, count in reads)
        sources: Dict[str, List[str]] = {}
        for pivot in sheet.pivots:
            name = pivot.source.split("!")[0].strip("'").replace("''", "'") if "!" in pivot.source else pivot.source
            if name:
                sources.setdefault(name, []).append(pivot.name)
        pivot_text = "、".join(
            f"{_escape(name)}（{len(names)} 個樞紐分析表）" if len(names) > 1 else f"{_escape(name)}（{_escape(names[0])}）"
            for name, names in sources.items()
        )
        if formula_text or pivot_text:
            rows.append(f"| {_escape(sheet.name)} | {formula_text or '-'} | {pivot_text or '-'} |")
    if not rows:
        return ""
    return (
        "### 工作表之間的資料流向\n\n"
        "改了右邊兩欄列出的工作表，左邊的工作表就會跟著變（公式在 Excel 存檔時重算；樞紐分析表要按「重新整理」）。\n\n"
        "| 工作表 | 公式讀取的工作表 | 樞紐分析表的資料來源 |\n|---|---|---|\n" + "\n".join(rows)
    )


class XlsxConverter(DocumentConverter):
    """Drop-in replacement for MarkItDown's .xlsx converter (register with priority 0)."""

    def __init__(self, options: Optional[XlsxOptions] = None):
        super().__init__()
        self.options = options or XlsxOptions()

    def accepts(self, file_stream: BinaryIO, stream_info: StreamInfo, **kwargs: Any) -> bool:
        extension = (stream_info.extension or "").lower()
        mimetype = (stream_info.mimetype or "").lower()
        return extension in ACCEPTED_EXTENSIONS or any(mimetype.startswith(p) for p in ACCEPTED_MIME_PREFIXES)

    def convert(self, file_stream: BinaryIO, stream_info: StreamInfo, **kwargs: Any) -> DocumentConverterResult:
        options = XlsxOptions.from_kwargs(self.options, kwargs)
        title = stream_info.filename or ""
        with zipfile.ZipFile(file_stream) as archive:
            markdown = render_workbook(archive, options, title)
        return DocumentConverterResult(markdown=markdown, title=title or None)
