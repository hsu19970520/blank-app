"""Token-efficient, lossless Excel package for Claude.

Instead of one big Markdown table, a workbook becomes:

* ``00_索引與邏輯.md`` — small: every sheet's columns with type and exact
  statistics computed here (so Claude never has to add up thousands of numbers
  itself), the formula logic per column, the data flow between sheets, and the
  estimated token cost of each data file;
* one data file per sheet (split into parts when large): tab-separated rows,
  full-precision values, the Excel row number first.

Small workbooks fit in one file (index + data). Every data file is parsed
back and compared with the workbook cell by cell before it is written, so a
package is either exactly right or not produced at all.
"""

from __future__ import annotations

import math
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from xlsx_converter import (
    Cell,
    Sheet,
    SheetRenderer,
    WorkbookReader,
    XlsxOptions,
    _dependency_table,
    col_letter,
    compress_columns,
    describe_cells,
    general_text,
    shift_formula,
)

INDEX_NAME = "00_索引與邏輯.md"
DEFAULT_BUDGET = 150_000  # estimated tokens: up to this, everything goes in one file
DEFAULT_PART_BUDGET = 100_000  # estimated tokens per data file when splitting

FORMAT_NOTE = (
    "資料格式：每張工作表是一段 Tab 分隔的資料。第一欄是 Excel 列號，列號後面有 `h` 表示該列在 Excel 中是隱藏的"
    "（常見原因是被篩選掉）。第一行列出 Excel 欄字母，各欄的名稱、類型和統計請看索引。空白代表儲存格是空的。"
    "數值是完整精度（最多 15 位有效數字，沒有千分位）；百分比欄以小數表示（0.25 = 25%）；日期為 YYYY-MM-DD（格式含時間的為 YYYY-MM-DD hh:mm:ss）。"
    "儲存格內的 Tab、換行和反斜線分別寫成 `\\t`、`\\n`、`\\\\`。"
)


# --------------------------------------------------------------------------- #
# Token estimate
# --------------------------------------------------------------------------- #
_CJK = re.compile(r"[\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF]")
_DIGITS = re.compile(r"\d+")


def estimate_tokens(text: str) -> int:
    """Rough Claude token count, deliberately on the high side.

    Calibrated against a Claude BPE tokenizer: digit runs cost about one token per
    3 digits, tabs/newlines about one each, CJK characters about 1.5, other
    characters about one per 3.5. Real counts vary by model."""
    cjk = len(_CJK.findall(text))
    digit_tokens = sum(math.ceil(len(m) / 3) for m in _DIGITS.findall(text))
    digits = sum(len(m) for m in _DIGITS.findall(text))
    separators = text.count("\t") + text.count("\n")
    other = len(text) - cjk - digits - separators
    return int((cjk * 1.5 + digit_tokens * 0.95 + separators + other / 3.0) * 1.1)


# --------------------------------------------------------------------------- #
# Cell text <-> TSV field
# --------------------------------------------------------------------------- #
def escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("\t", "\\t").replace("\r", "\\r").replace("\n", "\\n")


def unescape(field_text: str) -> str:
    out, i = [], 0
    while i < len(field_text):
        ch = field_text[i]
        if ch == "\\" and i + 1 < len(field_text):
            nxt = field_text[i + 1]
            out.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}.get(nxt, "\\" + nxt))
            i += 2
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .") or "sheet"


# --------------------------------------------------------------------------- #
# Package
# --------------------------------------------------------------------------- #
@dataclass
class CompactOptions:
    budget: int = DEFAULT_BUDGET
    part_budget: int = DEFAULT_PART_BUDGET
    visible_only: bool = False
    contains: List[str] = field(default_factory=list)  # keep only rows containing one of these
    header_rows: Dict[str, int] = field(default_factory=dict)
    recalc_note: str = ""


@dataclass
class Package:
    files: Dict[str, str]  # file name -> content (one entry when single)
    single: bool
    tokens: Dict[str, int]  # file name -> estimated tokens
    full_tokens: int  # what the whole workbook's data would cost
    rows_written: int = 0

    @property
    def index(self) -> str:
        return next(iter(self.files.values()))


@dataclass
class SheetData:
    sheet: Sheet
    header: int
    reason: str
    columns: List[int]
    names: Dict[int, str]
    rows: List[int]  # rows written (after filters)
    all_rows: int
    blocks: List[str]  # tab-separated data, already verified
    tokens: int
    legend: List[str]
    facts: List[str] = field(default_factory=list)
    category: Optional[int] = None  # column whose values classify the rows (e.g. Demand / Balance)


class PackageBuilder:
    def __init__(self, archive: zipfile.ZipFile, title: str, options: CompactOptions):
        self.archive = archive
        self.title = title
        self.options = options
        self.needles = [n.lower() for n in options.contains if n.strip()]
        self.xopts = XlsxOptions(header_rows=dict(options.header_rows), visible_only=options.visible_only,
                                 raw_values=True, formulas="summary", recalc_note=options.recalc_note)

    def build(self) -> Package:
        reader = WorkbookReader(self.archive, self.xopts)
        sheets = reader.read()
        if self.options.visible_only:
            sheets = [s for s in sheets if s.state == "visible"]
        data: List[SheetData] = []
        full_tokens = 0
        for sheet in sheets:
            reader.load_sheet(sheet)
            item = self._sheet(sheet)
            full_tokens += item.tokens if not self.needles else self._unfiltered_tokens
            data.append(item)
            sheet.cells.clear()
            for pattern in sheet.formula_patterns.values():
                pattern.cells.clear()

        index_text = self._index(sheets, data, full_tokens)
        data_tokens = sum(d.tokens for d in data)
        single = estimate_tokens(index_text) + data_tokens <= self.options.budget
        files: Dict[str, str] = {}
        if single:
            body = [index_text] + [self._sheet_file_text(d, None, d.blocks) for d in data if d.blocks]
            files[INDEX_NAME] = "\n\n".join(body)
        else:
            files[INDEX_NAME] = index_text
            for n, d in enumerate(data, start=1):
                if not d.blocks:
                    continue
                parts = self._split(d.blocks)
                for p, part in enumerate(parts, start=1):
                    suffix = f"_第{p}部分" if len(parts) > 1 else ""
                    name = f"{n:02d}_{safe_filename(d.sheet.name)}{suffix}.md"
                    files[name] = self._sheet_file_text(d, (p, len(parts)) if len(parts) > 1 else None, part)
            files[INDEX_NAME] = index_text + "\n\n" + self._file_list(files)
        tokens = {name: estimate_tokens(text) for name, text in files.items()}
        return Package(files=files, single=single, tokens=tokens, full_tokens=full_tokens,
                       rows_written=sum(len(d.rows) for d in data))

    # ----- one sheet -------------------------------------------------------- #
    def _sheet(self, sheet: Sheet) -> SheetData:
        cells = sheet.cells
        renderer = SheetRenderer(sheet, self.xopts)
        header, reason = renderer._main_header(cells) if cells else (0, "")
        all_columns = sorted({c for row in cells.values() for c in row})
        columns = [c for c in all_columns if not (self.options.visible_only and c in sheet.hidden_cols)]
        names = {c: self._column_name(renderer, header, c) for c in columns}
        rows = [r for r in sorted(cells) if not (self.options.visible_only and r in sheet.hidden_rows)]
        all_rows = len(rows)
        self._unfiltered_tokens = 0
        if self.needles:
            self._unfiltered_tokens = estimate_tokens("\n".join(self._line(sheet, r, columns) for r in rows))
            keep_context = {header, header - 1} if header else set()
            rows = [r for r in rows if r in keep_context or self._matches(cells.get(r, {}))]
            if not any(r not in keep_context for r in rows):
                rows = []
        lines = [self._line(sheet, r, columns) for r in rows]
        if lines:
            lines.insert(0, "列\t" + "\t".join(col_letter(c) for c in columns))
        self._verify(sheet, rows, columns, lines)
        text = "\n".join(lines)
        category = self._category_column(sheet, self._body(sheet, header), columns)
        item = SheetData(sheet, header, reason, columns, names, rows, all_rows, lines, estimate_tokens(text),
                         self._legend(sheet, columns, names, header, category), category=category)
        item.facts = self._facts(item)  # needs formula positions, which are freed after this sheet
        return item

    def _matches(self, row: Dict[int, Cell]) -> bool:
        return any(any(n in cell.text.lower() for n in self.needles) for cell in row.values())

    @staticmethod
    def _column_name(renderer: SheetRenderer, header: int, c: int) -> str:
        if not header:
            return ""
        cell = renderer._cell(renderer.sheet.cells, header, c, header=True)
        label = cell.text if cell else ""
        above = renderer._label_above(renderer.sheet.cells, header, c)
        return " ".join(f"{above or ''} {label}".split())

    @staticmethod
    def _line(sheet: Sheet, r: int, columns: List[int]) -> str:
        row = sheet.cells.get(r, {})
        values = [escape(row[c].text) if c in row else "" for c in columns]
        while values and values[-1] == "":
            values.pop()
        label = f"{r}h" if r in sheet.hidden_rows else str(r)
        return label + ("\t" + "\t".join(values) if values else "")

    @staticmethod
    def _verify(sheet: Sheet, rows: List[int], columns: List[int], lines: List[str]) -> None:
        """Parse the TSV back and compare it with the workbook, cell by cell."""
        if not lines:
            return
        letters = lines[0].split("\t")[1:]
        if letters != [col_letter(c) for c in columns]:
            raise ValueError(f"{sheet.name}: column header mismatch")
        for r, line in zip(rows, lines[1:]):
            fields = line.split("\t")
            label = fields[0]
            if label != (f"{r}h" if r in sheet.hidden_rows else str(r)):
                raise ValueError(f"{sheet.name}: row label mismatch at {r}")
            values = [unescape(f) for f in fields[1:]]
            values += [""] * (len(columns) - len(values))
            row = sheet.cells.get(r, {})
            for c, value in zip(columns, values):
                expected = row[c].text if c in row else ""
                if value != expected:
                    raise ValueError(f"{sheet.name}!{col_letter(c)}{r}: wrote {value!r}, workbook has {expected!r}")

    # ----- index ------------------------------------------------------------ #
    def _body(self, sheet: Sheet, header: int) -> List[int]:
        return [r for r in sorted(sheet.cells) if r > header and not (self.options.visible_only and r in sheet.hidden_rows)]

    def _legend(self, sheet: Sheet, columns: List[int], names: Dict[int, str], header: int,
                category: Optional[int]) -> List[str]:
        """Column table with exact statistics over the data rows (below the header).

        Runs of adjacent columns built from the same formula (e.g. 76 week columns)
        share one line; when the sheet has a "row type" column, totals are split by
        it, because adding Demand rows to Balance rows would be meaningless."""
        body = self._body(sheet, header)
        formulas = self._formulas_by_column(sheet)
        group_of = {r: sheet.cells[r][category].text for r in body if category and category in sheet.cells[r]}
        out = ["| 欄 | 名稱 | 類型 | 非空 | 統計（表頭以下所有列） | 公式 |", "|---|---|---|---|---|---|"]
        rows_out: List[Tuple[int, str]] = []  # (first column, table row)
        plain: List[Tuple[int, int, List[float], str]] = []  # numeric columns without formulas
        i = 0
        while i < len(columns):
            c = columns[i]
            j = i
            while (j + 1 < len(columns) and columns[j + 1] == columns[j] + 1 and c in formulas
                   and formulas.get(columns[j + 1], (None,))[0] == formulas[c][0]):
                j += 1
            run = columns[i:j + 1]
            if len(run) >= 3:
                rows_out.append((c, self._run_row(sheet, run, names, body, formulas[c])))
                i = j + 1
                continue
            cells = [(r, sheet.cells[r][c]) for r in body if c in sheet.cells.get(r, {})]
            kind, stats = self._column_stats([cell for _, cell in cells], group_of if c != category else {}, [r for r, _ in cells])
            name = (names.get(c) or "") + ("（隱藏欄）" if c in sheet.hidden_cols else "")
            rows_out.append((c, f"| {col_letter(c)} | {_md(name)} | {kind} | {len(cells):,} | {_md(stats)} | {self._formula_text(formulas.get(c))} |"))
            if kind == "數值" and c not in formulas and all(cell.kind == "n" for _, cell in cells):
                plain.append((c, len(cells), [float(cell.text) for _, cell in cells], name))
            i += 1
        # Long runs of adjacent plain numeric columns (e.g. 76 weeks of pasted values) become one line
        # with each column's total, instead of one line per column.
        merged: Dict[int, Optional[str]] = {}
        k = 0
        while k < len(plain):
            m = k
            while m + 1 < len(plain) and plain[m + 1][0] == plain[m][0] + 1:
                m += 1
            if m - k + 1 >= 6:
                merged.update({p[0]: None for p in plain[k:m + 1]})
                merged[plain[k][0]] = self._plain_run_row(sheet, plain[k:m + 1], names, bool(category))
            k = m + 1
        for c, row in rows_out:
            if c not in merged:
                out.append(row)
            elif merged[c]:
                out.append(merged[c])
        return out

    @staticmethod
    def _plain_run_row(sheet: Sheet, run, names: Dict[int, str], grouped: bool) -> str:
        cols = [c for c, _, _, _ in run]
        values = [v for _, _, vs, _ in run for v in vs]
        label = f"{names.get(cols[0]) or ''} … {names.get(cols[-1]) or ''}（{len(cols)} 欄）"
        hidden = sum(1 for c in cols if c in sheet.hidden_cols)
        if hidden:
            label += f"（其中 {hidden} 欄隱藏）"
        stats = f"整段最小 {general_text(min(values), 15)}，最大 {general_text(max(values), 15)}" if values else ""
        if values and not grouped:
            stats += "；各欄合計：" + "、".join(f"{col_letter(c)} {general_text(math.fsum(vs), 15)}" for c, _, vs, _ in run)
        count = sum(n for _, n, _, _ in run)
        return f"| {col_letter(cols[0])}:{col_letter(cols[-1])} | {_md(label)} | 數值 | {count:,} | {stats} |  |"

    def _run_row(self, sheet: Sheet, run: List[int], names: Dict[int, str], body: List[int], formula) -> str:
        values, count = [], 0
        for c in run:
            for r in body:
                cell = sheet.cells.get(r, {}).get(c)
                if cell is not None:
                    count += 1
                    if cell.kind == "n":
                        values.append(float(cell.text))
        first, last = names.get(run[0]) or "", names.get(run[-1]) or ""
        label = f"{first} … {last}（{len(run)} 欄，同一種公式）"
        hidden = sum(1 for c in run if c in sheet.hidden_cols)
        if hidden:
            label += f"（其中 {hidden} 欄隱藏）"
        stats = (f"每欄的值不同，請看資料檔；整段最小 {general_text(min(values), 15)}，最大 {general_text(max(values), 15)}"
                 if values else "")
        span = f"{col_letter(run[0])}:{col_letter(run[-1])}"
        return f"| {span} | {_md(label)} | 數值 | {count:,} | {stats} | {self._formula_text(formula)} |"

    @staticmethod
    def _category_column(sheet: Sheet, body: List[int], columns: List[int]) -> Optional[int]:
        """A column naming the kind of each row in a repeating cycle (e.g. Demand, ETD,
        Balance for every part number). A column that merely has few distinct values
        (Brand, Region) is not one: its rows can be added up across values."""
        best, best_key = None, None
        for c in columns:
            texts = [sheet.cells[r][c].text for r in body if c in sheet.cells.get(r, {}) and sheet.cells[r][c].kind == "s"]
            distinct = len(set(texts))
            if not (2 <= distinct <= 8 and len(texts) >= 0.5 * len(body) and len(texts) >= 20):
                continue
            pairs = Counter(zip(texts, texts[1:]))
            following: Dict[str, Tuple[str, int]] = {}
            for (a, b), n in pairs.items():
                if n > following.get(a, ("", 0))[1]:
                    following[a] = (b, n)
            cyclic = all(b != a for a, (b, _) in following.items())
            agreement = sum(n for _, n in following.values()) / max(1, len(texts) - 1)
            if cyclic and agreement >= 0.9:
                key = (distinct, len(texts))
                if best_key is None or key > best_key:
                    best, best_key = c, key
        return best

    @staticmethod
    def _column_stats(cells: List[Cell], group_of: Optional[Dict[int, str]] = None,
                      rows: Optional[List[int]] = None) -> Tuple[str, str]:
        if not cells:
            return "空", ""
        numbers = [(r, c) for r, c in zip(rows or [0] * len(cells), cells) if c.kind == "n"]
        texts = [c for c in cells if c.kind != "n"]
        if numbers and len(numbers) >= 0.8 * len(cells):
            values = [float(c.text) for _, c in numbers]
            percent = sum(1 for _, c in numbers if c.fmt and "%" in c.fmt) > len(numbers) / 2
            kind = "百分比（小數）" if percent else "數值"
            stats = f"{len(values):,} 個數值，最小 {general_text(min(values), 15)}，最大 {general_text(max(values), 15)}"
            if group_of:
                groups: Dict[str, List[float]] = {}
                for r, c in numbers:
                    groups.setdefault(group_of.get(r, "（無分類）"), []).append(float(c.text))
                stats += "；依分類合計：" + "、".join(f"{g} {general_text(math.fsum(v), 15)}" for g, v in groups.items())
            elif not percent:
                stats += f"，合計 {general_text(math.fsum(values), 15)}"
            if texts:
                other = Counter(c.text for c in texts).most_common(3)
                stats += "；另有文字 " + "、".join(f"{t}×{n}" for t, n in other)
            return kind, stats
        if all(c.kind == "d" for c in cells):
            values = sorted(c.text for c in cells)
            return "日期", f"{values[0]} ～ {values[-1]}，{len(set(values)):,} 個不同日期"
        counts = Counter(c.text for c in cells)
        kind = "文字" if not numbers else "混合"
        if len(counts) <= 12:
            return kind, "、".join(f"{t}×{n}" for t, n in counts.most_common())
        top = "、".join(f"{t}×{n}" for t, n in counts.most_common(5))
        return kind, f"{len(counts):,} 個不同值；最多的是 {top}"

    @staticmethod
    def _formulas_by_column(sheet: Sheet) -> Dict[int, Tuple[str, int, int, str, Tuple[int, int]]]:
        """Per column: (dominant pattern key, formula cells, patterns, example, example cell)."""
        per_col: Dict[int, List[Tuple[int, str, str, Tuple[int, int]]]] = {}
        for key, pattern in sheet.formula_patterns.items():
            for c, rows in pattern.cells.items():
                # Show the formula as written in this column's own first cell, not the pattern's first cell.
                r0, c0 = pattern.first_ref
                r = min(rows)
                example = pattern.example if (r, c) == (r0, c0) else shift_formula(pattern.example, r - r0, c - c0)
                per_col.setdefault(c, []).append((len(rows), key, example, (r, c)))
        out = {}
        for c, items in per_col.items():
            items.sort(key=lambda x: -x[0])
            n, key, example, ref = items[0]
            out[c] = (key, sum(i[0] for i in items), len(items), example, ref)
        return out

    @staticmethod
    def _formula_text(info) -> str:
        if not info:
            return ""
        _key, total, kinds, example, (r0, c0) = info
        text = f"{total:,} 格；`{example.replace('|', chr(92) + '|')}`（{col_letter(c0)}{r0} 的寫法）"
        return text + (f"，另有 {kinds - 1} 種" if kinds > 1 else "")

    def _index(self, sheets: List[Sheet], data: List[SheetData], full_tokens: int) -> str:
        lines = [f"# {self.title or 'Excel 活頁簿'}（給 Claude 的精簡版）", ""]
        values = (f"數值：{self.options.recalc_note}。" if self.options.recalc_note
                  else "數值是 Excel 上次存檔時算好的結果（樞紐分析表也是上次重新整理的內容）。")
        lines.append(f"> {values}{FORMAT_NOTE}")
        lines.append("")
        lines.append("> 給 Claude：請先根據本索引判斷回答問題需要哪些工作表；統計數字已由程式精確算出，可直接引用，"
                     "不要自行加總大量數字。如果需要的資料檔不在對話中，請告訴使用者要上傳哪個檔案。")
        if self.needles:
            lines += ["", f"> ⚠️ 已依關鍵字篩選：資料檔只包含含有「{'、'.join(self.options.contains)}」的列（以及表頭列）。"
                          "下方統計仍是整張工作表的結果。"]
        if self.options.visible_only:
            lines += ["", "> ⚠️ 只輸出 Excel 中看得到的工作表、列和欄；隱藏或被篩掉的資料不在檔案裡，統計也只計算看得到的部分。"]
        lines += ["", "## 工作表", "", "| 工作表 | 狀態 | 列數 | 欄數 | 表頭 | 公式 | 資料估計 token |", "|---|---|---|---|---|---|---|"]
        for d in data:
            s = d.sheet
            state = {"visible": "顯示", "hidden": "隱藏", "veryHidden": "深度隱藏"}.get(s.state, s.state)
            header = f"第 {d.header} 列（{d.reason}）" if d.header else "無"
            written = f"{len(d.rows):,}" + (f" / {d.all_rows:,}" if len(d.rows) != d.all_rows else "")
            formula = f"{s.formula_count:,}（{len(s.formula_patterns)} 種）" if s.formula_count else ""
            lines.append(f"| {_md(s.name)} | {state} | {written} | {len(d.columns)} | {header} | {formula} | {d.tokens:,} |")
        deps = _dependency_table(sheets)
        if deps:
            lines += ["", deps.replace("### ", "## ")]
        for d in data:
            lines += ["", f"## {d.sheet.name}" + ("（隱藏工作表）" if d.sheet.state != "visible" else ""), ""]
            if d.facts:
                lines += [f"- {f}" for f in d.facts] + [""]
            lines += d.legend
        return "\n".join(lines)

    @staticmethod
    def _facts(d: SheetData) -> List[str]:
        s = d.sheet
        facts = []
        if s.hidden_rows:
            facts.append(f"隱藏列：{len(s.hidden_rows):,} 列（資料中列號加 `h`）")
        if s.hidden_cols:
            facts.append(f"隱藏欄：{compress_columns(sorted(s.hidden_cols))}")
        if s.autofilter:
            facts.append(f"自動篩選：{s.autofilter}" + ("；" + "；".join(s.filter_columns) if s.filter_columns else ""))
        if s.pivots:
            facts.append("樞紐分析表：" + "、".join(f"{p.name}（{col_letter(p.ref[1])}{p.ref[0]}:{col_letter(p.ref[3])}{p.ref[2]}，來源 {p.source or '未知'}）" for p in s.pivots[:8])
                         + (f" 等 {len(s.pivots)} 個" if len(s.pivots) > 8 else ""))
        if s.validations:
            facts.append("資料驗證：" + "；".join(s.validations[:4]))
        if s.merged:
            facts.append("合併儲存格：" + "、".join(s.merged[:12]) + (f" 等 {len(s.merged)} 個" if len(s.merged) > 12 else "") + "（值只在左上角）")
        if s.comments:
            facts.append(f"註解：{sum(len(v) for v in s.comments.values())} 則，附在這張工作表資料檔的最後")
        category = d.category
        if category:
            facts.append(f"列的分類：{col_letter(category)} 欄（{_md(d.names.get(category, ''))}）。不同分類的列代表不同的量，"
                         "不能直接相加；下表的合計已依此分類分開計算")
        if s.formula_patterns:
            facts.append(f"公式：{s.formula_count:,} 格、{len(s.formula_patterns)} 種寫法；用得最多的：")
            for p in sorted(s.formula_patterns.values(), key=lambda p: -p.count)[:8]:
                facts.append(f"　{p.count:,} 格｜{describe_cells(p.cells, limit=2)}｜`{p.example.replace('|', chr(92) + '|')}`")
            if len(s.formula_patterns) > 8:
                facts.append(f"　…其餘 {len(s.formula_patterns) - 8} 種見下表各欄")
        return facts

    # ----- data files ------------------------------------------------------- #
    def _split(self, lines: List[str]) -> List[List[str]]:
        header, body = lines[0], lines[1:]
        parts, current, size = [], [header], estimate_tokens(header)
        for line in body:
            cost = estimate_tokens(line) + 1
            if size + cost > self.options.part_budget and len(current) > 1:
                parts.append(current)
                current, size = [header], estimate_tokens(header)
            current.append(line)
            size += cost
        parts.append(current)
        return parts

    def _sheet_file_text(self, d: SheetData, part: Optional[Tuple[int, int]], lines: List[str]) -> str:
        first = lines[1].split("\t")[0].rstrip("h") if len(lines) > 1 else "?"
        last = lines[-1].split("\t")[0].rstrip("h") if len(lines) > 1 else "?"
        title = f"## 資料：{d.sheet.name}" + (f"（第 {part[0]}/{part[1]} 部分）" if part else "")
        note = f"Excel 第 {first}–{last} 列，{len(lines) - 1:,} 列 × {len(d.columns)} 欄。欄名與統計見 {INDEX_NAME}。"
        text = f"{title}\n\n{note}\n\n```tsv\n" + "\n".join(lines) + "\n```"
        if d.sheet.comments and (part is None or part[0] == 1):
            notes = []
            for ref, entries in sorted(d.sheet.comments.items()):
                for i, (author, body) in enumerate(entries):
                    who = f"（{author}）" if author else ""
                    body = " ".join(body.split())
                    notes.append(f"- {ref}{who}：{body}" if i == 0 else f"  - 回覆{who}：{body}")
            text += f"\n\n### 註解（{d.sheet.name}）\n\n" + "\n".join(notes)
        return text

    @staticmethod
    def _file_list(files: Dict[str, str]) -> str:
        rows = ["## 資料檔", "", "| 檔案 | 估計 token |", "|---|---|"]
        for name, text in files.items():
            if name != INDEX_NAME:
                rows.append(f"| {name} | {estimate_tokens(text):,} |")
        return "\n".join(rows)


def _md(text: str) -> str:
    return escape(text).replace("|", "\\|")


def build_package(path_or_stream, title: str = "", options: Optional[CompactOptions] = None) -> Package:
    with zipfile.ZipFile(path_or_stream) as archive:
        return PackageBuilder(archive, title, options or CompactOptions()).build()
