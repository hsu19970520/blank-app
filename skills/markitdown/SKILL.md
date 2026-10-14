---
name: markitdown
description: Convert documents and web pages to Markdown with Microsoft MarkItDown. Handles PDF, Word (.docx), PowerPoint (.pptx), Excel (.xlsx/.xls), HTML, CSV, JSON, XML/RSS, EPUB, ZIP archives, Jupyter notebooks, Outlook .msg, images (EXIF metadata), audio (speech transcription) and URLs (web pages, YouTube transcripts, Wikipedia). Use when the user wants to read, extract text or tables from, summarize, search, or convert such files to Markdown, or batch-convert a whole folder — e.g. "轉成 Markdown", "把 PDF 轉文字", "讀這份 Word/Excel/PPT", "convert this deck to md".
---

# MarkItDown: anything → Markdown

This skill turns office documents, PDFs, and web content into clean Markdown that
preserves headings, lists, tables, and links, which is the best format for an LLM
to read or for saving as notes.

## 1. Pick the runner (once per session)

All commands below are written as `RUN …`. Resolve `RUN` from this skill's base
directory, in this order:

1. **Windows exe (no Python needed):** if `bin/markitdown.exe` exists in this skill
   folder, `RUN` = `"<skill-dir>/bin/markitdown.exe"`.
2. **Python:** otherwise `RUN` = `python "<skill-dir>/scripts/convert.py"`
   (use `python3` on macOS/Linux). If it fails with `No module named 'markitdown'`,
   install the engine first: `pip install "markitdown[all]"`.

Check it works with `RUN --version`.

## 2. Convert

| Goal | Command |
|---|---|
| Read one file (Markdown on stdout) | `RUN report.pdf` |
| Save one file | `RUN report.pdf -o report.md` |
| Several files → `.md` next to each source | `RUN a.docx b.pptx c.xlsx` |
| Whole folder (with sub-folders) into another folder | `RUN ./docs -r --out-dir ./docs_md` |
| Web page / YouTube transcript / Wikipedia | `RUN "https://www.youtube.com/watch?v=ID"` |
| Bytes from stdin | `RUN -x pdf < file.bin` |
| Print a whole batch to stdout | `RUN a.pdf b.pdf --stdout` |

Excel options: `--xlsx-format compact|readable` (default `compact`, the token-saving
package described in section 4), `--xlsx-rows-containing TEXT` (compact: keep only rows
containing TEXT, case-insensitive, repeatable), `--xlsx-latest N` (compact: in sheets
with a version column, i.e. snapshot dates repeated for every item as in a forecast
waterfall, keep only the newest N versions), `--xlsx-part-tokens N` (compact: split
into files of about N tokens), `--xlsx-header-row "SHEET=ROW"` (override the detected
header row; repeatable), `--xlsx-visible-only` (leave out hidden sheets / rows / columns,
i.e. what Excel shows), `--xlsx-raw-values` (readable: full-precision numbers instead of
Excel's display format), `--xlsx-formulas summary|cells|none` (how formulas are shown,
see below; `cells` implies `readable`),
`--xlsx-recalc` (Windows with Microsoft Excel only: let Excel recalculate formulas and
refresh pivot tables in a copy before converting; the CLI exits 2 with the reason when
Excel isn't available), `--xlsx-classic` (MarkItDown's original pandas converter).

Other options: `--skip-existing` (resume a batch), `--keep-data-uris` (keep base64
images), `-c big5` (charset hint for legacy text/CSV files), `-x`/`-m` (extension /
MIME hints), `-p` (enable installed MarkItDown plugins),
`--docintel-endpoint URL` (Azure Document Intelligence for scanned PDFs and
images; reads `MARKITDOWN_DOCINTEL_ENDPOINT` if set).

Exit code is `0` when everything converted, `1` if any input failed (each failure
is reported on stderr as `[FAIL] <input>: <reason>`), and `2` for usage errors.
Batch runs report each item as `(i/n) [OK] src -> dest.md`.

## 3. Workflow guidance

- **Reading for the user:** for a single document, convert to stdout and work from
  that. For anything large (long PDFs, big spreadsheets) write to a file with `-o`
  first, then read it in parts. Don't pull a huge document into context at once.
- **Excel:** convert with `--out-dir <folder>` (a big workbook becomes several files),
  read `00_索引與邏輯.md` first, then only the data files the question needs. When the
  question is about specific items (part numbers, customers), re-run with
  `--xlsx-rows-containing ITEM` (repeatable) so the data holds only those rows.
- **Converting for the user:** write the `.md` files where the user asked; by default
  a batch goes next to the originals. When asked to convert a folder, use
  `--out-dir` so the source folder stays clean, and add `-r` for sub-folders.
- Batch mode never overwrites the source file. If two inputs would map to the
  same name (e.g. `doc.pdf` and `doc.docx`), the second becomes `doc.docx.md`.
- When scanning a folder, `.md`/`.txt` files are skipped. Explicitly named files
  are always attempted.
- After converting, tell the user where the files went and mention any `[FAIL]`
  lines.

## 4. What to expect per format

- **PDF:** text layer only. A scanned PDF with no text layer yields little or no
  output. Say so, and suggest OCR or `--docintel-endpoint`.
- **Excel (.xlsx/.xlsm), default `compact` package:** small workbooks give one `.md`
  file; workbooks over about 30k tokens (or a filtered extract over about 150k) give a folder `<name>_md/` holding
  `00_索引與邏輯.md` (the index) and one file per sheet, split into `_第N部分` parts when
  large. How to use it:
  - **Read the index first.** Per sheet it lists every column (Excel letter, name, type,
    non-empty count) with **exact statistics computed by the converter**: min, max and
    total, split by the row-category column when the sheet has one (e.g. Demand / ETD /
    Balance rows, whose values must not be added together). Quote these numbers instead
    of adding up data yourself. It also gives each column's formula, the most used
    formulas, the data flow between sheets, filters, hidden rows/columns, validation,
    pivot tables, and the estimated tokens of each data file. Adjacent columns built by
    the same formula (e.g. 76 week columns) share one line.
  - **Then read only the data you need.** Data is a ```` ```tsv ```` block: the first line
    is `列` plus Excel column letters, each following line starts with the Excel row
    number (`17h` = row hidden in Excel, usually filtered out), then the cells in
    column order (empty = blank cell; trailing blanks omitted). Values are exact: full
    precision without thousands separators, percentages as fractions (0.25 = 25%), dates
    as `YYYY-MM-DD` or `YYYY-MM-DD hh:mm:ss[.fff]` even when the cell's format shows less.
    `\t`, `\n`, `\\` stand for a tab, line break and backslash inside a cell. Cell
    `C15` is column `C` of the line starting `15`. Comments follow the sheet's data.
  - In a chat where the user uploaded only some of the files, say which file you need
    (the index lists them all) rather than guessing.
  - Before writing, the converter reads every value back and compares it with the
    workbook; a package is never produced with a wrong cell.
- **Excel, `--xlsx-format readable`:** a workbook overview table first (sheet, visible/hidden, range,
  where the header came from), then one `## SheetName` section per sheet; hidden sheets
  are grouped at the end under `# 隱藏的工作表`. In each table the first column `列` is the
  Excel row number and every column name ends with its Excel letter, e.g. `MPN (A)`, so
  you can point the user to exact cells. `(隱)` / `·隱` mark rows / columns hidden in
  Excel (often by an active autofilter, described under `#### 自動篩選`). The header row
  comes from the autofilter, pivot layout or freeze panes; rows above it (totals,
  parameters) are shown separately. Numbers appear exactly as Excel displays them
  (`12,345,678`, `(1,234)`, `206.92%`, ISO dates). Comments (`[註n]`), data validation,
  merged cells and pivot tables (fields, source, last refresh) are listed per sheet.
  Values are the results Excel cached at the last save (pivot tables are snapshots),
  unless `--xlsx-recalc` was used; the line under the title says which. Formulas: each
  sheet ends with `#### 公式`, every distinct formula pattern with where it is used and
  its first cell's formula (copies filled down/across count as one pattern), and the
  overview has `### 工作表之間的資料流向`, which sheets each sheet's formulas and pivot
  tables read. To explain how a number is produced, use those; for one exact cell use
  `--xlsx-formulas cells`, which puts the formula under each value (`54,566,399<br>` +
  `` `=XLOOKUP(...)` ``) but makes the output several times larger. Conditional
  formatting (red/yellow/green) is not rendered yet.
  Big workbooks produce big Markdown: write to a file with `-o` and read it in parts.
- **PowerPoint:** one section per slide, including tables, chart data, and speaker notes.
- **Word:** headings, lists, tables, and links are preserved. Equations become LaTeX.
- **Images:** no OCR offline. Output is EXIF metadata, and only if `exiftool` is
  installed. For text in images, use `--docintel-endpoint`, or look at the image yourself.
- **Audio (.wav/.mp3/.m4a/.mp4):** speech transcription goes through Google's free
  web speech API, so it needs internet. Non-WAV input also needs `ffmpeg` on PATH.
- **ZIP:** each contained file is converted and concatenated.
- **URLs:** need internet access. YouTube returns title, description, and transcript
  when one is available.

## 5. Troubleshooting

- `MissingDependencyException`: run `pip install "markitdown[all]"` (Python runner).
- Garbled Chinese/Japanese text from a `.txt`/`.csv`: retry with `-c big5`, `-c gbk`,
  or `-c shift_jis`.
- The Windows exe takes a few seconds to start (it unpacks itself each run), so
  prefer one batch command over many single-file calls.
