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
- **Excel:** every sheet becomes a `## SheetName` section with a Markdown table.
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
