#!/usr/bin/env python3
"""End-to-end smoke test for the MarkItDown skill script or the packaged exe.

Generates small sample documents of every major format, converts them through
the given command, and checks that the expected text comes out.

    python tests/smoke_test.py                                   # test skills/markitdown/scripts/convert.py
    python tests/smoke_test.py --cmd dist/markitdown.exe         # test the Windows build
"""

from __future__ import annotations

import argparse
import json
import shlex
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CMD = [sys.executable, str(ROOT / "skills" / "markitdown" / "scripts" / "convert.py")]

MARKER = "MarkItDownSmoke"
CJK = "繁體中文測試"


def make_pdf(path: Path) -> None:
    """Write a minimal one-page PDF with a line of Helvetica text."""
    text = f"{MARKER} PDF page"
    stream = f"BT /F1 24 Tf 72 720 Td ({text}) Tj ET".encode("latin-1")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for offset in offsets:
        out += b"%010d 00000 n \n" % offset
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, xref)
    path.write_bytes(bytes(out))


def make_docx(path: Path) -> None:
    """Write a minimal .docx (WordprocessingML) with a heading and a paragraph."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        f"<w:p><w:r><w:t>{MARKER} Word document</w:t></w:r></w:p>"
        f"<w:p><w:r><w:t>{CJK}</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)


def make_xlsx(path: Path) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Value"])
    ws.append([f"{MARKER} Excel", 42])
    ws.append([CJK, 7])
    wb.save(path)


def make_pptx(path: Path) -> None:
    from pptx import Presentation

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"{MARKER} PowerPoint"
    slide.placeholders[1].text = CJK
    prs.save(path)


def make_samples(folder: Path) -> dict:
    """Create sample files; returns {filename: expected substring}."""
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "page.html").write_text(
        f"<html><head><title>t</title></head><body><h1>{MARKER} HTML</h1>"
        f"<p>{CJK}</p><ul><li>item</li></ul></body></html>",
        encoding="utf-8",
    )
    (folder / "table.csv").write_text(f"name,value\n{MARKER} CSV,1\n{CJK},2\n", encoding="utf-8")
    (folder / "data.json").write_text(json.dumps({"title": f"{MARKER} JSON"}), encoding="utf-8")
    (folder / "notebook.ipynb").write_text(
        json.dumps(
            {
                "cells": [{"cell_type": "markdown", "metadata": {}, "source": [f"# {MARKER} Notebook"]}],
                "metadata": {},
                "nbformat": 4,
                "nbformat_minor": 5,
            }
        ),
        encoding="utf-8",
    )
    make_pdf(folder / "doc.pdf")
    make_docx(folder / "doc.docx")
    make_xlsx(folder / "sheet.xlsx")
    make_pptx(folder / "slides.pptx")
    with zipfile.ZipFile(folder / "bundle.zip", "w") as z:
        z.writestr("inner.html", f"<h2>{MARKER} ZIP</h2>")
    # doc.pdf and doc.docx both want doc.md, which exercises the collision naming;
    # sub/ exercises the recursive scan.
    (folder / "sub").mkdir(exist_ok=True)
    (folder / "sub" / "nested.html").write_text(f"<p>{MARKER} nested</p>", encoding="utf-8")
    return {
        "page.html": f"{MARKER} HTML",
        "table.csv": f"{MARKER} CSV",
        "data.json": f"{MARKER} JSON",
        "notebook.ipynb": f"{MARKER} Notebook",
        "doc.pdf": f"{MARKER} PDF page",
        "doc.docx": f"{MARKER} Word document",
        "sheet.xlsx": f"{MARKER} Excel",
        "slides.pptx": f"{MARKER} PowerPoint",
        "bundle.zip": f"{MARKER} ZIP",
    }


def run(cmd: list, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd + list(args),
        capture_output=True,
        timeout=300,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--cmd", help="command to test (default: the skill's convert.py)")
    args = parser.parse_args()
    if args.cmd:
        cmd = [part.strip('"') for part in shlex.split(args.cmd, posix=(sys.platform != "win32"))]
        if Path(cmd[0]).is_file():
            cmd[0] = str(Path(cmd[0]).resolve())
    else:
        cmd = DEFAULT_CMD

    failures = []

    def check(name: str, ok: bool, detail: str = "") -> None:
        print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  -- {detail}" if detail and not ok else ""))
        if not ok:
            failures.append(name)

    proc = run(cmd, "--version")
    out = proc.stdout.decode("utf-8", "replace")
    check("--version", proc.returncode == 0 and "markitdown" in out, out + proc.stderr.decode("utf-8", "replace"))

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        samples = tmp / "samples"
        expected = make_samples(samples)

        # 1. Single file -> stdout, for every format.
        for name, needle in expected.items():
            proc = run(cmd, str(samples / name))
            out = proc.stdout.decode("utf-8", "replace")
            check(f"stdout {name}", proc.returncode == 0 and needle in out, (out + proc.stderr.decode("utf-8", "replace"))[-800:])

        # 2. Unicode survives the stdout pipe.
        proc = run(cmd, str(samples / "doc.docx"))
        check("stdout keeps CJK text", CJK in proc.stdout.decode("utf-8", "replace"))

        # 3. -o writes a UTF-8 file.
        out_file = tmp / "single" / "out.md"
        proc = run(cmd, str(samples / "sheet.xlsx"), "-o", str(out_file))
        check(
            "-o output file",
            proc.returncode == 0 and out_file.exists() and CJK in out_file.read_text(encoding="utf-8"),
            proc.stderr.decode("utf-8", "replace"),
        )

        # 4. Folder batch with --out-dir and -r mirrors sub-folders.
        out_dir = tmp / "batch"
        proc = run(cmd, str(samples), "-r", "--out-dir", str(out_dir))
        batch_log = proc.stdout.decode("utf-8", "replace") + proc.stderr.decode("utf-8", "replace")
        check("folder batch exit code", proc.returncode == 0, batch_log[-1500:])
        produced = sorted(p.relative_to(out_dir).as_posix() for p in out_dir.rglob("*.md"))
        want = {"doc.md", "doc.pdf.md", "sub/nested.md", "page.md", "sheet.md", "slides.md", "bundle.md"}
        check("folder batch outputs", want.issubset(produced), f"got {produced}")

        # 5. Several files without --out-dir -> written next to the sources.
        proc = run(cmd, str(samples / "page.html"), str(samples / "table.csv"))
        check(
            "multi-file beside sources",
            proc.returncode == 0 and (samples / "page.md").exists() and (samples / "table.md").exists(),
            proc.stderr.decode("utf-8", "replace"),
        )

        # 6. stdin with an extension hint.
        proc = subprocess.run(cmd + ["-x", "html"], input=(samples / "page.html").read_bytes(), capture_output=True, timeout=300)
        check("stdin with -x", proc.returncode == 0 and f"{MARKER} HTML" in proc.stdout.decode("utf-8", "replace"))

        # 7. A missing file fails cleanly with a non-zero exit code.
        proc = run(cmd, str(samples / "does-not-exist.pdf"))
        check("missing file -> exit 1", proc.returncode == 1)

    print()
    if failures:
        print(f"{len(failures)} check(s) failed: {', '.join(failures)}")
        return 1
    print("All smoke tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
