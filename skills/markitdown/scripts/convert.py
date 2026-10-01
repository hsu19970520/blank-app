#!/usr/bin/env python3
"""Convert files, folders, or URLs to Markdown with Microsoft MarkItDown.

This script is the engine behind the ``markitdown`` Claude skill and the
``markitdown.exe`` Windows build. It wraps the ``markitdown`` library and adds
batch conversion (many files / whole folders) on top of the official CLI.

Examples:
    python convert.py report.pdf                      # Markdown to stdout
    python convert.py report.pdf -o report.md         # write one file
    python convert.py a.docx b.pptx c.xlsx            # write a.md, b.md, c.md beside the sources
    python convert.py ./docs -r --out-dir ./docs_md   # whole folder tree
    python convert.py https://example.com/article     # web page / YouTube / Wikipedia
    type data.bin | python convert.py -x pdf          # read from stdin

Requires: pip install "markitdown[all]"
"""

from __future__ import annotations

import argparse
import codecs
import io
import os
import re
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Set
from urllib.parse import parse_qs, urlparse

# pydub warns at import time when ffmpeg is missing; that only matters for audio.
warnings.filterwarnings("ignore", message=".*ffmpeg.*", category=RuntimeWarning)

# ONNX Runtime (used by Magika for file-type detection) sends usage telemetry to
# Microsoft by default. Converting local documents shouldn't, so opt out before
# it loads. Set ORT_DISABLE_TELEMETRY=0 to allow it.
os.environ.setdefault("ORT_DISABLE_TELEMETRY", "1")

TOOL_VERSION = "1.1.0"

# File types MarkItDown can convert. Used when scanning folders; explicit file
# arguments are always attempted regardless of extension.
SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".pptx",
    ".xlsx",
    ".xls",
    ".html",
    ".htm",
    ".csv",
    ".json",
    ".jsonl",
    ".xml",
    ".rss",
    ".atom",
    ".epub",
    ".ipynb",
    ".msg",
    ".zip",
    ".jpg",
    ".jpeg",
    ".png",
    ".wav",
    ".mp3",
    ".m4a",
    ".mp4",
}

# Text formats that are already Markdown-ish; skipped when scanning folders so a
# re-run never feeds its own output back in.
SKIP_IN_FOLDER_SCAN = {".md", ".markdown", ".txt", ".text"}

URL_SCHEMES = ("http", "https", "file", "data")


@dataclass
class Source:
    """One thing to convert: a local file, a URL, or stdin ("-")."""

    value: str
    # For files found by scanning a folder: the folder they were found in, so
    # --out-dir can mirror the sub-folder structure.
    scan_root: Optional[Path] = None

    @property
    def is_url(self) -> bool:
        return is_url(self.value)

    @property
    def is_stdin(self) -> bool:
        return self.value == "-"


@dataclass
class Result:
    source: Source
    output: Optional[Path]
    markdown: Optional[str]
    error: Optional[str]
    error_type: Optional[str] = None  # exception class name, for friendlier UI messages

    @property
    def ok(self) -> bool:
        return self.error is None


def is_url(value: str) -> bool:
    return urlparse(value.strip()).scheme.lower() in URL_SCHEMES and "://" in value


def collect_sources(inputs: Iterable[str], recursive: bool = False) -> List[Source]:
    """Expand the raw CLI/GUI inputs into a flat list of sources (deduplicated)."""
    sources: List[Source] = []
    seen: Set[str] = set()

    def add(source: Source) -> None:
        key = source.value if (source.is_url or source.is_stdin) else str(Path(source.value).resolve())
        if key not in seen:
            seen.add(key)
            sources.append(source)

    for raw in inputs:
        raw = raw.strip().strip('"')
        if not raw:
            continue
        if raw == "-" or is_url(raw):
            add(Source(raw))
            continue
        path = Path(raw).expanduser()
        if path.is_dir():
            for source in scan_folder(path, recursive):
                add(source)
        else:
            add(Source(str(path)))
    return sources


def scan_folder(folder: Path, recursive: bool) -> List[Source]:
    pattern = "**/*" if recursive else "*"
    found = []
    for p in sorted(folder.glob(pattern)):
        if not p.is_file():
            continue
        name = p.name
        if name.startswith(".") or name.startswith("~$"):  # hidden / Office lock files
            continue
        ext = p.suffix.lower()
        if ext in SKIP_IN_FOLDER_SCAN or ext not in SUPPORTED_EXTENSIONS:
            continue
        found.append(Source(str(p), scan_root=folder))
    return found


def url_to_filename(url: str) -> str:
    parsed = urlparse(url)
    host = parsed.netloc.replace("www.", "")
    video_id = parse_qs(parsed.query).get("v", [None])[0]
    if "youtube" in host and video_id:
        name = f"youtube_{video_id}"
    else:
        last = Path(parsed.path.rstrip("/")).stem
        name = f"{host}_{last}" if last else host or "page"
    name = re.sub(r"[^\w\-.]+", "_", name).strip("._") or "page"
    return name[:80]


def plan_output(
    source: Source,
    out_dir: Optional[Path],
    taken: Set[Path],
) -> Path:
    """Pick a non-colliding .md path for a source (batch mode)."""
    if source.is_url:
        base_dir = out_dir or Path.cwd()
        stem, alt_stem = url_to_filename(source.value), None
    elif source.is_stdin:
        base_dir = out_dir or Path.cwd()
        stem, alt_stem = "stdin", None
    else:
        src = Path(source.value)
        if out_dir is None:
            base_dir = src.parent
        elif source.scan_root is not None:
            base_dir = out_dir / src.parent.relative_to(source.scan_root)
        else:
            base_dir = out_dir
        stem = src.stem
        alt_stem = f"{src.stem}.{src.suffix.lstrip('.')}" if src.suffix else None

    candidates = [base_dir / f"{stem}.md"]
    if alt_stem:
        candidates.append(base_dir / f"{alt_stem}.md")
    for n in range(2, 1000):
        candidates.append(base_dir / f"{stem} ({n}).md")

    src_resolved = None if (source.is_url or source.is_stdin) else Path(source.value).resolve()
    for cand in candidates:
        resolved = cand.resolve()
        if resolved == src_resolved or resolved in taken:
            continue
        taken.add(resolved)
        return cand
    raise RuntimeError(f"Could not find a free output name for {source.value}")


def build_converter(
    use_plugins: bool = False,
    docintel_endpoint: Optional[str] = None,
):
    """Create the MarkItDown instance (slow: loads the Magika model)."""
    from markitdown import MarkItDown

    if os.environ.get("ORT_DISABLE_TELEMETRY") == "1":
        try:
            import onnxruntime

            onnxruntime.disable_telemetry_events()
        except Exception:
            pass

    kwargs = {"enable_plugins": use_plugins}
    if docintel_endpoint:
        kwargs["docintel_endpoint"] = docintel_endpoint
    return MarkItDown(**kwargs)


def make_stream_info(extension: Optional[str], mime_type: Optional[str], charset: Optional[str]):
    if not (extension or mime_type or charset):
        return None
    from markitdown import StreamInfo

    if extension:
        extension = extension.strip().lower()
        if not extension.startswith("."):
            extension = "." + extension
    if charset:
        charset = codecs.lookup(charset.strip()).name
    return StreamInfo(extension=extension or None, mimetype=mime_type or None, charset=charset or None)


def convert_source(converter, source: Source, stream_info=None, keep_data_uris: bool = False) -> str:
    if source.is_stdin:
        data = io.BytesIO(sys.stdin.buffer.read())
        result = converter.convert_stream(data, stream_info=stream_info, keep_data_uris=keep_data_uris)
    else:
        result = converter.convert(source.value, stream_info=stream_info, keep_data_uris=keep_data_uris)
    return result.markdown or ""


def describe_error(exc: BaseException) -> str:
    message = str(exc).strip() or exc.__class__.__name__
    # MarkItDown error messages can be very long tracebacks; keep the gist.
    first_lines = "\n".join(message.splitlines()[:6])
    return f"{exc.__class__.__name__}: {first_lines}"


def convert_batch(
    converter,
    sources: List[Source],
    out_dir: Optional[Path],
    stream_info=None,
    keep_data_uris: bool = False,
    skip_existing: bool = False,
    on_result: Optional[Callable[[int, int, Result], None]] = None,
    on_start: Optional[Callable[[int, int, Source], None]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
) -> List[Result]:
    """Convert every source to its own .md file. Shared by the CLI and the GUI.

    ``should_stop`` is checked between files, so a stop request finishes the
    file in progress and leaves the rest unconverted.
    """
    taken: Set[Path] = set()
    results: List[Result] = []
    total = len(sources)
    for index, source in enumerate(sources, start=1):
        if should_stop and should_stop():
            break
        if on_start:
            on_start(index, total, source)
        output = plan_output(source, out_dir, taken)
        if skip_existing and output.exists():
            result = Result(source, output, None, None)
        else:
            try:
                markdown = convert_source(converter, source, stream_info, keep_data_uris)
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_text(markdown, encoding="utf-8")
                result = Result(source, output, markdown, None)
            except Exception as exc:  # report and keep going with the rest of the batch
                result = Result(source, None, None, describe_error(exc), type(exc).__name__)
        results.append(result)
        if on_result:
            on_result(index, total, result)
    return results


def build_parser(prog: str = "markitdown") -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=prog,
        description=(
            "Convert PDF, Word, PowerPoint, Excel, HTML, CSV, JSON, XML, EPUB, ZIP, "
            "images, audio, Outlook .msg, Jupyter notebooks and web pages (incl. YouTube) "
            "to Markdown using Microsoft MarkItDown."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Output rules:\n"
            "  * one input, no -o/--out-dir  -> Markdown is printed to stdout\n"
            "  * -o FILE                     -> one input written to FILE\n"
            "  * --out-dir DIR               -> every input written to DIR/<name>.md\n"
            "  * several inputs or a folder  -> <name>.md written next to each source\n"
            "\n"
            "Examples:\n"
            f"  {prog} report.pdf\n"
            f"  {prog} report.pdf -o report.md\n"
            f"  {prog} a.docx b.pptx c.xlsx\n"
            f"  {prog} ./docs -r --out-dir ./docs_md\n"
            f"  {prog} https://www.youtube.com/watch?v=VIDEO_ID\n"
        ),
    )
    parser.add_argument("inputs", nargs="*", help="files, folders, URLs, or - for stdin")
    parser.add_argument("-v", "--version", action="store_true", help="show versions and exit")
    parser.add_argument("-o", "--output", help="output file (single input only)")
    parser.add_argument("--out-dir", help="write one .md per input into this folder")
    parser.add_argument("-r", "--recursive", action="store_true", help="include sub-folders when an input is a folder")
    parser.add_argument("--stdout", action="store_true", help="print all results to stdout instead of writing files")
    parser.add_argument("--skip-existing", action="store_true", help="in batch mode, skip inputs whose .md already exists")
    parser.add_argument("--keep-data-uris", action="store_true", help="keep base64 images in the output (truncated by default)")
    parser.add_argument("-x", "--extension", help="file-extension hint, e.g. pdf (useful with stdin)")
    parser.add_argument("-m", "--mime-type", help="MIME-type hint")
    parser.add_argument("-c", "--charset", help="charset hint, e.g. UTF-8 or big5")
    parser.add_argument("-p", "--use-plugins", action="store_true", help="enable installed 3rd-party MarkItDown plugins")
    parser.add_argument(
        "--docintel-endpoint",
        default=os.environ.get("MARKITDOWN_DOCINTEL_ENDPOINT") or None,
        help="Azure Document Intelligence endpoint (default: $MARKITDOWN_DOCINTEL_ENDPOINT)",
    )
    return parser


def versions_text() -> str:
    try:
        from markitdown.__about__ import __version__ as md_version
    except Exception:  # pragma: no cover - only when markitdown is missing
        md_version = "not installed"
    return f"markitdown-skill {TOOL_VERSION} (markitdown {md_version}, Python {sys.version.split()[0]})"


def _use_utf8_stdout() -> None:
    # When output is piped or redirected (e.g. by an agent or `> out.md`), Windows
    # would otherwise encode with the ANSI code page and mangle non-Latin text.
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and not stream.isatty():
                stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv: Optional[List[str]] = None, prog: str = "markitdown") -> int:
    _use_utf8_stdout()
    parser = build_parser(prog)
    args = parser.parse_args(argv)

    if args.version:
        print(versions_text())
        return 0

    inputs = list(args.inputs)
    if not inputs:
        if sys.stdin is not None and not sys.stdin.isatty():
            inputs = ["-"]
        else:
            parser.print_help()
            return 2

    sources = collect_sources(inputs, recursive=args.recursive)
    if not sources:
        print("No convertible files found.", file=sys.stderr)
        return 1

    if args.output and len(sources) != 1:
        parser.error("-o/--output works with exactly one input; use --out-dir for several")
    if args.output and args.out_dir:
        parser.error("use either -o/--output or --out-dir, not both")

    try:
        stream_info = make_stream_info(args.extension, args.mime_type, args.charset)
    except LookupError:
        parser.error(f"unknown charset: {args.charset}")

    converter = build_converter(args.use_plugins, args.docintel_endpoint)

    single_to_stdout = len(sources) == 1 and not args.output and not args.out_dir
    has_folder_input = any(s.scan_root is not None for s in sources)

    if args.stdout or (single_to_stdout and not has_folder_input):
        failures = 0
        for source in sources:
            try:
                markdown = convert_source(converter, source, stream_info, args.keep_data_uris)
            except Exception as exc:
                failures += 1
                print(f"[FAIL] {source.value}: {describe_error(exc)}", file=sys.stderr)
                continue
            if len(sources) > 1:
                print(f"\n<!-- source: {source.value} -->\n")
            print(markdown)
        return 1 if failures else 0

    if args.output:
        source = sources[0]
        try:
            markdown = convert_source(converter, source, stream_info, args.keep_data_uris)
        except Exception as exc:
            print(f"[FAIL] {source.value}: {describe_error(exc)}", file=sys.stderr)
            return 1
        output = Path(args.output)
        if output.parent and not output.parent.exists():
            output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(markdown, encoding="utf-8")
        print(f"[OK]   {source.value} -> {output}")
        return 0

    out_dir = Path(args.out_dir) if args.out_dir else None

    def report(index: int, total: int, result: Result) -> None:
        prefix = f"({index}/{total})"
        if result.error:
            print(f"{prefix} [FAIL] {result.source.value}: {result.error}", file=sys.stderr)
        elif result.markdown is None:
            print(f"{prefix} [SKIP] {result.source.value} (exists: {result.output})")
        else:
            print(f"{prefix} [OK]   {result.source.value} -> {result.output}")
        sys.stdout.flush()

    results = convert_batch(
        converter,
        sources,
        out_dir,
        stream_info=stream_info,
        keep_data_uris=args.keep_data_uris,
        skip_existing=args.skip_existing,
        on_result=report,
    )
    failed = sum(1 for r in results if not r.ok)
    print(f"Done: {len(results) - failed} succeeded, {failed} failed.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
