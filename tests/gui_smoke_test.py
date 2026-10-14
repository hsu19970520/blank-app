#!/usr/bin/env python3
"""Smoke test for the GUI of markitdown.exe.

1. Opens the converter window from source and walks the main flow: empty
   state, adding a folder / file / missing file, converting, previewing a
   result, the friendly error for a failed item, and stopping a run.
2. With --exe, also launches the packaged `markitdown.exe --gui` and checks
   that it stays up (i.e. did not crash on start-up).

    python tests/gui_smoke_test.py --screenshots shots
    python tests/gui_smoke_test.py --appearance dark --screenshots shots
    python tests/gui_smoke_test.py --exe dist/markitdown.exe --screenshots shots
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app"))
sys.path.insert(0, str(ROOT / "tests"))


# The Windows console may not be UTF-8: never let a check's text crash the run.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


def grab(path: Path, bbox=None) -> None:
    try:
        from PIL import ImageGrab

        ImageGrab.grab(bbox=bbox).save(path)
        print(f"screenshot: {path}")
    except Exception as exc:  # screenshots are a nice-to-have
        print(f"(screenshot skipped: {exc})")


def test_window(shots: Path | None, prefix: str, timeout: float = 240) -> list:
    import markitdown_app
    import smoke_test

    failures = []

    def check(name: str, ok: bool, detail: str = "") -> None:
        print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  -- {detail}" if detail and not ok else ""))
        if not ok:
            failures.append(name)

    tmp = Path(tempfile.mkdtemp())
    samples = tmp / "samples"
    smoke_test.make_samples(samples)
    out = tmp / "out"
    missing = samples / "missing.pdf"

    root, app = markitdown_app.create_window([])
    deadline = time.time() + timeout

    def pump(seconds: float = 0.3) -> None:
        end = time.time() + seconds
        while time.time() < end:
            root.update()
            time.sleep(0.02)

    def shot(name: str) -> None:
        if shots is None:
            return
        pump(0.4)
        x, y = root.winfo_rootx(), root.winfo_rooty()
        grab(shots / f"{prefix}{name}.png", (x, y, x + root.winfo_width(), y + root.winfo_height()))

    def wait(condition, what: str) -> bool:
        while time.time() < deadline:
            if condition():
                return True
            pump(0.1)
        failures.append(f"timed out waiting for {what}")
        return False

    def preview_text() -> str:
        return app.preview.get("1.0", "end")

    def leaf(name: str) -> str:
        return next(i for i in app._leaf_iids() if Path(app.items[i].value).name == name)

    pump(0.5)
    check("empty state shown", bool(app.drop_zone.winfo_ismapped()))
    check("convert disabled when empty", str(app.primary.cget("state")) == "disabled")
    shot("1_empty")

    app.add_inputs([str(samples), str(samples / "doc.pdf"), str(missing)])
    leaves = app._leaf_iids()
    names = sorted(Path(app.items[i].value).name for i in leaves)
    check("folder expanded into files (no duplicates)", names.count("doc.pdf") == 1 and "nested.html" in names, str(names))
    check("missing file listed", "missing.pdf" in names)
    pump(0.3)
    check("empty state hidden", not app.drop_zone.winfo_ismapped())
    check("primary button counts items", str(app.primary.cget("text")) == f"轉換 {len(leaves)} 個項目", app.primary.cget("text"))

    app.output_mode.set("folder")
    app.output_dir.set(str(out))
    app._refresh()
    wait(lambda: app.converter is not None, "the conversion engine")
    shot("2_queued")

    app.start()
    check("primary button becomes Stop", str(app.primary.cget("text")) == "停止")
    wait(lambda: not app.busy, "conversion to finish")

    produced = sorted(p.relative_to(out).as_posix() for p in out.rglob("*.md"))
    want = {"doc.md", "doc.pdf.md", "page.md", "sheet.md", "slides.md", "bundle.md", "sub/nested.md"}
    check("outputs written", want.issubset(produced), str(produced))
    check("no duplicate conversion", "doc (2).md" not in produced, str(produced))
    sheet_md = (out / "sheet.md").read_text(encoding="utf-8") if (out / "sheet.md").exists() else ""
    check("Excel gets the compact package by default", "給 Claude 的精簡版" in sheet_md and "```tsv" in sheet_md, sheet_md[:200])
    statuses = {Path(app.items[i].value).name: app.items[i].status for i in app._leaf_iids()}
    check("missing file failed", statuses.get("missing.pdf") == "failed", str(statuses))
    check("others done", all(s == "done" for n, s in statuses.items() if n != "missing.pdf"), str(statuses))
    check("summary in status line", "已轉換" in app.status_text.get() and "失敗" in app.status_text.get(), app.status_text.get())

    app.tree.selection_set(leaf("slides.pptx"))
    pump(0.3)
    check("preview shows Markdown", "MarkItDownSmoke PowerPoint" in preview_text(), preview_text()[:200])
    check("copy enabled for a result", str(app.copy_button.cget("state")) == "normal")
    app.copy_markdown()
    check("copy puts Markdown on the clipboard", "MarkItDownSmoke PowerPoint" in root.clipboard_get())
    shot("3_preview")

    app.tree.selection_set(leaf("missing.pdf"))
    pump(0.3)
    check("failed item explains the problem", "找不到這個檔案" in preview_text(), preview_text()[:200])
    check("copy disabled for a failure", str(app.copy_button.cget("state")) == "disabled")
    shot("4_error")

    # Re-convert everything and stop straight away: the run ends early and cleanly.
    app.start(only=app._leaf_iids())
    app.stop()
    wait(lambda: not app.busy, "the stopped run to end")
    check("stop ends the run", "已停止" in app.status_text.get(), app.status_text.get())
    check("stopped run leaves items pending", any(app.items[i].status == "pending" for i in app._leaf_iids()))

    # Quick Convert panel: drop files in, get draggable result cards out.
    panel = app.show_panel()
    pump(0.3)
    check("panel opens with a large well", panel.well.winfo_height() > 150 and not panel.holder.winfo_ismapped())
    panel.add([str(samples / "sheet.xlsx"), str(samples / "page.html"), str(missing)])
    wait(lambda: not panel.busy, "the panel to convert")
    wait(lambda: not panel.animating, "the panel's animations to settle")
    pump(0.1)
    names = [c.result.output.name if c.result.ok else "error" for c in panel.cards]
    check("panel shows newest result first", names == ["error", "page.md", "sheet.md"], str(names))
    check("well shrinks to one row once there are results", panel.well.winfo_height() < 100 and panel.holder.winfo_ismapped())
    sheet_card = panel.cards[-1]
    check("Excel card carries a token estimate", sheet_card.result.tokens > 0 and sheet_card.result.excel)

    class FakeDnd:
        COPY, DND_FILES = "copy", "DND_Files"

    panel.dnd = FakeDnd
    drag = sheet_card._drag_init(None)
    check("dragging a card hands over its .md file", drag[1] == "DND_Files" and drag[2][0].endswith("sheet.md"), str(drag))
    panel.dnd = app.dnd
    sheet_card.copy()
    check("copy puts the card's Markdown on the clipboard", "MarkItDownSmoke" in root.clipboard_get())
    shot("5_panel")
    # Waterfall workbooks: the card offers "only the newest versions", one click makes that file.
    from compact_test import build_waterfall

    import panel as panel_module

    waterfall = samples / "waterfall.xlsx"
    build_waterfall(waterfall)
    panel_module.OFFER_MIN_TOKENS = 0  # the fixture is small; real workbooks need over 30k tokens
    panel.add([str(waterfall)])
    wait(lambda: not panel.busy and not panel.animating, "the waterfall to convert")
    offers = panel.cards[0].result.offers
    check("waterfall card offers the newest one and two versions", sorted(offers) == [1, 2] and offers[1] < offers[2],
          str(offers))
    panel.convert_latest(str(waterfall), 1)
    wait(lambda: not panel.busy and not panel.animating, "the newest-version extract")
    extract = panel.cards[0].result
    check("newest-version extract is its own small file",
          extract.output.name == "waterfall_最新1版.md" and extract.variant == "只含最新 1 版"
          and extract.tokens < panel.cards[1].result.tokens, f"{extract.output} {extract.tokens}")
    panel.copy_instructions()
    check("Claude Project instructions on the clipboard", "00_索引與邏輯.md" in root.clipboard_get())

    panel.clear()
    wait(lambda: not panel.animating, "the well to open up again")
    pump(0.1)
    check("clearing brings the large well back", not panel.holder.winfo_ismapped() and panel.well.winfo_height() > 150)

    # Motion: easing curves end where they should, colours blend, tweens retarget from where they are.
    from motion import EASE_IN_OUT, EASE_OUT, Motion, blend

    samples = [i / 100 for i in range(101)]
    check("easing curves run 0 -> 1 without turning back",
          all(f(0) == 0 and f(1) == 1 and all(f(b) >= f(a) for a, b in zip(samples, samples[1:]))
              for f in (EASE_OUT, EASE_IN_OUT)))
    check("ease-out is front-loaded, ease-in-out starts slowly", EASE_OUT(0.25) > 0.7 and EASE_IN_OUT(0.25) < 0.1)
    check("colour blend stands in for opacity", blend("#000000", "#FFFFFF", 0.5) == "#808080")
    m = Motion(root)
    m.set("x", 0.0)
    m.to("x", 1.0, 200, EASE_OUT, lambda: None)
    pump(0.08)
    midway = m.get("x")
    m.to("x", 0.0, 100, EASE_OUT, lambda: None)
    pump(0.02)
    check("an interrupted tween restarts from its current value", 0 < midway < 1 and 0 < m.get("x") <= midway,
          f"{midway} -> {m.get('x')}")
    wait(lambda: not m.active, "the tween to finish")
    check("a retargeted tween lands on the new target", m.get("x") == 0.0)
    panel.close()
    pump(0.2)
    check("closing the panel keeps the main window", panel.win.state() == "withdrawn" and root.winfo_exists())

    app.tree.selection_set(app._leaf_iids())
    app.remove_selected()
    pump(0.3)
    check("removing everything brings back the empty state", bool(app.drop_zone.winfo_ismapped()))
    root.destroy()
    return failures


def test_exe(exe: str, shots: Path | None, wait: float = 25) -> list:
    """The packaged exe starts and stays up, as the main window and as the panel alone."""
    failures = []
    for mode in ("--gui", "--panel"):
        proc = subprocess.Popen([str(Path(exe).resolve()), mode])
        try:
            time.sleep(wait)
            if proc.poll() is not None:
                failures.append(f"{exe} {mode} exited early with code {proc.returncode}")
                continue
            print(f"PASS  {exe} {mode} is still running after {wait:.0f}s")
            if shots is not None:
                grab(shots / f"exe{mode.replace('--', '_')}.png")
        finally:
            if proc.poll() is None:
                proc.kill()
                proc.wait()
    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--exe", help="also launch this packaged exe with --gui")
    parser.add_argument("--screenshots", help="folder to save screenshots in")
    parser.add_argument("--appearance", choices=["light", "dark"], help="force an appearance (default: follow the system)")
    args = parser.parse_args()
    shots = Path(args.screenshots) if args.screenshots else None
    if shots:
        shots.mkdir(parents=True, exist_ok=True)
    if args.appearance:
        os.environ["MARKITDOWN_APPEARANCE"] = args.appearance

    failures = test_window(shots, prefix=f"{args.appearance or 'system'}_")
    if args.exe:
        failures += test_exe(args.exe, shots)

    print()
    if failures:
        print(f"{len(failures)} GUI check(s) failed: {', '.join(failures)}")
        return 1
    print("GUI smoke test passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
