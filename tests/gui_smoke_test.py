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

    app.tree.selection_set(app._leaf_iids())
    app.remove_selected()
    pump(0.3)
    check("removing everything brings back the empty state", bool(app.drop_zone.winfo_ismapped()))
    root.destroy()
    return failures


def test_exe(exe: str, shots: Path | None, wait: float = 25) -> list:
    proc = subprocess.Popen([str(Path(exe).resolve()), "--gui"])
    try:
        time.sleep(wait)
        if proc.poll() is not None:
            return [f"{exe} --gui exited early with code {proc.returncode}"]
        print(f"PASS  {exe} --gui is still running after {wait:.0f}s")
        if shots is not None:
            grab(shots / "exe_gui.png")
        return []
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait()


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
