#!/usr/bin/env python3
"""Smoke test for the GUI of markitdown.exe.

1. Opens the converter window from source, queues sample files, presses
   "開始轉換" and checks the .md files that come out.
2. With --exe, also launches the packaged `markitdown.exe --gui` and checks
   that it stays up (i.e. did not crash on start-up).

    python tests/gui_smoke_test.py --screenshots shots
    python tests/gui_smoke_test.py --exe dist/markitdown.exe --screenshots shots
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app"))
sys.path.insert(0, str(ROOT / "tests"))

import markitdown_app  # noqa: E402
import smoke_test  # noqa: E402


def grab(path: Path, bbox=None) -> None:
    try:
        from PIL import ImageGrab

        ImageGrab.grab(bbox=bbox).save(path)
        print(f"screenshot: {path}")
    except Exception as exc:  # screenshots are a nice-to-have
        print(f"(screenshot skipped: {exc})")


def test_window(shots: Path | None, timeout: float = 240) -> list:
    failures = []
    tmp = Path(tempfile.mkdtemp())
    samples = tmp / "samples"
    smoke_test.make_samples(samples)
    out = tmp / "out"

    root, app = markitdown_app.create_window([str(samples), str(samples / "doc.pdf")])
    app.dest_mode.set("folder")
    app.out_dir_var.set(str(out))
    deadline = time.time() + timeout
    state = {"started": False}

    def shot(name: str) -> None:
        if shots is None:
            return
        root.update()
        time.sleep(0.3)
        root.update()
        x, y = root.winfo_rootx(), root.winfo_rooty()
        grab(shots / f"{name}.png", (x, y, x + root.winfo_width(), y + root.winfo_height()))

    def tick() -> None:
        if time.time() > deadline:
            failures.append("GUI conversion timed out")
            root.destroy()
            return
        if not state["started"]:
            if app.converter is not None:
                shot("gui_1_ready")
                app._start()
                state["started"] = True
        elif not app.busy:
            shot("gui_2_done")
            root.destroy()
            return
        root.after(200, tick)

    root.after(200, tick)
    root.mainloop()

    produced = sorted(p.relative_to(out).as_posix() for p in out.rglob("*.md"))
    want = {"doc.md", "doc.pdf.md", "page.md", "sheet.md", "slides.md", "bundle.md", "sub/nested.md"}
    print("GUI produced:", produced)
    if not want.issubset(produced):
        failures.append(f"GUI outputs missing: {sorted(want - set(produced))}")
    if "doc (2).md" in produced:
        failures.append("GUI converted a duplicate input twice")
    if len(app.results) != len(produced):
        failures.append(f"result list has {len(app.results)} entries, expected {len(produced)}")
    return failures


def test_exe(exe: str, shots: Path | None, wait: float = 25) -> list:
    proc = subprocess.Popen([str(Path(exe).resolve()), "--gui"])
    try:
        time.sleep(wait)
        if proc.poll() is not None:
            return [f"{exe} --gui exited early with code {proc.returncode}"]
        print(f"{exe} --gui is still running after {wait:.0f}s")
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
    args = parser.parse_args()
    shots = Path(args.screenshots) if args.screenshots else None
    if shots:
        shots.mkdir(parents=True, exist_ok=True)

    failures = test_window(shots)
    if args.exe:
        failures += test_exe(args.exe, shots)

    if failures:
        for f in failures:
            print("FAIL ", f)
        return 1
    print("GUI smoke test passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
