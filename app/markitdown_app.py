#!/usr/bin/env python3
"""MarkItDown desktop app: the entry point of markitdown.exe.

* Double-click the exe            -> graphical converter (GUI)
* Drag files/folders onto the exe -> converts them to .md next to the originals
* Run it from a terminal          -> full CLI (see `markitdown.exe --help`)

The conversion logic lives in skills/markitdown/scripts/convert.py so the
Claude skill and the exe always behave the same.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "skills" / "markitdown" / "scripts"))

import convert  # noqa: E402  (path set up above; bundled directly by PyInstaller)

APP_TITLE = "MarkItDown 轉換器"


# --------------------------------------------------------------------------- #
# Windows console helpers
# --------------------------------------------------------------------------- #
def _is_onefile_build() -> bool:
    meipass = getattr(sys, "_MEIPASS", None)
    if not getattr(sys, "frozen", False) or not meipass:
        return False
    return Path(meipass).resolve().parent != Path(sys.executable).resolve().parent


def launched_from_explorer() -> bool:
    """True when Windows created a console just for us (double-click / drag & drop).

    From cmd/PowerShell the shell is attached to the same console too, so the
    process count is higher. A PyInstaller one-file build runs as a bootloader
    plus a child process, both attached.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        buf = (ctypes.c_uint32 * 16)()
        count = ctypes.windll.kernel32.GetConsoleProcessList(buf, 16)
    except Exception:
        return False
    own_processes = 2 if _is_onefile_build() else 1
    return 0 < count <= own_processes


def hide_console() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass


def enable_high_dpi() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Drag & drop onto the exe icon
# --------------------------------------------------------------------------- #
def run_dropped(paths: list) -> int:
    print(f"{APP_TITLE}  ({convert.versions_text()})\n")
    sources = convert.collect_sources(paths, recursive=True)
    if not sources:
        print("找不到可轉換的檔案。")
        input("\n按 Enter 鍵關閉視窗…")
        return 1
    print("正在載入轉換引擎…")
    converter = convert.build_converter()

    def report(index: int, total: int, result: convert.Result) -> None:
        if result.error:
            print(f"({index}/{total}) [失敗] {result.source.value}\n        {result.error}")
        else:
            print(f"({index}/{total}) [完成] {result.source.value}\n        -> {result.output}")

    results = convert.convert_batch(converter, sources, None, on_result=report)
    failed = sum(1 for r in results if not r.ok)
    print(f"\n完成：成功 {len(results) - failed} 個，失敗 {failed} 個。")
    input("\n按 Enter 鍵關閉視窗…")
    return 1 if failed else 0


# --------------------------------------------------------------------------- #
# GUI
# --------------------------------------------------------------------------- #
def run_gui(initial_inputs=()) -> int:
    enable_high_dpi()
    root, _app = create_window(initial_inputs)
    root.mainloop()
    return 0


def create_window(initial_inputs=()):
    import gui

    return gui.create_window(initial_inputs)


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main() -> int:
    args = sys.argv[1:]
    from_explorer = launched_from_explorer()

    if args and args[0] == "--gui":
        if from_explorer:
            hide_console()
        return run_gui(args[1:])

    if not args:
        stdin_is_pipe = sys.stdin is not None and not sys.stdin.isatty()
        if stdin_is_pipe and not from_explorer:
            return convert.main([])  # e.g. `type file.pdf | markitdown.exe`
        if from_explorer:
            hide_console()
        else:
            print("正在開啟圖形介面… (命令列用法請執行: markitdown --help)")
        return run_gui()

    if from_explorer and all(not a.startswith("-") and os.path.exists(a) for a in args):
        return run_dropped(args)

    return convert.main(args)


if __name__ == "__main__":
    sys.exit(main())
