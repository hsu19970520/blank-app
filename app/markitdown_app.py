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
import queue
import subprocess
import sys
import threading
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "skills" / "markitdown" / "scripts"))

import convert  # noqa: E402  (path set up above; bundled directly by PyInstaller)

APP_TITLE = "MarkItDown 轉換器"
DEFAULT_URL_DIR = Path.home() / "Downloads"


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


def open_in_file_manager(path: Path, select: bool = False) -> None:
    try:
        if sys.platform == "win32":
            if select and path.is_file():
                subprocess.Popen(["explorer", "/select,", str(path)])
            else:
                os.startfile(str(path if path.is_dir() else path.parent))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", str(path)] if select else ["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path if path.is_dir() else path.parent)])
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
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    DND_FILES, root = None, None
    if sys.platform == "win32":  # optional: drag files straight into the window
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD

            root = TkinterDnD.Tk()
        except Exception:
            DND_FILES = None
    if root is None:
        root = tk.Tk()

    app = ConverterWindow(root, tk, ttk, filedialog, messagebox, DND_FILES)
    for item in initial_inputs:
        app.add_input(item)
    return root, app


class ConverterWindow:
    def __init__(self, root, tk, ttk, filedialog, messagebox, dnd_files):
        self.root = root
        self.tk = tk
        self.ttk = ttk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.converter = None
        self.converter_lock = threading.Lock()
        self.busy = False
        self.results: list = []

        root.title(APP_TITLE)
        scale = max(1.0, root.winfo_fpixels("1i") / 96.0)  # high-DPI screens
        root.geometry(f"{int(960 * scale)}x{int(720 * scale)}")
        root.minsize(int(760 * scale), int(580 * scale))
        self._build(dnd_files)
        self._log(f"{convert.versions_text()}")
        self._log("加入檔案、資料夾或網址後，按「開始轉換」。")
        # Load the (slow) conversion engine in the background right away.
        threading.Thread(target=self._get_converter, daemon=True).start()
        root.after(100, self._poll_events)

    # ----- layout ---------------------------------------------------------- #
    def _build(self, dnd_files) -> None:
        tk, ttk = self.tk, self.ttk
        pad = {"padx": 8, "pady": 4}
        main = ttk.Frame(self.root, padding=8)
        main.pack(fill="both", expand=True)

        # Inputs
        inputs = ttk.LabelFrame(main, text="1. 要轉換的檔案 / 資料夾 / 網址", padding=6)
        inputs.pack(fill="both", expand=False, **pad)
        bar = ttk.Frame(inputs)
        bar.pack(fill="x")
        ttk.Button(bar, text="加入檔案…", command=self._choose_files).pack(side="left")
        ttk.Button(bar, text="加入資料夾…", command=self._choose_folder).pack(side="left", padx=4)
        ttk.Button(bar, text="移除選取", command=self._remove_selected).pack(side="left")
        ttk.Button(bar, text="清空", command=self._clear_inputs).pack(side="left", padx=4)

        url_row = ttk.Frame(inputs)
        url_row.pack(fill="x", pady=(6, 4))
        ttk.Label(url_row, text="網址：").pack(side="left")
        self.url_var = tk.StringVar()
        url_entry = ttk.Entry(url_row, textvariable=self.url_var)
        url_entry.pack(side="left", fill="x", expand=True)
        url_entry.bind("<Return>", lambda _e: self._add_url())
        ttk.Button(url_row, text="加入網址", command=self._add_url).pack(side="left", padx=4)

        list_frame = ttk.Frame(inputs)
        list_frame.pack(fill="both", expand=True)
        self.input_list = tk.Listbox(list_frame, height=7, selectmode="extended", activestyle="none")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.input_list.yview)
        self.input_list.configure(yscrollcommand=scroll.set)
        self.input_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        if dnd_files:
            self.input_list.drop_target_register(dnd_files)
            self.input_list.dnd_bind("<<Drop>>", self._on_drop)
            self.root.drop_target_register(dnd_files)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        hint = (
            "提示：可以直接把檔案或資料夾拖曳到這個視窗。"
            if dnd_files
            else "提示：把檔案拖曳到 markitdown.exe 圖示上，會直接在原位置產生 .md。"
        )
        ttk.Label(inputs, text=hint, foreground="#666").pack(anchor="w", pady=(4, 0))

        # Output options
        options = ttk.LabelFrame(main, text="2. 輸出設定", padding=6)
        options.pack(fill="x", **pad)
        self.dest_mode = tk.StringVar(value="beside")
        row1 = ttk.Frame(options)
        row1.pack(fill="x")
        ttk.Radiobutton(row1, text="存在原檔案旁邊（網址存到「下載」資料夾）", variable=self.dest_mode, value="beside").pack(side="left")
        row2 = ttk.Frame(options)
        row2.pack(fill="x", pady=2)
        ttk.Radiobutton(row2, text="全部存到：", variable=self.dest_mode, value="folder").pack(side="left")
        self.out_dir_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.out_dir_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row2, text="瀏覽…", command=self._choose_out_dir).pack(side="left", padx=4)
        row3 = ttk.Frame(options)
        row3.pack(fill="x", pady=(4, 0))
        self.recursive_var = tk.BooleanVar(value=True)
        self.skip_var = tk.BooleanVar(value=False)
        self.keep_uris_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="資料夾包含子資料夾", variable=self.recursive_var).pack(side="left")
        ttk.Checkbutton(row3, text="略過已存在的 .md", variable=self.skip_var).pack(side="left", padx=12)
        ttk.Checkbutton(row3, text="保留內嵌圖片 (base64)", variable=self.keep_uris_var).pack(side="left")

        # Run
        run_row = ttk.Frame(main)
        run_row.pack(fill="x", **pad)
        self.run_button = ttk.Button(run_row, text="3. 開始轉換", command=self._start)
        self.run_button.pack(side="left")
        self.progress = ttk.Progressbar(run_row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=8)
        self.status_var = tk.StringVar(value="正在載入轉換引擎…")
        ttk.Label(run_row, textvariable=self.status_var, width=28).pack(side="left")

        # Results / log
        notebook = ttk.Notebook(main)
        notebook.pack(fill="both", expand=True, **pad)
        self.notebook = notebook

        preview_tab = ttk.Frame(notebook, padding=4)
        notebook.add(preview_tab, text="結果預覽")
        paned = ttk.PanedWindow(preview_tab, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left = ttk.Frame(paned)
        self.result_list = tk.Listbox(left, width=34, activestyle="none", exportselection=False)
        self.result_list.pack(fill="both", expand=True)
        self.result_list.bind("<<ListboxSelect>>", lambda _e: self._show_selected_result())
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=(4, 0))
        ttk.Button(btns, text="複製 Markdown", command=self._copy_selected).pack(side="left")
        ttk.Button(btns, text="開啟位置", command=self._open_selected).pack(side="left", padx=4)
        paned.add(left, weight=1)
        right = ttk.Frame(paned)
        self.preview = tk.Text(right, wrap="word", undo=False)
        pscroll = ttk.Scrollbar(right, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=pscroll.set, state="disabled")
        self.preview.pack(side="left", fill="both", expand=True)
        pscroll.pack(side="left", fill="y")
        paned.add(right, weight=3)

        log_tab = ttk.Frame(notebook, padding=4)
        notebook.add(log_tab, text="轉換紀錄")
        self.log_text = tk.Text(log_tab, wrap="word", height=8)
        lscroll = ttk.Scrollbar(log_tab, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lscroll.set, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)
        lscroll.pack(side="left", fill="y")

    # ----- inputs ---------------------------------------------------------- #
    def add_input(self, value: str) -> None:
        value = value.strip().strip('"')
        if value and value not in self.input_list.get(0, "end"):
            self.input_list.insert("end", value)

    def _choose_files(self) -> None:
        types = [
            ("支援的檔案", " ".join(f"*{ext}" for ext in sorted(convert.SUPPORTED_EXTENSIONS))),
            ("所有檔案", "*.*"),
        ]
        for path in self.filedialog.askopenfilenames(title="選擇要轉換的檔案", filetypes=types):
            self.add_input(path)

    def _choose_folder(self) -> None:
        folder = self.filedialog.askdirectory(title="選擇資料夾")
        if folder:
            self.add_input(folder)

    def _choose_out_dir(self) -> None:
        folder = self.filedialog.askdirectory(title="選擇輸出資料夾")
        if folder:
            self.out_dir_var.set(folder)
            self.dest_mode.set("folder")

    def _add_url(self) -> None:
        url = self.url_var.get().strip()
        if not url:
            return
        if not convert.is_url(url):
            url = "https://" + url
        self.add_input(url)
        self.url_var.set("")

    def _remove_selected(self) -> None:
        for index in reversed(self.input_list.curselection()):
            self.input_list.delete(index)

    def _clear_inputs(self) -> None:
        self.input_list.delete(0, "end")

    def _on_drop(self, event) -> None:
        for path in self.root.tk.splitlist(event.data):
            self.add_input(path)

    # ----- conversion ------------------------------------------------------ #
    def _get_converter(self):
        with self.converter_lock:
            if self.converter is None:
                try:
                    self.converter = convert.build_converter()
                    self.events.put(("ready",))
                except Exception as exc:
                    self.events.put(("log", f"載入轉換引擎失敗：{convert.describe_error(exc)}"))
                    raise
            return self.converter

    def _start(self) -> None:
        if self.busy:
            return
        inputs = list(self.input_list.get(0, "end"))
        if not inputs:
            self.messagebox.showinfo(APP_TITLE, "請先加入要轉換的檔案、資料夾或網址。")
            return
        out_dir = None
        if self.dest_mode.get() == "folder":
            text = self.out_dir_var.get().strip()
            if not text:
                self.messagebox.showwarning(APP_TITLE, "請選擇輸出資料夾。")
                return
            out_dir = Path(text)
        self.busy = True
        self.run_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=1)
        self.notebook.select(1)
        options = (inputs, out_dir, self.recursive_var.get(), self.skip_var.get(), self.keep_uris_var.get())
        threading.Thread(target=self._worker, args=options, daemon=True).start()

    def _worker(self, inputs, out_dir, recursive, skip_existing, keep_data_uris) -> None:
        try:
            self.events.put(("status", "正在載入轉換引擎…"))
            converter = self._get_converter()
            file_inputs = [i for i in inputs if not convert.is_url(i)]
            url_inputs = [i for i in inputs if convert.is_url(i)]
            sources = convert.collect_sources(file_inputs, recursive=recursive)
            url_sources = convert.collect_sources(url_inputs)
            total = len(sources) + len(url_sources)
            if total == 0:
                self.events.put(("log", "找不到可轉換的檔案（資料夾中沒有支援的格式）。"))
                return
            self.events.put(("start", total))
            done = [0]

            def on_result(_i, _t, result) -> None:
                done[0] += 1
                self.events.put(("result", done[0], total, result))

            if sources:
                convert.convert_batch(
                    converter, sources, out_dir,
                    keep_data_uris=keep_data_uris, skip_existing=skip_existing, on_result=on_result,
                )
            if url_sources:
                url_dir = out_dir or DEFAULT_URL_DIR
                convert.convert_batch(
                    converter, url_sources, url_dir,
                    keep_data_uris=keep_data_uris, skip_existing=skip_existing, on_result=on_result,
                )
        except Exception as exc:
            self.events.put(("log", f"錯誤：{convert.describe_error(exc)}"))
        finally:
            self.events.put(("done",))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "ready":
            if not self.busy:
                self.status_var.set("就緒")
        elif kind == "status":
            self.status_var.set(event[1])
        elif kind == "log":
            self._log(event[1])
        elif kind == "start":
            total = event[1]
            self.progress.configure(maximum=total, value=0)
            self.status_var.set(f"轉換中 0/{total}")
            self._log(f"開始轉換 {total} 個項目…")
        elif kind == "result":
            index, total, result = event[1:]
            self.progress.configure(value=index)
            self.status_var.set(f"轉換中 {index}/{total}")
            if result.error:
                self._log(f"[失敗] {result.source.value}\n       {result.error}")
            elif result.markdown is None:
                self._log(f"[略過] {result.source.value}（已存在 {result.output}）")
            else:
                self._log(f"[完成] {result.source.value}\n       -> {result.output}")
                self.results.append(result)
                self.result_list.insert("end", Path(str(result.output)).name)
        elif kind == "done":
            self.busy = False
            self.run_button.configure(state="normal")
            ok = sum(1 for r in self.results if r.ok)
            self.status_var.set("完成")
            self._log("全部處理完畢。")
            if ok:
                self.notebook.select(0)
                self.result_list.selection_clear(0, "end")
                self.result_list.selection_set("end")
                self.result_list.see("end")
                self._show_selected_result()

    # ----- results --------------------------------------------------------- #
    def _selected_result(self):
        selection = self.result_list.curselection()
        return self.results[selection[0]] if selection else None

    def _show_selected_result(self) -> None:
        result = self._selected_result()
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        if result is not None:
            self.preview.insert("1.0", result.markdown or "")
        self.preview.configure(state="disabled")

    def _copy_selected(self) -> None:
        result = self._selected_result()
        if result is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(result.markdown or "")
        self.status_var.set("已複製到剪貼簿")

    def _open_selected(self) -> None:
        result = self._selected_result()
        if result is not None and result.output:
            open_in_file_manager(Path(result.output), select=True)

    def _log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


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
