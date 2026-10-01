"""Converter window for markitdown.exe.

Built on Tkinter with the Sun Valley theme so it looks and behaves like a
Windows 11 app, follows the system light/dark setting, and keeps one job in
focus: drop documents in, read and save the Markdown that comes out.
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import urlparse

import convert

APP_NAME = "MarkItDown 轉換器"
DEFAULT_URL_DIR = Path.home() / "Downloads"
PREVIEW_LIMIT = 200_000  # characters shown in the preview; the saved file is always complete

# Colour tokens. Text contrast against `surface` / `bg` (WCAG):
#   light: text 17.0:1, secondary 6.6:1, syntax 5.1:1, success 5.4:1, danger 5.7:1, caution 5.3:1
#   dark:  text 15.0:1, secondary 6.9:1, syntax 5.6:1, success 7.4:1, danger 7.4:1, caution 11.4:1
PALETTES = {
    "light": {
        "bg": "#FAFAFA", "surface": "#FFFFFF", "border": "#E5E5E5", "text": "#1C1C1C",
        "secondary": "#5D5D5D", "syntax": "#6E6E6E", "code_bg": "#F3F3F3", "accent": "#005FB8",
        "drop": "#EAF2FB", "success": "#0F7B0F", "danger": "#C42B1C", "caution": "#9D5D00",
    },
    "dark": {
        "bg": "#1C1C1C", "surface": "#232323", "border": "#333333", "text": "#FAFAFA",
        "secondary": "#ABABAB", "syntax": "#9A9A9A", "code_bg": "#2D2D2D", "accent": "#57C8FF",
        "drop": "#1E3140", "success": "#6CCB5F", "danger": "#FF99A4", "caution": "#FCE100",
    },
}

TYPE_LABELS = {
    ".pdf": "PDF", ".docx": "Word", ".pptx": "PowerPoint", ".xlsx": "Excel", ".xls": "Excel",
    ".html": "HTML", ".htm": "HTML", ".csv": "CSV", ".json": "JSON", ".jsonl": "JSON",
    ".xml": "XML", ".rss": "RSS", ".atom": "RSS", ".epub": "EPUB", ".ipynb": "Notebook",
    ".msg": "Outlook", ".zip": "ZIP", ".jpg": "圖片", ".jpeg": "圖片", ".png": "圖片",
    ".wav": "音訊", ".mp3": "音訊", ".m4a": "音訊", ".mp4": "影片", ".txt": "文字", ".md": "Markdown",
}

STATUS_TEXT = {
    "pending": "等待轉換",
    "working": "轉換中…",
    "done": "✓ 已完成",
    "empty": "⚠ 沒有文字",
    "skipped": "✓ 已有 .md",
    "failed": "✕ 失敗",
}
LEAF_KINDS = ("file", "url")


# --------------------------------------------------------------------------- #
# Platform helpers
# --------------------------------------------------------------------------- #
def asset_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "assets" / name


def system_appearance() -> str:
    """'dark' or 'light', following the OS setting (no app-specific switch)."""
    override = os.environ.get("MARKITDOWN_APPEARANCE", "").lower()
    if override in PALETTES:
        return override
    if sys.platform == "win32":
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value else "dark"
        except OSError:
            return "light"
    if sys.platform == "darwin":
        try:
            out = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"], capture_output=True, text=True, timeout=2
            )
            return "dark" if "Dark" in out.stdout else "light"
        except Exception:
            return "light"
    return "light"


def set_title_bar_dark(root, dark: bool) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        for attribute in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (new / pre-20H1 builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value), 4) == 0:
                break
    except Exception:
        pass


def open_path(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def reveal_path(path: Path) -> None:
    try:
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", str(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path.parent)])
    except Exception:
        pass


def friendly_error(error_type: Optional[str]) -> tuple:
    """(what went wrong, how to fix it) in plain language."""
    network = {"ConnectionError", "Timeout", "ConnectTimeout", "ReadTimeout", "HTTPError", "SSLError", "TooManyRedirects"}
    if error_type == "FileNotFoundError":
        return "找不到這個檔案", "檔案可能已被移動、重新命名或刪除。請確認位置後重新加入。"
    if error_type == "PermissionError":
        return "無法儲存 .md 檔", "請關閉正在使用這個檔案的程式，或在「儲存到」改選其他資料夾。"
    if error_type == "UnsupportedFormatException":
        return "不支援這種檔案格式", "請改用 PDF、Word、PowerPoint、Excel、HTML 等格式。完整清單請見「說明 › 支援的格式」。"
    if error_type == "MissingDependencyException":
        return "缺少轉換這種格式所需的元件", "這種格式需要額外安裝的元件，目前的版本無法轉換。"
    if error_type in network:
        return "無法取得網頁內容", "請確認網路連線和網址是否正確，然後再轉換一次。"
    if error_type == "FileConversionException":
        return "無法讀取這個檔案", "檔案可能已損毀、受密碼保護，或內容不完整。請確認它能在原本的程式中開啟。"
    return "轉換失敗", "請確認檔案能在原本的程式中正常開啟，然後再試一次。"


@dataclass
class Item:
    kind: str  # "file", "url" or "folder"
    value: str  # path or URL
    scan_root: Optional[Path] = None  # folder a file was found in (mirrors sub-folders in "save to")
    status: str = "pending"
    output: Optional[Path] = None
    markdown: Optional[str] = None
    error: Optional[str] = None
    error_type: Optional[str] = None

    @property
    def key(self) -> str:
        return self.value if self.kind == "url" else os.path.normcase(str(Path(self.value).resolve()))

    @property
    def display_name(self) -> str:
        if self.kind == "url":
            parsed = urlparse(self.value)
            return (parsed.netloc + parsed.path).rstrip("/") or self.value
        return Path(self.value).name or self.value

    @property
    def type_label(self) -> str:
        if self.kind == "folder":
            return "資料夾"
        if self.kind == "url":
            return "YouTube" if "youtube.com" in self.value or "youtu.be" in self.value else "網頁"
        return TYPE_LABELS.get(Path(self.value).suffix.lower(), Path(self.value).suffix.lstrip(".").upper() or "檔案")


# --------------------------------------------------------------------------- #
# Window
# --------------------------------------------------------------------------- #
def create_window(initial_inputs=()):
    import tkinter as tk

    root = None
    dnd = None
    if sys.platform == "win32":  # native drag & drop into the window
        try:
            import tkinterdnd2

            root = tkinterdnd2.TkinterDnD.Tk()
            dnd = tkinterdnd2
        except Exception:
            root = None
    if root is None:
        root = tk.Tk()
    app = ConverterWindow(root, dnd)
    app.add_inputs(initial_inputs)
    return root, app


class ConverterWindow:
    def __init__(self, root, dnd=None):
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
        from tkinter import font as tkfont

        self.root, self.tk, self.ttk = root, tk, ttk
        self.filedialog, self.messagebox, self.tkfont = filedialog, messagebox, tkfont
        self.dnd = dnd

        self.items: Dict[str, Item] = {}
        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.converter = None
        self.converter_error: Optional[str] = None
        self.converter_lock = threading.Lock()
        self.stop_event = threading.Event()
        self.busy = False
        self.dragging = False
        self.log_lines: List[str] = []
        self.log_window = None
        self.run_summary: Dict[str, object] = {}
        self._repaint_pending = False

        self.output_mode = tk.StringVar(value="beside")
        self.output_dir = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.skip_existing = tk.BooleanVar(value=False)
        self.keep_data_uris = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar()
        self.url_text = tk.StringVar()

        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        self.appearance = system_appearance()
        self._load_fonts()
        self._configure_styles()
        self._setup_window()
        self._build_menu()
        self._build_layout()
        self._bind_keys()
        self._apply_palette()
        # Sun Valley resets plain Tk widget colours (tk_setPalette) on <<ThemeChanged>>;
        # repaint just those afterwards. Restyling ttk here would fire the event again.
        self.root.bind("<<ThemeChanged>>", lambda _e: self._schedule_repaint(), add="+")
        self._refresh()

        self._set_status("正在準備轉換引擎，第一次啟動約需數秒…")
        self._log(convert.versions_text())
        threading.Thread(target=self._load_converter, daemon=True).start()
        root.after(100, self._poll_events)
        root.after(2000, self._watch_appearance)

    # ----- theme ----------------------------------------------------------- #
    def _load_fonts(self) -> None:
        try:
            import sv_ttk

            sv_ttk.set_theme(self.appearance, self.root)
            self.sv_ttk = sv_ttk
        except Exception:
            self.sv_ttk = None
        families = set(self.tkfont.families(self.root))
        cjk_ui = next((f for f in ("Microsoft JhengHei UI", "Microsoft JhengHei") if f in families), None)
        mono = next((f for f in ("Cascadia Mono", "Consolas", "DejaVu Sans Mono", "Courier New") if f in families), "TkFixedFont")

        def px(size: int) -> int:
            return -round(size * self.scale)

        # Sun Valley sizes fonts in pixels; scale them for high-DPI screens and
        # use the Traditional Chinese UI face when Windows has it.
        sizes = {
            "SunValleyCaptionFont": (12, "normal"), "SunValleyBodyFont": (14, "normal"),
            "SunValleyBodyStrongFont": (14, "bold"), "SunValleyBodyLargeFont": (18, "normal"),
            "SunValleySubtitleFont": (20, "bold"), "SunValleyTitleFont": (28, "bold"),
        }
        for name, (size, weight) in sizes.items():
            try:
                font = self.tkfont.nametofont(name, root=self.root)
            except Exception:
                font = self.tkfont.Font(root=self.root, name=name, exists=False)
            options = {"size": px(size)}
            if cjk_ui:
                options.update(family=cjk_ui, weight=weight)
            font.configure(**options)
        body = self.tkfont.nametofont("SunValleyBodyFont", root=self.root)
        family = body.actual("family")
        self.fonts = {
            "body": body,
            "strong": self.tkfont.nametofont("SunValleyBodyStrongFont", root=self.root),
            "caption": self.tkfont.nametofont("SunValleyCaptionFont", root=self.root),
            "title": self.tkfont.Font(root=self.root, family=family, size=px(20), weight="bold"),
            "h1": self.tkfont.Font(root=self.root, family=family, size=px(24), weight="bold"),
            "h2": self.tkfont.Font(root=self.root, family=family, size=px(19), weight="bold"),
            "h3": self.tkfont.Font(root=self.root, family=family, size=px(16), weight="bold"),
            "bold": self.tkfont.Font(root=self.root, family=family, size=px(14), weight="bold"),
            "mono": self.tkfont.Font(root=self.root, family=mono, size=px(13)),
        }

    def _configure_styles(self) -> None:
        """Per-theme style overrides; ttk keeps settings per theme, so re-run after a switch."""
        style = self.ttk.Style(self.root)
        linespace = self.fonts["body"].metrics("linespace")
        style.configure("Treeview", rowheight=max(linespace + round(12 * self.scale), round(30 * self.scale)))
        style.configure("Treeview", font=self.fonts["body"])
        style.configure("Heading", font=self.fonts["caption"])
        style.configure("Secondary.TLabel", font=self.fonts["caption"])
        style.configure("SecondaryBody.TLabel", font=self.fonts["body"])
        style.configure("Title.TLabel", font=self.fonts["strong"])

    def _watch_appearance(self) -> None:
        current = system_appearance()
        if current != self.appearance:
            self.appearance = current
            if self.sv_ttk:
                self.sv_ttk.set_theme(current, self.root)
            self._configure_styles()
            self._apply_palette()
        self.root.after(2000, self._watch_appearance)

    @property
    def colors(self) -> dict:
        return PALETTES[self.appearance]

    def _apply_palette(self) -> None:
        c = self.colors
        style = self.ttk.Style(self.root)
        style.configure("Secondary.TLabel", foreground=c["secondary"])
        style.configure("SecondaryBody.TLabel", foreground=c["secondary"])
        style.configure("Placeholder.TEntry", foreground=c["secondary"])
        self.tree.tag_configure("pending", foreground=c["secondary"])
        self.tree.tag_configure("working", foreground=c["text"])
        self.tree.tag_configure("failed", foreground=c["danger"])
        self.tree.tag_configure("empty", foreground=c["caution"])
        self.tree.tag_configure("folder", foreground=c["text"])
        set_title_bar_dark(self.root, self.appearance == "dark")
        self._paint_tk_widgets()

    def _schedule_repaint(self) -> None:
        # The event reaches every widget; repaint once per burst.
        if not self._repaint_pending:
            self._repaint_pending = True
            self.root.after_idle(self._paint_tk_widgets)

    def _paint_tk_widgets(self) -> None:
        self._repaint_pending = False
        c = self.colors
        self.preview.configure(
            background=c["surface"], foreground=c["text"], insertbackground=c["text"],
            selectbackground="#2F60D8", selectforeground="#FFFFFF", highlightthickness=1,
            highlightbackground=c["border"], highlightcolor=c["border"],
        )
        for tag, color in (("syntax", c["syntax"]), ("muted", c["secondary"]), ("quote", c["secondary"]),
                           ("danger", c["danger"]), ("caution", c["caution"]), ("success", c["success"])):
            self.preview.tag_configure(tag, foreground=color)
        self.preview.tag_configure("code", background=c["code_bg"])
        self.drop_zone.configure(background=c["bg"])
        self._draw_drop_zone()
        if self.log_window is not None:
            self.log_text.configure(background=c["surface"], foreground=c["text"])

    # ----- layout ---------------------------------------------------------- #
    def _setup_window(self) -> None:
        root = self.root
        root.title(APP_NAME)
        width = min(round(1120 * self.scale), int(root.winfo_screenwidth() * 0.9))
        height = min(round(740 * self.scale), int(root.winfo_screenheight() * 0.85))
        root.geometry(f"{width}x{height}")
        root.minsize(round(820 * self.scale), round(560 * self.scale))
        try:
            self._icon = self.tk.PhotoImage(file=str(asset_path("icon.png")))
            root.iconphoto(True, self._icon)
            if sys.platform == "win32":
                root.iconbitmap(default=str(asset_path("icon.ico")))
        except Exception:
            pass
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_menu(self) -> None:
        tk = self.tk
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="加入檔案…", accelerator="Ctrl+O", command=self.choose_files)
        file_menu.add_command(label="加入資料夾…", accelerator="Ctrl+Shift+O", command=self.choose_folder)
        file_menu.add_command(label="加入網址", accelerator="Ctrl+L", command=self._focus_url)
        file_menu.add_separator()
        file_menu.add_command(label="轉換", accelerator="Ctrl+Enter", command=self.start)
        file_menu.add_command(label="停止", accelerator="Esc", command=self.stop)
        file_menu.add_separator()
        file_menu.add_command(label="結束", accelerator="Alt+F4", command=self._on_close)
        menubar.add_cascade(label="檔案(F)", underline=3, menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=False)
        edit_menu.add_command(label="複製 Markdown", accelerator="Ctrl+Shift+C", command=self.copy_markdown)
        edit_menu.add_separator()
        edit_menu.add_command(label="全選", accelerator="Ctrl+A", command=self._select_all)
        edit_menu.add_command(label="移除", accelerator="Delete", command=self.remove_selected)
        edit_menu.add_command(label="清除清單", command=self.clear_items)
        menubar.add_cascade(label="編輯(E)", underline=3, menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=False)
        view_menu.add_command(label="轉換紀錄", command=self.show_log)
        menubar.add_cascade(label="檢視(V)", underline=3, menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="支援的格式", command=self._show_formats)
        help_menu.add_command(label=f"關於 {APP_NAME}", command=self._show_about)
        menubar.add_cascade(label="說明(H)", underline=3, menu=help_menu)

        self.root.configure(menu=menubar)
        self.file_menu, self.edit_menu = file_menu, edit_menu

        self.context_menu = tk.Menu(self.root, tearoff=False)
        self.context_menu.add_command(label="開啟", command=self.open_output)
        self.context_menu.add_command(label="開啟檔案位置", command=self.reveal_output)
        self.context_menu.add_command(label="複製 Markdown", command=self.copy_markdown)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="重新轉換", command=self.reconvert_selected)
        self.context_menu.add_command(label="移除", command=self.remove_selected)

    def _build_layout(self) -> None:
        tk, ttk = self.tk, self.ttk
        gap = round(12 * self.scale)
        small = round(6 * self.scale)
        main = ttk.Frame(self.root, padding=(gap, gap, gap, small))
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        # Toolbar: add things on the leading side, the one primary action on the trailing side.
        toolbar = ttk.Frame(main)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(2, weight=1)
        ttk.Button(toolbar, text="加入檔案…", command=self.choose_files).grid(row=0, column=0)
        ttk.Button(toolbar, text="加入資料夾…", command=self.choose_folder).grid(row=0, column=1, padx=(small, gap))
        self.url_entry = ttk.Entry(toolbar, textvariable=self.url_text)
        self.url_entry.grid(row=0, column=2, sticky="ew")
        self.url_entry.bind("<Return>", lambda _e: self.add_url())
        self._install_placeholder(self.url_entry, "貼上網址後按 Enter 加入（網頁、YouTube、Wikipedia）")
        self.primary = ttk.Button(toolbar, text="轉換", style="Accent.TButton", command=self._primary_action, width=14)
        self.primary.grid(row=0, column=3, padx=(gap, 0))

        # Where to save, and the less common options behind a pull-down.
        options = ttk.Frame(main)
        options.grid(row=1, column=0, sticky="ew", pady=(small, small))
        options.columnconfigure(4, weight=1)
        ttk.Label(options, text="儲存到", style="SecondaryBody.TLabel").grid(row=0, column=0, padx=(0, small))
        self.beside_radio = ttk.Radiobutton(options, text="原始檔案所在的資料夾", value="beside",
                                            variable=self.output_mode, command=self._output_mode_changed)
        self.beside_radio.grid(row=0, column=1)
        self.folder_radio = ttk.Radiobutton(options, text="其他資料夾", value="folder",
                                            variable=self.output_mode, command=self._output_mode_changed)
        self.folder_radio.grid(row=0, column=2, padx=(gap, 0))
        self.folder_button = ttk.Button(options, text="選擇…", command=self.choose_output_dir)
        self.folder_button.grid(row=0, column=3, padx=(small, 0))
        self.folder_label = ttk.Label(options, textvariable=self.output_dir, style="Secondary.TLabel")
        self.folder_label.grid(row=0, column=4, sticky="w", padx=(small, gap))
        self.options_button = ttk.Menubutton(options, text="選項")
        self.options_button.grid(row=0, column=5, sticky="e")
        options_menu = tk.Menu(self.options_button, tearoff=False)
        options_menu.add_checkbutton(label="加入資料夾時包含子資料夾", variable=self.recursive, command=self._rescan_folders)
        options_menu.add_checkbutton(label="略過已經有 .md 的檔案", variable=self.skip_existing)
        options_menu.add_checkbutton(label="在 Markdown 中保留內嵌圖片（base64）", variable=self.keep_data_uris)
        self.options_button["menu"] = options_menu

        # Content: the list of documents and the Markdown preview, side by side.
        paned = ttk.PanedWindow(main, orient="horizontal")
        paned.grid(row=2, column=0, sticky="nsew", pady=(small, 0))

        list_side = ttk.Frame(paned)
        list_side.rowconfigure(0, weight=1)
        list_side.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(list_side, columns=("type", "status"), show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="名稱", anchor="w")
        self.tree.heading("type", text="類型", anchor="w")
        self.tree.heading("status", text="狀態", anchor="w")
        self.tree.column("#0", width=round(200 * self.scale), minwidth=round(120 * self.scale), stretch=True)
        self.tree.column("type", width=round(92 * self.scale), minwidth=round(70 * self.scale), stretch=False)
        self.tree.column("status", width=round(118 * self.scale), minwidth=round(100 * self.scale), stretch=False)
        tree_scroll = ttk.Scrollbar(list_side, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_selection())
        self.tree.bind("<Double-1>", lambda _e: self.open_output())
        self.tree.bind("<Button-3>", self._open_context_menu)
        self.tree.bind("<Delete>", lambda _e: self.remove_selected())

        # Empty state, drawn over the list until something is added.
        self.drop_zone = tk.Canvas(list_side, highlightthickness=0, borderwidth=0)
        self.drop_zone.bind("<Configure>", lambda _e: self._draw_drop_zone())
        self.drop_zone_button = ttk.Button(self.drop_zone, text="選擇檔案…", command=self.choose_files)
        paned.add(list_side, weight=2)

        preview_side = ttk.Frame(paned, padding=(gap, 0, 0, 0))
        preview_side.rowconfigure(1, weight=1)
        preview_side.columnconfigure(0, weight=1)
        header = ttk.Frame(preview_side)
        header.grid(row=0, column=0, sticky="ew", pady=(0, small))
        header.columnconfigure(0, weight=1)
        self.preview_title = ttk.Label(header, text="預覽", style="Title.TLabel")
        self.preview_title.grid(row=0, column=0, sticky="w")
        self.preview_subtitle = ttk.Label(header, text="", style="Secondary.TLabel")
        self.preview_subtitle.grid(row=1, column=0, sticky="w")
        self.copy_button = ttk.Button(header, text="複製", command=self.copy_markdown)
        self.copy_button.grid(row=0, column=1, rowspan=2, padx=(small, 0))
        self.open_button = ttk.Button(header, text="開啟", command=self.open_output)
        self.open_button.grid(row=0, column=2, rowspan=2, padx=(small, 0))
        self.reveal_button = ttk.Button(header, text="開啟檔案位置", command=self.reveal_output)
        self.reveal_button.grid(row=0, column=3, rowspan=2, padx=(small, 0))
        text_frame = ttk.Frame(preview_side)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.preview = tk.Text(
            text_frame, wrap="word", relief="flat", borderwidth=0, font=self.fonts["body"],
            padx=round(18 * self.scale), pady=round(14 * self.scale), spacing1=round(2 * self.scale),
            spacing3=round(2 * self.scale), undo=False,
        )
        preview_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.grid(row=0, column=0, sticky="nsew")
        preview_scroll.grid(row=0, column=1, sticky="ns")
        self._make_read_only(self.preview)
        for name in ("h1", "h2", "h3", "bold", "mono"):
            self.preview.tag_configure(name, font=self.fonts[name])
        self.preview.tag_configure("h1", spacing1=round(10 * self.scale), spacing3=round(6 * self.scale))
        self.preview.tag_configure("h2", spacing1=round(8 * self.scale), spacing3=round(4 * self.scale))
        self.preview.tag_configure("h3", spacing1=round(6 * self.scale), spacing3=round(2 * self.scale))
        self.preview.tag_configure("code", font=self.fonts["mono"])
        self.preview.tag_configure("table", font=self.fonts["mono"])
        self.preview.tag_configure("link", underline=True)
        self.preview.tag_configure("quote", lmargin1=round(12 * self.scale), lmargin2=round(12 * self.scale))
        self.preview.tag_configure("message_title", font=self.fonts["title"], spacing1=round(24 * self.scale),
                                   spacing3=round(6 * self.scale), justify="center")
        self.preview.tag_configure("message", justify="center", spacing3=round(4 * self.scale))
        self.preview.tag_configure("detail", font=self.fonts["caption"], spacing1=round(16 * self.scale))
        paned.add(preview_side, weight=3)
        self.paned = paned

        def first_split(event):  # once the pane has a real width, give the list ~44%
            if event.width > 200:
                paned.sashpos(0, round(event.width * 0.44))
                paned.unbind("<Configure>", split_binding)

        split_binding = paned.bind("<Configure>", first_split, add="+")

        # Status line: information only; every action lives above.
        status = ttk.Frame(main)
        status.grid(row=3, column=0, sticky="ew", pady=(small, 0))
        status.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(status, textvariable=self.status_text, style="Secondary.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(status, mode="determinate", length=round(220 * self.scale))
        self.progress.grid(row=0, column=1, sticky="e")
        self.progress.grid_remove()

        if self.dnd is not None:
            for widget in (self.root, self.tree, self.drop_zone, self.preview):
                widget.drop_target_register(self.dnd.DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._on_drag_enter)
                widget.dnd_bind("<<DropPosition>>", lambda e: e.action)
                widget.dnd_bind("<<DropLeave>>", self._on_drag_leave)
                widget.dnd_bind("<<Drop>>", self._on_drop)

    def _install_placeholder(self, entry, text: str) -> None:
        def show(_e=None):
            if not self.url_text.get():
                entry.configure(style="Placeholder.TEntry")
                self.url_text.set(text)
                entry._placeholder = True

        def hide(_e=None):
            if getattr(entry, "_placeholder", False):
                self.url_text.set("")
                entry.configure(style="TEntry")
                entry._placeholder = False

        entry.bind("<FocusIn>", hide, add="+")
        entry.bind("<FocusOut>", show, add="+")
        show()

    def _make_read_only(self, text) -> None:
        navigation = {"Up", "Down", "Left", "Right", "Prior", "Next", "Home", "End", "Tab"}

        def on_key(event):
            ctrl = event.state & 0x4
            if event.keysym in navigation or (ctrl and event.keysym.lower() in ("c", "a")):
                return None
            return "break"

        text.bind("<Key>", on_key)
        for sequence in ("<<Paste>>", "<<Cut>>", "<<Clear>>"):
            text.bind(sequence, lambda _e: "break")
        text.bind("<Control-a>", lambda _e: (text.tag_add("sel", "1.0", "end"), "break")[1])

    def _bind_keys(self) -> None:
        r = self.root
        r.bind_all("<Control-o>", lambda _e: self.choose_files())
        r.bind_all("<Control-O>", lambda _e: self.choose_folder())
        r.bind_all("<Control-l>", lambda _e: self._focus_url())
        r.bind_all("<Control-Return>", lambda _e: self.start())
        r.bind_all("<Control-C>", lambda _e: self.copy_markdown())
        r.bind_all("<Escape>", lambda _e: self.stop())
        self.tree.bind("<Control-a>", lambda _e: (self._select_all(), "break")[1])

    # ----- empty state & drag and drop ------------------------------------ #
    def _draw_drop_zone(self) -> None:
        canvas = self.drop_zone
        canvas.delete("all")
        c = self.colors
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 10 or h < 10:
            return
        inset = round(4 * self.scale)
        active = self.dragging
        canvas.create_rectangle(
            inset, inset, w - inset, h - inset,
            outline=c["accent"] if active else c["border"], width=max(1, round(1.5 * self.scale)),
            dash=() if active else (round(6 * self.scale), round(4 * self.scale)),
            fill=c["drop"] if active else c["bg"],
        )
        k = self.scale
        cx, cy = w / 2, h / 2 - 60 * k
        ink = c["accent"] if active else c["secondary"]
        stroke = max(2, round(2 * k))
        # A page with a folded corner, and an arrow dropping into it.
        left, top, right, bottom, fold = cx - 22 * k, cy - 26 * k, cx + 22 * k, cy + 30 * k, 12 * k
        canvas.create_line(left, top, right - fold, top, right, top + fold, right, bottom, left, bottom, left, top,
                           fill=ink, width=stroke, joinstyle="round")
        canvas.create_line(right - fold, top, right - fold, top + fold, right, top + fold, fill=ink, width=stroke)
        canvas.create_line(cx, cy - 12 * k, cx, cy + 16 * k, fill=ink, width=stroke, arrow="last",
                           arrowshape=(10 * k, 10 * k, 5 * k))
        title = "放開以加入" if active else ("把檔案或資料夾拖曳到這裡" if self.dnd else "加入要轉換的檔案")
        canvas.create_text(cx, cy + 58 * k, text=title, font=self.fonts["strong"], fill=c["text"])
        canvas.create_text(
            cx, cy + 92 * k, justify="center", font=self.fonts["caption"], fill=c["secondary"],
            text="PDF、Word、PowerPoint、Excel、HTML、\nCSV、EPUB、ZIP 等格式都會轉成 Markdown",
        )
        canvas.create_window(cx, cy + 136 * k, window=self.drop_zone_button)

    def _on_drag_enter(self, event):
        self.dragging = True
        if not self.items:
            self._draw_drop_zone()
        else:
            self._set_status("放開以加入清單")
        return event.action

    def _on_drag_leave(self, event):
        self.dragging = False
        self._draw_drop_zone()
        self._refresh_status_idle()
        return event.action

    def _on_drop(self, event):
        self.dragging = False
        self.add_inputs(self.root.tk.splitlist(event.data))
        self._draw_drop_zone()
        return event.action

    # ----- adding & removing ---------------------------------------------- #
    def _all_keys(self) -> set:
        return {item.key for item in self.items.values()}

    def add_inputs(self, values) -> int:
        added = 0
        keys = self._all_keys()
        for raw in values:
            value = str(raw).strip().strip('"')
            if not value:
                continue
            if convert.is_url(value):
                item = Item("url", value)
            elif Path(value).is_dir():
                item = Item("folder", str(Path(value)))
            else:
                item = Item("file", str(Path(value)))
            if item.key in keys:
                continue
            keys.add(item.key)
            iid = self.tree.insert("", "end", text=item.display_name, open=True)
            self.items[iid] = item
            if item.kind == "folder":
                self._scan_folder(iid)
                keys |= {self.items[c].key for c in self.tree.get_children(iid)}
            self._render_row(iid)
            added += 1
        if added:
            self._log(f"加入 {added} 個項目")
            if not self.busy:
                self.run_summary = {}
        self._refresh()
        return added

    def _scan_folder(self, folder_iid: str) -> None:
        folder = self.items[folder_iid]
        found = convert.scan_folder(Path(folder.value), self.recursive.get())
        existing = {self.items[c].key: c for c in self.tree.get_children(folder_iid)}
        other_keys = self._all_keys() - set(existing)
        wanted = []
        for source in found:
            child = Item("file", source.value, scan_root=source.scan_root)
            if child.key in other_keys:
                continue
            wanted.append(child)
        wanted_keys = {c.key for c in wanted}
        for key, iid in existing.items():
            if key not in wanted_keys:
                self.tree.delete(iid)
                self.items.pop(iid, None)
        for child in wanted:
            if child.key in existing:
                continue
            rel = Path(child.value).relative_to(folder.value).as_posix()
            iid = self.tree.insert(folder_iid, "end", text=rel)
            self.items[iid] = child
            self._render_row(iid)

    def _rescan_folders(self) -> None:
        if self.busy:
            return
        for iid, item in list(self.items.items()):
            if item.kind == "folder" and self.tree.exists(iid):
                self._scan_folder(iid)
                self._render_row(iid)
        self._refresh()

    def add_url(self) -> None:
        url = self.url_text.get().strip()
        if not url or getattr(self.url_entry, "_placeholder", False):
            return
        if not convert.is_url(url):
            url = "https://" + url
        if self.add_inputs([url]):
            self.url_text.set("")
        else:
            self._set_status("這個網址已經在清單中")

    def choose_files(self) -> None:
        types = [
            ("支援的檔案", " ".join(f"*{ext}" for ext in sorted(convert.SUPPORTED_EXTENSIONS))),
            ("所有檔案", "*.*"),
        ]
        paths = self.filedialog.askopenfilenames(parent=self.root, title="加入檔案", filetypes=types)
        self.add_inputs(paths)

    def choose_folder(self) -> None:
        folder = self.filedialog.askdirectory(parent=self.root, title="加入資料夾")
        if folder:
            self.add_inputs([folder])

    def choose_output_dir(self) -> bool:
        folder = self.filedialog.askdirectory(parent=self.root, title="選擇儲存 .md 檔的資料夾")
        if folder:
            self.output_dir.set(str(Path(folder)))
            self.output_mode.set("folder")
        self._refresh()
        return bool(folder)

    def _output_mode_changed(self) -> None:
        if self.output_mode.get() == "folder" and not self.output_dir.get():
            if not self.choose_output_dir():
                self.output_mode.set("beside")
        self._refresh()

    def remove_selected(self) -> None:
        if self.busy:
            self._set_status("轉換進行中，完成或停止後才能移除項目")
            return
        touched_folders = set()
        for iid in self.tree.selection():
            if self.tree.exists(iid):
                touched_folders.add(self.tree.parent(iid))
                for child in self.tree.get_children(iid):
                    self.items.pop(child, None)
                self.tree.delete(iid)
                self.items.pop(iid, None)
        for folder in touched_folders - {""}:
            if not self.tree.exists(folder):
                continue
            if self.tree.get_children(folder):
                self._render_row(folder)
            else:  # every file in it was removed, so the folder goes too
                self.tree.delete(folder)
                self.items.pop(folder, None)
        self._refresh()

    def clear_items(self) -> None:
        if self.busy:
            self._set_status("轉換進行中，完成或停止後才能清除清單")
            return
        self.tree.delete(*self.tree.get_children(""))
        self.items.clear()
        self.run_summary = {}
        self._refresh()

    def _select_all(self) -> None:
        self.tree.selection_set(self._leaf_iids())

    def _focus_url(self) -> None:
        self.url_entry.focus_set()

    # ----- conversion ------------------------------------------------------ #
    def _leaf_iids(self) -> List[str]:
        result = []
        for iid in self.tree.get_children(""):
            if self.items[iid].kind == "folder":
                result.extend(self.tree.get_children(iid))
            else:
                result.append(iid)
        return result

    def _targets(self, only: Optional[List[str]] = None) -> List[str]:
        if only is not None:
            return list(only)
        leaves = self._leaf_iids()
        todo = [i for i in leaves if self.items[i].status in ("pending", "failed")]
        return todo or leaves

    def _primary_action(self) -> None:
        if self.busy:
            self.stop()
        else:
            self.start()

    def _load_converter(self):
        with self.converter_lock:
            if self.converter is None and self.converter_error is None:
                try:
                    self.converter = convert.build_converter()
                    self.events.put(("ready",))
                except Exception as exc:
                    self.converter_error = convert.describe_error(exc)
                    self.events.put(("engine_failed", self.converter_error))
            return self.converter

    def start(self, only: Optional[List[str]] = None) -> None:
        if self.busy:
            return
        targets = self._targets(only)
        if not targets:
            return
        out_dir = Path(self.output_dir.get()) if self.output_mode.get() == "folder" and self.output_dir.get() else None
        for iid in targets:
            item = self.items[iid]
            item.status, item.output, item.markdown, item.error, item.error_type = "pending", None, None, None, None
            self._render_row(iid)
        self.busy = True
        self.stop_event.clear()
        self.run_summary = {"total": len(targets)}
        self.progress.configure(mode="determinate", maximum=len(targets), value=0)
        self.progress.grid()
        self._refresh()
        jobs = (targets, out_dir, self.skip_existing.get(), self.keep_data_uris.get())
        threading.Thread(target=self._worker, args=jobs, daemon=True).start()

    def reconvert_selected(self) -> None:
        selected = [i for i in self.tree.selection() if self.items.get(i) and self.items[i].kind in LEAF_KINDS]
        if selected:
            self.start(only=selected)

    def stop(self) -> None:
        if self.busy and not self.stop_event.is_set():
            self.stop_event.set()
            self._set_status("正在停止，目前的檔案完成後就會停下…")

    def _worker(self, targets, out_dir, skip_existing, keep_data_uris) -> None:
        try:
            if self.converter is None:
                self.events.put(("waiting_engine",))
            converter = self._load_converter()
            if converter is None:
                return
            files = [i for i in targets if self.items[i].kind == "file"]
            urls = [i for i in targets if self.items[i].kind == "url"]
            groups = [(files, out_dir), (urls, out_dir or DEFAULT_URL_DIR)]
            for iids, destination in groups:
                if not iids or self.stop_event.is_set():
                    continue
                sources = [convert.Source(self.items[i].value, self.items[i].scan_root) for i in iids]
                convert.convert_batch(
                    converter, sources, destination,
                    keep_data_uris=keep_data_uris, skip_existing=skip_existing,
                    on_start=lambda n, _t, _s, ids=iids: self.events.put(("working", ids[n - 1])),
                    on_result=lambda n, _t, r, ids=iids: self.events.put(("result", ids[n - 1], r)),
                    should_stop=self.stop_event.is_set,
                )
        except Exception as exc:
            self.events.put(("log", f"錯誤：{convert.describe_error(exc)}"))
        finally:
            self.events.put(("finished",))

    def _poll_events(self) -> None:
        try:
            while True:
                self._handle_event(self.events.get_nowait())
        except queue.Empty:
            pass
        self.root.after(80, self._poll_events)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "ready":
            if not self.busy:
                self._refresh_status_idle()
        elif kind == "engine_failed":
            self._log(f"無法啟動轉換引擎：{event[1]}")
            self._set_status("無法啟動轉換引擎，詳細資訊請見「檢視 › 轉換紀錄」")
        elif kind == "waiting_engine":
            self.progress.configure(mode="indeterminate")
            self.progress.start(12)
            self._set_status("正在準備轉換引擎…")
        elif kind == "log":
            self._log(event[1])
        elif kind == "working":
            iid = event[1]
            if str(self.progress.cget("mode")) == "indeterminate":
                self.progress.stop()
                self.progress.configure(mode="determinate", value=0)
            if iid in self.items:
                self.items[iid].status = "working"
                self._render_row(iid)
                done = int(self.progress.cget("value"))
                self._set_status(f"正在轉換 {self.items[iid].display_name}（{done + 1}/{self.run_summary['total']}）")
        elif kind == "result":
            iid, result = event[1], event[2]
            self.progress.configure(value=float(self.progress.cget("value")) + 1)
            self._apply_result(iid, result)
        elif kind == "finished":
            self._finish_run()

    def _apply_result(self, iid: str, result: "convert.Result") -> None:
        item = self.items.get(iid)
        if item is None:
            return
        item.output = Path(result.output) if result.output else None
        if result.error:
            item.status, item.error, item.error_type = "failed", result.error, result.error_type
            self._log(f"✕ {item.value}\n   {result.error}")
        elif result.markdown is None:
            item.status = "skipped"
            self._log(f"– {item.value}（已有 {result.output}）")
        else:
            item.markdown = result.markdown
            item.status = "done" if result.markdown.strip() else "empty"
            self._log(f"✓ {item.value} → {result.output}")
        self.run_summary[item.status] = self.run_summary.get(item.status, 0) + 1
        self._render_row(iid)
        parent = self.tree.parent(iid)
        if parent:
            self._render_row(parent)
        if iid in self.tree.selection():
            self._show_selection()

    def _finish_run(self) -> None:
        self.progress.stop()
        self.progress.grid_remove()
        stopped = self.stop_event.is_set()
        self.busy = False
        for iid, item in self.items.items():
            if item.status == "working":
                item.status = "pending"
                self._render_row(iid)
        s = self.run_summary
        converted = s.get("done", 0) + s.get("empty", 0)
        parts = [f"已轉換 {converted} 個項目"]
        if s.get("skipped"):
            parts.append(f"略過 {s['skipped']} 個已有 .md 的檔案")
        if s.get("failed"):
            parts.append(f"{s['failed']} 個失敗")
        if s.get("empty"):
            parts.append(f"{s['empty']} 個沒有擷取到文字")
        message = "，".join(parts) + "。"
        if stopped:
            message = "已停止。" + message
        self.run_summary["message"] = message
        self._set_status(message)
        self._log(message)
        if not self.tree.selection():
            first = next((i for i in self._leaf_iids() if self.items[i].status in ("failed", "done", "empty")), None)
            if first:
                self.tree.selection_set(first)
                self.tree.see(first)
        self._refresh(keep_status=True)

    # ----- rendering ------------------------------------------------------- #
    def _render_row(self, iid: str) -> None:
        if not self.tree.exists(iid):
            return
        item = self.items[iid]
        if item.kind == "folder":
            children = [self.items[c] for c in self.tree.get_children(iid)]
            finished = sum(1 for c in children if c.status in ("done", "empty", "skipped"))
            if not children:
                status = "沒有可轉換的檔案"
            elif finished:
                status = f"{finished}/{len(children)} 已完成"
            else:
                status = f"{len(children)} 個檔案"
            self.tree.item(iid, values=(item.type_label, status), tags=("folder",))
        else:
            self.tree.item(iid, values=(item.type_label, STATUS_TEXT[item.status]), tags=(item.status,))

    def _refresh(self, keep_status: bool = False) -> None:
        has_items = bool(self.items)
        if has_items:
            self.drop_zone.place_forget()
        else:
            self.drop_zone.place(x=0, y=0, relwidth=1, relheight=1)
            self._draw_drop_zone()
        leaves = self._leaf_iids()
        todo = [i for i in leaves if self.items[i].status in ("pending", "failed")]
        if self.busy:
            self.primary.configure(text="停止", style="TButton", state="normal")
        elif todo:
            self.primary.configure(text=f"轉換 {len(todo)} 個項目", style="Accent.TButton", state="normal")
        elif leaves:
            self.primary.configure(text=f"重新轉換 {len(leaves)} 個", style="Accent.TButton", state="normal")
        else:
            self.primary.configure(text="轉換", style="Accent.TButton", state="disabled")
        idle = "disabled" if self.busy else "normal"
        for widget in (self.beside_radio, self.folder_radio, self.options_button):
            widget.configure(state=idle)
        folder_mode = self.output_mode.get() == "folder"
        self.folder_button.configure(state="normal" if (folder_mode and not self.busy) else "disabled")
        self.folder_label.grid() if folder_mode else self.folder_label.grid_remove()
        for menu, label in ((self.file_menu, "轉換"), (self.edit_menu, "清除清單"), (self.edit_menu, "移除")):
            menu.entryconfigure(label, state="disabled" if (self.busy or not leaves) else "normal")
        self.file_menu.entryconfigure("停止", state="normal" if self.busy else "disabled")
        self._show_selection()
        if not keep_status and not self.busy:
            self._refresh_status_idle()

    def _refresh_status_idle(self) -> None:
        if self.busy:
            return
        if self.converter_error:
            self._set_status("無法啟動轉換引擎，詳細資訊請見「檢視 › 轉換紀錄」")
            return
        leaves = self._leaf_iids()
        if self.run_summary.get("total"):
            return  # keep the summary of the last run until the list changes
        if not leaves:
            if self.converter is None:
                self._set_status("正在準備轉換引擎，第一次啟動約需數秒…")
            else:
                self._set_status("加入檔案、資料夾或網址就可以開始。")
        else:
            self._set_status(f"清單中有 {len(leaves)} 個項目。按「轉換」或 Ctrl+Enter 開始。")

    def _selected_item(self) -> Optional[Item]:
        selection = self.tree.selection()
        if len(selection) != 1:
            return None
        return self.items.get(selection[0])

    def _show_selection(self) -> None:
        selection = self.tree.selection()
        item = self._selected_item()
        has_output = bool(item and item.output and item.status in ("done", "empty", "skipped"))
        self.copy_button.configure(state="normal" if has_output else "disabled")
        self.open_button.configure(state="normal" if has_output else "disabled")
        self.reveal_button.configure(state="normal" if has_output else "disabled")

        if not self.items:
            self._preview_message("還沒有任何檔案", "轉換完成後，Markdown 會顯示在這裡。")
            self._set_preview_header("預覽", "")
        elif item is None:
            title = f"已選取 {len(selection)} 個項目" if selection else "預覽"
            self._set_preview_header(title, "")
            self._preview_message("選取一個項目來預覽", "已完成的項目會在這裡顯示轉換後的 Markdown。")
        elif item.kind == "folder":
            children = [self.items[c] for c in self.tree.get_children(selection[0])]
            finished = sum(1 for c in children if c.status in ("done", "empty", "skipped"))
            self._set_preview_header(item.display_name, item.value)
            self._preview_message(f"{len(children)} 個檔案", f"已完成 {finished} 個。選取資料夾中的檔案來預覽。")
        elif item.status == "pending":
            self._set_preview_header(item.display_name, item.value)
            self._preview_message("尚未轉換", "按「轉換」或 Ctrl+Enter 開始。")
        elif item.status == "working":
            self._set_preview_header(item.display_name, item.value)
            self._preview_message("轉換中…", "")
        elif item.status == "failed":
            title, advice = friendly_error(item.error_type)
            self._set_preview_header(item.display_name, item.value)
            self._preview_message(title, advice, detail=item.error, tone="danger")
        else:
            self._set_preview_header(Path(str(item.output)).name, str(item.output))
            if item.markdown is None and item.output and item.output.exists():
                try:
                    item.markdown = item.output.read_text(encoding="utf-8")
                except OSError:
                    item.markdown = ""
            if item.status == "empty" or not (item.markdown or "").strip():
                self._preview_message(
                    "沒有擷取到文字",
                    "這個檔案可能只有圖片（例如掃描的 PDF），MarkItDown 無法辨識圖片中的文字。已產生空白的 .md 檔。",
                    tone="caution",
                )
            else:
                self._preview_markdown(item.markdown or "")

    def _set_preview_header(self, title: str, subtitle: str) -> None:
        self.preview_title.configure(text=title)
        self.preview_subtitle.configure(text=self._ellipsize_middle(subtitle, 90))

    @staticmethod
    def _ellipsize_middle(text: str, limit: int) -> str:
        if len(text) <= limit:
            return text
        keep = (limit - 1) // 2
        return text[:keep] + "…" + text[-keep:]

    def _preview_message(self, title: str, body: str, detail: Optional[str] = None, tone: str = "") -> None:
        p = self.preview
        p.delete("1.0", "end")
        p.insert("end", title + "\n", ("message_title",) + ((tone,) if tone else ()))
        if body:
            p.insert("end", body + "\n", ("message", "muted"))
        if detail:
            p.insert("end", "技術細節\n" + detail.strip() + "\n", ("detail", "muted"))
        p.yview_moveto(0)

    def _preview_markdown(self, markdown: str) -> None:
        p = self.preview
        p.delete("1.0", "end")
        truncated = len(markdown) > PREVIEW_LIMIT
        segments = render_markdown(markdown[:PREVIEW_LIMIT])
        if truncated:
            segments.append(("\n…預覽只顯示前 200,000 個字元，完整內容請按「開啟」。\n", ("muted",)))
        for start in range(0, len(segments), 400):
            args = []
            for text, tags in segments[start:start + 400]:
                args.extend((text, tags))
            p.insert("end", *args)
        p.yview_moveto(0)

    def _set_status(self, text: str) -> None:
        self.status_text.set(text)

    # ----- result actions --------------------------------------------------- #
    def copy_markdown(self) -> None:
        item = self._selected_item()
        if not item or item.status not in ("done", "empty", "skipped"):
            return
        if item.markdown is None and item.output and item.output.exists():
            item.markdown = item.output.read_text(encoding="utf-8")
        self.root.clipboard_clear()
        self.root.clipboard_append(item.markdown or "")
        self._set_status(f"已複製 {Path(str(item.output)).name} 的 Markdown")
        self.root.after(4000, self._restore_status)

    def _restore_status(self) -> None:
        """Brief confirmations give way to the lasting status again."""
        if self.busy:
            return
        if self.run_summary.get("message"):
            self._set_status(self.run_summary["message"])
        else:
            self._refresh_status_idle()

    def open_output(self) -> None:
        item = self._selected_item()
        if item and item.output and item.output.exists():
            open_path(item.output)

    def reveal_output(self) -> None:
        item = self._selected_item()
        if item and item.output and item.output.exists():
            reveal_path(item.output)

    def _open_context_menu(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if row not in self.tree.selection():
            self.tree.selection_set(row)
        item = self._selected_item()
        has_output = bool(item and item.output and item.status in ("done", "empty", "skipped"))
        for label in ("開啟", "開啟檔案位置", "複製 Markdown"):
            self.context_menu.entryconfigure(label, state="normal" if has_output else "disabled")
        for label in ("重新轉換", "移除"):
            self.context_menu.entryconfigure(label, state="disabled" if self.busy else "normal")
        self.context_menu.tk_popup(event.x_root, event.y_root)

    # ----- log, help, about ------------------------------------------------ #
    def _log(self, line: str) -> None:
        self.log_lines.append(line)
        if self.log_window is not None:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", line + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

    def show_log(self) -> None:
        if self.log_window is not None:
            self.log_window.deiconify()
            self.log_window.lift()
            return
        win = self.tk.Toplevel(self.root)
        win.title("轉換紀錄")
        win.geometry(f"{round(640 * self.scale)}x{round(400 * self.scale)}")
        frame = self.ttk.Frame(win, padding=round(12 * self.scale))
        frame.pack(fill="both", expand=True)
        self.log_text = self.tk.Text(frame, wrap="word", relief="flat", font=self.fonts["mono"],
                                     padx=round(10 * self.scale), pady=round(8 * self.scale))
        scroll = self.ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set, background=self.colors["surface"], foreground=self.colors["text"])
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.log_text.insert("end", "\n".join(self.log_lines) + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")
        set_title_bar_dark(win, self.appearance == "dark")

        def close():
            self.log_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close)
        self.log_window = win

    def _show_formats(self) -> None:
        self.messagebox.showinfo(
            "支援的格式",
            "文件：PDF、Word (.docx)、PowerPoint (.pptx)、Excel (.xlsx、.xls)、EPUB\n"
            "網頁與資料：HTML、CSV、JSON、XML、RSS、Jupyter Notebook (.ipynb)\n"
            "其他：Outlook 郵件 (.msg)、ZIP（逐一轉換內含檔案）、圖片（EXIF 資訊）、音訊（語音轉文字，需要網路）\n"
            "網址：一般網頁、YouTube（標題、說明、字幕）、Wikipedia\n\n"
            "掃描的 PDF 或圖片中的文字無法辨識（沒有 OCR）。",
            parent=self.root,
        )

    def _show_about(self) -> None:
        self.messagebox.showinfo(
            f"關於 {APP_NAME}",
            f"{APP_NAME}\n{convert.versions_text()}\n\n"
            "以 Microsoft MarkItDown（MIT 授權）為核心，把文件與網頁轉成 Markdown。",
            parent=self.root,
        )

    def _on_close(self) -> None:
        self.stop_event.set()
        self.root.destroy()


# --------------------------------------------------------------------------- #
# Markdown preview: quiet syntax colouring
# --------------------------------------------------------------------------- #
_INLINE = re.compile(r"(`[^`\n]+`)|(\*\*[^*\n]+\*\*)|(!?\[[^\]\n]*\]\([^)\n]*\))")
_HEADING = re.compile(r"^(#{1,6})(\s+)(.*)$")
_LIST = re.compile(r"^(\s*(?:[-*+]|\d+[.)])\s+)(.*)$")
_RULE = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,})\s*$")
_COMMENT = re.compile(r"^\s*<!--.*-->\s*$")


def _inline(text: str, base: tuple) -> list:
    out = []
    pos = 0
    for m in _INLINE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], base))
        token = m.group(0)
        if m.group(1):
            out += [("`", base + ("syntax",)), (token[1:-1], base + ("code",)), ("`", base + ("syntax",))]
        elif m.group(2):
            out += [("**", base + ("syntax",)), (token[2:-2], base + ("bold",)), ("**", base + ("syntax",))]
        else:
            label_end = token.index("](")
            prefix = "![" if token.startswith("!") else "["
            out += [
                (prefix, base + ("syntax",)),
                (token[len(prefix):label_end], base + ("link",)),
                (token[label_end:], base + ("syntax",)),
            ]
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], base))
    return out


def render_markdown(markdown: str) -> list:
    """Split Markdown into (text, tags) runs: headings and emphasis stand out,
    the syntax characters recede, and the text stays exactly what is saved."""
    segments = []
    in_code = False
    for line in markdown.split("\n"):
        if line.lstrip().startswith("```"):
            in_code = not in_code
            segments.append((line, ("syntax", "code")))
        elif in_code:
            segments.append((line, ("code",)))
        elif m := _HEADING.match(line):
            level = "h" + str(min(len(m.group(1)), 3))
            segments.append((m.group(1) + m.group(2), ("syntax", level)))
            segments += _inline(m.group(3), (level,))
        elif line.lstrip().startswith("|"):
            for part in re.split(r"(\|)", line):
                if part:
                    is_pipe = part == "|" or set(part.strip()) <= set("-: ")
                    segments.append((part, ("table", "syntax") if is_pipe else ("table",)))
        elif _RULE.match(line) or _COMMENT.match(line):
            segments.append((line, ("syntax",)))
        elif line.startswith(">"):
            segments.append((">", ("syntax", "quote")))
            segments += _inline(line[1:], ("quote",))
        elif m := _LIST.match(line):
            segments.append((m.group(1), ("syntax",)))
            segments += _inline(m.group(2), ())
        else:
            segments += _inline(line, ())
        segments.append(("\n", ()))
    return segments
