"""Quick Convert: a small floating panel between Explorer and Claude.

Drop files on the well; each result becomes a file card that can be dragged
straight into Claude Desktop (or anywhere that takes files). One job, three
states: waiting for a drop, converting, results to hand over. Every drag has a
non-drag alternative: "選擇檔案…" in, and copy / open / show-in-folder out.
"""

from __future__ import annotations

import queue
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import convert
from motion import EASE, EASE_IN_OUT, EASE_OUT, Motion, blend, reduced_motion
from xlsx_compact import CONVERSATION_BUDGET, INDEX_NAME, SMALL_BUDGET, estimate_tokens

PANEL_TITLE = "快速轉換"
MAX_CARDS = 30
OFFER_MIN_TOKENS = SMALL_BUDGET  # "only the newest versions" is offered only when the whole is this big

# Pasted once into a Claude Project's instructions, so every conversation starts small.
PROJECT_INSTRUCTIONS = """我會上傳用「MarkItDown 轉換器」從 Excel 轉出的檔案。請照這個方式處理：
1. 先只讀「00_索引與邏輯.md」。它列出每張工作表、每一欄的類型，以及程式精確算好的合計、最大值、最小值和公式邏輯；這些數字可以直接引用，不要自己加總大量數字。
2. 能用索引回答就直接回答。需要明細時，告訴我要上傳哪一個資料檔（檔名和大約 token 數在索引最後的「資料檔」表），一次只要需要的那幾個，不要要求整包。
3. 問題只跟某個料號、客戶，或只跟最新的預測版本有關時，提醒我用轉換器的關鍵字篩選或「只轉最新一版」重新轉一次，檔案會小很多。
4. 資料是 Tab 分隔，第一欄是 Excel 列號（後面有 h 表示該列在 Excel 中隱藏），數值是完整精度。引用數字時請註明工作表和儲存格位置（例如 Summary!F12）。"""

# Motion budget (see motion.py). Nothing animates on a keyboard action, and with
# Windows' animation effects off nothing moves or resizes; colour fades stay, shorter.
HOVER_IN_MS, HOVER_OUT_MS = 150, 100  # drop-target highlight: in on drag enter, out faster
RESIZE_MS = 260  # the well folding to one row / opening up again
ENTER_MS = 240  # a result card arriving
STAGGER_MS = 50  # between cards that arrive together
PRESS_IN_MS, PRESS_OUT_MS = 100, 160  # card press: down quickly, back a little slower
PROGRESS_MS = 220
PRESS_SCALE = 0.97
ENTER_RISE = 8  # px the card content rises while it fades in
COMPACT_WELL = 76  # px, one-row well once there are results
LAYOUT_SWITCH = 132  # px; below this the well draws its one-row layout


@dataclass
class PanelResult:
    source: str
    output: Optional[Path] = None  # the file to hand over: the .md, or the index of a package
    tokens: int = 0
    data_files: int = 0  # data files beside the index (multi-file Excel package)
    data_tokens: int = 0
    excel: bool = False
    error: Optional[str] = None
    error_type: Optional[str] = None
    total_tokens: int = 0  # index + every data file
    latest: str = ""  # newest version of the workbook's version column, if it has one
    offers: Dict[int, int] = field(default_factory=dict)  # newest N versions -> estimated tokens
    variant: str = ""  # e.g. "只含最新 1 版" for an extract

    @property
    def ok(self) -> bool:
        return self.error is None and self.output is not None


def describe_tokens(n: int) -> str:
    """約 22,000 token / 約 540 萬 token: rounded, as an estimate should read."""
    if n >= 100_000_000:
        return f"約 {n / 100_000_000:.1f} 億 token"
    if n >= 100_000:
        return f"約 {round(n / 10_000):,} 萬 token"
    if n >= 1_000:
        return f"約 {round(n, -2):,} token"
    return f"約 {max(n, 1):,} token"


def measure(result: "convert.Result", variant: str = "") -> PanelResult:
    """Token estimates for a conversion result (runs on the worker thread)."""
    source = result.source.value
    excel = convert.is_excel_file(result.source)
    if result.error or result.output is None:
        return PanelResult(source, error=result.error or "沒有輸出", error_type=result.error_type, excel=excel)
    output = Path(result.output)
    text = result.markdown if result.markdown is not None else _read(output)
    item = PanelResult(source, output, estimate_tokens(text), excel=excel, variant=variant)
    if output.name == INDEX_NAME:
        for data in sorted(output.parent.glob("[0-9][0-9]_*.md")):
            item.data_files += 1
            item.data_tokens += estimate_tokens(_read(data))
    item.total_tokens = item.tokens + item.data_tokens
    versions = (result.info or {}).get("versions") or {}
    item.latest = str(versions.get("latest", ""))
    if item.total_tokens > OFFER_MIN_TOKENS:  # a small file is handed over whole; no extra choices
        item.offers = {int(n): int(t) for n, t in (versions.get("estimates") or {}).items()}
    return item


def _read(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return ""


def rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    """A rounded rectangle as a smoothed polygon (Tk has no native one)."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    points = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
              x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(points, smooth=True, splinesteps=12, **kw)


def fit_text(font, text: str, width: int) -> str:
    """Shorten from the middle so both the start and the extension stay visible."""
    if font.measure(text) <= width:
        return text
    keep = len(text)
    while keep > 4:
        keep -= 1
        head = text[: (keep + 1) // 2]
        tail = text[len(text) - keep // 2:]
        candidate = f"{head}…{tail}"
        if font.measure(candidate) <= width:
            return candidate
    return text[:1] + "…"


class QuickPanel:
    """The floating panel. Owned by the main ConverterWindow, whose engine,
    options, palette and fonts it shares."""

    def __init__(self, app, standalone: bool = False):
        tk, ttk = app.tk, app.ttk
        self.app, self.tk, self.ttk = app, tk, ttk
        self.standalone = standalone
        self.scale = app.scale
        self.dnd = app.dnd
        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.pending: List[Tuple[List[str], dict]] = []  # (paths, extra options) waiting for the current batch
        self.busy = False
        self.hovering = False
        self.status = ""
        self.progress_value = (0, 0)  # (finished, total) across everything dropped since idle
        self._offset = 0  # files finished before the batch now running
        self.cards: List[ResultCard] = []
        self.pinned = tk.BooleanVar(value=True)
        self.reduce_motion = reduced_motion()
        body = app.fonts["body"]
        self.link_font = app.tkfont.Font(root=app.root, family=body.actual("family"), size=body.actual("size"),
                                         underline=True)
        self._compact: Optional[bool] = None  # well layout: None until first laid out
        self._next_enter = 0.0  # perf_counter time the next card may start entering (stagger)

        k = self.scale
        win = tk.Toplevel(app.root)
        self.win = win
        win.title(PANEL_TITLE)
        width, height = round(380 * k), round(540 * k)
        x = max(0, win.winfo_screenwidth() - width - round(32 * k))
        win.geometry(f"{width}x{height}+{x}+{round(96 * k)}")
        win.minsize(round(320 * k), round(420 * k))
        win.attributes("-topmost", True)
        if not standalone and app.dnd is not None:
            try:  # a panel: no minimize button, no extra taskbar entry
                win.attributes("-toolwindow", True)
            except Exception:
                pass
        win.protocol("WM_DELETE_WINDOW", self.close)
        self.motion = Motion(win)
        self._build()
        self._bind_keys()
        self.repaint()
        win.after(80, self._poll)

    # ----- layout ---------------------------------------------------------- #
    def _build(self) -> None:
        tk, ttk, k = self.tk, self.ttk, self.scale
        gap = round(14 * k)
        frame = ttk.Frame(self.win, padding=(gap, gap, gap, round(10 * k)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        self.frame = frame

        # The well: drop target, progress, and the "choose files" alternative.
        self.well = tk.Canvas(frame, height=round(188 * k), highlightthickness=0, borderwidth=0)
        self.well.grid(row=0, column=0, sticky="nsew")
        self.well.bind("<Configure>", lambda _e: self._draw_well())
        self.choose_button = ttk.Button(self.well, text="選擇檔案…", command=self.choose_files)
        self.progress = ttk.Progressbar(self.well, mode="determinate")

        # Results: newest first, each a card that can be dragged out.
        header = ttk.Frame(frame)
        header.columnconfigure(0, weight=1)
        self.results_label = ttk.Label(header, text="轉換結果", style="Title.TLabel")
        self.results_label.grid(row=0, column=0, sticky="w")
        self.clear_button = ttk.Button(header, text="清除", command=self.clear)
        self.clear_button.grid(row=0, column=1, sticky="e")
        hint = ("把卡片拖進 Claude 的對話框；按兩下開啟，按右鍵可複製" if self.dnd is not None
                else "按兩下開啟；按右鍵可複製內容或開啟檔案位置")
        self.hint_label = ttk.Label(header, text=hint, style="Secondary.TLabel")
        self.hint_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(round(2 * k), 0))
        self.header = header

        holder = ttk.Frame(frame)
        self.holder = holder
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        self.list_canvas = tk.Canvas(holder, highlightthickness=0, borderwidth=0)
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(holder, orient="vertical", command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=self._on_scroll)
        self.list_frame = tk.Frame(self.list_canvas, borderwidth=0, highlightthickness=0)
        self.list_window = self.list_canvas.create_window(0, 0, window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>", lambda _e: self._update_scroll_region())
        self.list_canvas.bind("<Configure>", self._on_list_resize)
        for widget in (self.list_canvas, self.list_frame):
            widget.bind("<MouseWheel>", self._on_wheel)
            widget.bind("<Button-4>", lambda _e: self.list_canvas.yview_scroll(-1, "units"))
            widget.bind("<Button-5>", lambda _e: self.list_canvas.yview_scroll(1, "units"))

        # Footer: what the settings are, where to change them, whether to float.
        footer = ttk.Frame(frame)
        self.footer = footer
        footer.grid(row=3, column=0, sticky="ew", pady=(round(10 * k), 0))
        footer.columnconfigure(0, weight=1)
        # Row 1: what the settings are (full width, so it is never cut off). Row 2: float switch,
        # and the less frequent actions behind "更多".
        self.settings_label = ttk.Label(footer, style="Secondary.TLabel")
        self.settings_label.grid(row=0, column=0, columnspan=2, sticky="w")
        self.pin_switch = ttk.Checkbutton(footer, text="浮在其他視窗上方", variable=self.pinned, command=self._apply_pin,
                                          style="Switch.TCheckbutton" if self.app.sv_ttk else "TCheckbutton")
        self.pin_switch.grid(row=1, column=0, sticky="w", pady=(round(6 * k), 0))
        self.more_button = ttk.Menubutton(footer, text="更多")
        self.more_button.grid(row=1, column=1, sticky="e", pady=(round(6 * k), 0))
        more = tk.Menu(self.more_button, tearoff=False)
        more.add_command(label="轉換選項…", command=self.show_main_window)
        more.add_command(label="複製 Claude 專案指示", command=self.copy_instructions)
        self.more_button["menu"] = more

        if self.dnd is not None:
            for widget in (self.win, self.well, self.list_canvas, self.list_frame):
                widget.drop_target_register(self.dnd.DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._on_drag_enter)
                widget.dnd_bind("<<DropPosition>>", lambda e: e.action)
                widget.dnd_bind("<<DropLeave>>", self._on_drag_leave)
                widget.dnd_bind("<<Drop>>", self._on_drop)
        self._refresh()

    def _bind_keys(self) -> None:
        w = self.win
        # "break" keeps the main window's application-wide shortcuts from firing too.
        w.bind("<Control-o>", lambda _e: (self.choose_files(), "break")[1])
        w.bind("<Escape>", lambda _e: (self.close(), "break")[1])
        w.bind("<Control-C>", lambda _e: (self._copy_focused(), "break")[1])
        w.bind("<Control-Return>", lambda _e: "break")

    # ----- appearance ------------------------------------------------------ #
    @property
    def colors(self) -> dict:
        return self.app.colors

    def repaint(self) -> None:
        c = self.colors
        for widget in (self.well, self.list_canvas, self.list_frame):
            widget.configure(background=c["bg"])
        from gui import set_title_bar_dark

        set_title_bar_dark(self.win, self.app.appearance == "dark")
        self._draw_well()
        for card in self.cards:
            card.draw()

    @property
    def animating(self) -> bool:
        return self.motion.active or any(card.motion.active for card in self.cards)

    def _ms(self, duration: int, moves: bool = True) -> int:
        """Movement is dropped under reduced motion; colour fades get shorter."""
        if not self.reduce_motion:
            return duration
        return 0 if moves else round(duration * 0.7)

    def _draw_well(self) -> None:
        canvas, c, k = self.well, self.colors, self.scale
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 20 or h < 20:
            return
        hover = self.motion.get("hover", 0.0)
        fill = blend(c["surface"], c["drop"], hover)
        inset = round(2 * k)
        rounded_rect(
            canvas, inset, inset, w - inset, h - inset, round(14 * k), fill=fill,
            outline=blend(c["border"], c["accent"], hover), width=max(1, round((1.5 + 0.5 * hover) * k)),
            dash=() if hover >= 0.5 else (round(6 * k), round(4 * k)),
        )
        # The layout follows the well's current height, so it stays right while it resizes.
        # Content dips out and back in at the two switch points (layout, drag wording) so the
        # swap reads as one change instead of two states overlapping.
        switch = round(LAYOUT_SWITCH * k)
        compact = h < switch
        alpha = min(1.0, abs(h - switch) / (28 * k)) if "well_h" in self.motion.jobs else 1.0
        if 0.0 < hover < 1.0:
            alpha *= abs(hover - 0.5) * 2
        active = hover >= 0.5

        def ink(color: str) -> str:
            return blend(fill, color, alpha)

        fonts = self.app.fonts
        show_widgets = alpha > 0.6 and "well_h" not in self.motion.jobs  # Tk widgets can't fade: hide mid-resize
        if self.busy:
            done, total = self.progress_value
            count = f"第 {min(done + 1, total)} 個，共 {total} 個" + ("　·　可以繼續拖進來" if self.dnd else "")
            self.progress.configure(maximum=max(total, 1))
            if compact:  # one row: what is happening, then the bar
                pad = round(18 * k)
                bar = round((w - 2 * pad) * 0.45)
                canvas.create_text(pad, h / 2 - round(14 * k), anchor="w", font=fonts["strong"], fill=ink(c["text"]),
                                   text=fit_text(fonts["strong"], self.status or "正在轉換…", w - 2 * pad))
                if show_widgets:
                    canvas.create_window(pad, h / 2 + round(14 * k), anchor="w", window=self.progress, width=bar)
                canvas.create_text(pad + bar + round(12 * k), h / 2 + round(14 * k), anchor="w", font=fonts["caption"],
                                   fill=ink(c["secondary"]), text=f"{min(done + 1, total)}/{total}")
            else:
                cx, cy = w / 2, h * 0.42
                canvas.create_text(cx, cy, text=self.status or "正在轉換…", font=fonts["strong"], fill=ink(c["text"]),
                                   width=w - round(40 * k), justify="center")
                canvas.create_text(cx, cy + round(30 * k), font=fonts["caption"], fill=ink(c["secondary"]), text=count)
                if show_widgets:
                    canvas.create_window(cx, cy + round(64 * k), window=self.progress, width=w - round(64 * k))
            return
        icon_ink = ink(blend(c["secondary"], c["accent"], hover))
        nudge = 0 if self.reduce_motion else 4 * k * hover  # the arrow dips toward the page while a drag hovers
        if active:
            title, note = "放開就開始轉換", "轉好的檔案會出現在下方"
        elif self.dnd:
            title, note = ("把檔案拖到這裡", "轉好後，把檔案卡片直接拖進 Claude") if not self.cards \
                else ("再拖入檔案", "可以一次拖入好幾個")
        else:
            title, note = "選擇要轉換的檔案", "轉好後可以複製內容，或開啟檔案位置"
        if compact:  # one row: icon, words, button
            pad = round(18 * k)
            self._page_icon(canvas, pad + 13 * k, h / 2, 0.75, icon_ink, nudge * 0.75)
            text_x = pad + round(40 * k)
            canvas.create_text(text_x, h / 2 - round(10 * k), anchor="w", text=title, font=fonts["strong"], fill=ink(c["text"]))
            canvas.create_text(text_x, h / 2 + round(12 * k), anchor="w", text=note, font=fonts["caption"],
                               fill=ink(c["secondary"]))
            if not active and show_widgets:
                canvas.create_window(w - pad, h / 2, anchor="e", window=self.choose_button)
            return
        cx, cy = w / 2, h / 2 - round(46 * k)
        self._page_icon(canvas, cx, cy, 1.0, icon_ink, nudge)
        canvas.create_text(cx, cy + round(50 * k), text=title, font=fonts["strong"], fill=ink(c["text"]))
        canvas.create_text(cx, cy + round(76 * k), text=note, font=fonts["caption"], fill=ink(c["secondary"]),
                           width=w - round(32 * k), justify="center")
        if not active and show_widgets:
            canvas.create_window(cx, cy + round(114 * k), window=self.choose_button)

    def _page_icon(self, canvas, cx, cy, size: float, ink: str, nudge: float = 0.0) -> None:
        """A page with an arrow dropping in, as on the main window's empty state."""
        k = self.scale * size
        stroke = max(2, round(2 * self.scale))
        left, top, right, bottom, fold = cx - 17 * k, cy - 21 * k, cx + 17 * k, cy + 23 * k, 10 * k
        canvas.create_line(left, top, right - fold, top, right, top + fold, right, bottom, left, bottom, left, top,
                           fill=ink, width=stroke, joinstyle="round")
        canvas.create_line(right - fold, top, right - fold, top + fold, right, top + fold, fill=ink, width=stroke)
        canvas.create_line(cx, cy - 9 * k + nudge, cx, cy + 13 * k + nudge, fill=ink, width=stroke, arrow="last",
                           arrowshape=(9 * k, 9 * k, 4 * k))

    def _set_layout(self, compact: bool) -> None:
        """Big well while there is nothing to hand over; one row above the results after."""
        if compact == self._compact:
            return
        first = self._compact is None
        self._compact = compact
        k = self.scale
        instant = first or self.reduce_motion or not self.win.winfo_viewable()
        current = self.well.winfo_height()
        if compact:
            target = round(COMPACT_WELL * k)
            self.frame.rowconfigure(0, weight=0)
            self.frame.rowconfigure(2, weight=1)
            self.header.grid(row=1, column=0, sticky="ew", pady=(round(14 * k), round(8 * k)))
            self.holder.grid(row=2, column=0, sticky="nsew")
            if instant:
                self.motion.set("well_h", target)
                self.well.configure(height=target)
                return
            # Start where it is, less the room the results header now takes, so nothing below
            # (the footer) is pushed out of the window; the list then opens up underneath.
            start = max(target, current - self.header.winfo_reqheight() - round(22 * k))
            self.well.configure(height=start)
            self.motion.to("well_h", target, RESIZE_MS, EASE_IN_OUT, self._apply_well_height, start=start)
        else:
            self.header.grid_remove()
            self.holder.grid_remove()
            self.frame.rowconfigure(2, weight=0)
            if instant:
                self.frame.rowconfigure(0, weight=1)
                return
            # Grow back into the space the results took, then follow the window again.
            target = self.frame.winfo_height() - round(24 * k) - self.footer.winfo_height() - round(10 * k)
            self.frame.rowconfigure(0, weight=0)
            self.motion.to("well_h", max(target, current), RESIZE_MS, EASE_IN_OUT, self._apply_well_height,
                           start=current, on_done=lambda: self.frame.rowconfigure(0, weight=1))

    def _apply_well_height(self) -> None:
        self.well.configure(height=round(self.motion.get("well_h")))
        self._draw_well()

    def _animate_progress(self) -> None:
        done, total = self.progress_value
        self.progress.configure(maximum=max(total, 1))
        self.motion.to("progress", done, self._ms(PROGRESS_MS), EASE_OUT,
                       lambda: self.progress.configure(value=self.motion.get("progress", 0.0)))

    def _refresh(self) -> None:
        options = self.app.conversion_options()
        excel = "Excel：給 Claude 的精簡版" if options["xlsx_format"] == "compact" else "Excel：完整表格"
        if options.get("xlsx_rows_containing"):
            excel += f"（只留含「{'、'.join(options['xlsx_rows_containing'][:2])}」的列）"
        where = "存到原檔旁" if self.app.output_destination() is None else f"存到 {self.app.output_destination().name}"
        self.settings_label.configure(text=f"{excel} · {where}")
        self._set_layout(bool(self.cards))
        self._draw_well()

    # ----- scrolling ------------------------------------------------------- #
    def _update_scroll_region(self) -> None:
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def _on_list_resize(self, event) -> None:
        self.list_canvas.itemconfigure(self.list_window, width=event.width)

    def _on_scroll(self, first, last) -> None:
        self.scrollbar.set(first, last)
        if float(first) <= 0 and float(last) >= 1:
            self.scrollbar.grid_remove()
        else:
            self.scrollbar.grid(row=0, column=1, sticky="ns")

    def _on_wheel(self, event) -> None:
        if self.list_frame.winfo_height() > self.list_canvas.winfo_height():
            self.list_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    # ----- dropping in ----------------------------------------------------- #
    def _hover(self, on: bool) -> None:
        """Drop-target highlight: a colour transition, retargeted if the drag leaves early."""
        if on == self.hovering:
            return
        self.hovering = on
        duration = HOVER_IN_MS if on else HOVER_OUT_MS
        self.motion.to("hover", 1.0 if on else 0.0, self._ms(duration, moves=False), EASE, self._draw_well)

    def _on_drag_enter(self, event):
        self._hover(True)
        return event.action

    def _on_drag_leave(self, event):
        self._hover(False)
        return event.action

    def _on_drop(self, event):
        self._hover(False)
        self.add(self.win.tk.splitlist(event.data))
        return event.action

    def choose_files(self) -> None:
        paths = self.app.filedialog.askopenfilenames(parent=self.win, title="選擇要轉換的檔案")
        if paths:
            self.add(paths)

    def add(self, paths, extra: Optional[dict] = None) -> None:
        """Convert files; ``extra`` adds options for this batch (e.g. only the newest versions)."""
        values = [str(p) for p in paths if str(p).strip()]
        if not values:
            return
        if self.busy:
            self.pending.append((values, dict(extra or {})))
            self.progress_value = (self.progress_value[0], self.progress_value[1] + len(convert.collect_sources(values, recursive=True)))
            self._animate_progress()
            self._draw_well()
            return
        self._start(values, extra)

    def convert_latest(self, source: str, n: int) -> None:
        """Re-convert a workbook keeping only the newest ``n`` versions: a small file to hand over."""
        self.add([source], {"xlsx_latest_versions": n, "output_suffix": f"_最新{n}版",
                            "variant": f"只含最新 {n} 版"})

    def _options(self, extra: Optional[dict]) -> Tuple[dict, str]:
        options = dict(self.app.conversion_options())
        extra = dict(extra or {})
        variant = extra.pop("variant", "")
        options.update(extra)
        return options, variant

    def _start(self, values: List[str], extra: Optional[dict] = None) -> None:
        sources = convert.collect_sources(values, recursive=self.app.recursive.get())
        if not sources:
            self._add_card(PanelResult(values[0], error="沒有可以轉換的檔案", error_type="UnsupportedFormatException"))
            return
        self.busy = True
        self.status = "正在準備轉換引擎…" if self.app.converter is None else f"正在轉換 {Path(sources[0].value).name}"
        self.progress_value = (0, len(sources))
        self.motion.set("progress", 0.0)
        self.progress.configure(value=0)
        self._draw_well()
        options, variant = self._options(extra)
        args = (sources, self.app.output_destination(), self.app.keep_data_uris.get(), options, variant)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, sources, out_dir, keep_data_uris, options, variant="") -> None:
        try:
            converter = self.app._load_converter()
            if converter is None:
                self.events.put(("failed", self.app.converter_error or "無法啟動轉換引擎"))
                return
            convert.convert_batch(
                converter, sources, out_dir, keep_data_uris=keep_data_uris, options=options,
                on_start=lambda n, t, s: self.events.put(("working", n, t, Path(s.value).name)),
                on_result=lambda n, t, r: self.events.put(("result", n, measure(r, variant))),
                progress=lambda message: self.events.put(("status", message)),
            )
        except Exception as exc:
            self.events.put(("failed", convert.describe_error(exc)))
        finally:
            self.events.put(("finished",))

    def _poll(self) -> None:
        try:
            while True:
                self._handle(self.events.get_nowait())
        except queue.Empty:
            pass
        if self.win.winfo_exists():
            self.win.after(80, self._poll)

    def _handle(self, event: tuple) -> None:
        kind = event[0]
        offset = self._offset
        if kind == "working":
            _, n, _total, name = event
            self.progress_value = (offset + n - 1, self.progress_value[1])
            self.status = f"正在轉換 {name}"
            self._animate_progress()
        elif kind == "status":
            self.status = event[1]
        elif kind == "result":
            self.progress_value = (offset + event[1], self.progress_value[1])
            self._animate_progress()
            self._add_card(event[2])
        elif kind == "failed":
            self._add_card(PanelResult("轉換引擎", error=event[1]))
        elif kind == "finished":
            if self.pending:
                values, extra = self.pending.pop(0)
                self._offset = self.progress_value[0]
                sources = convert.collect_sources(values, recursive=self.app.recursive.get())
                options, variant = self._options(extra)
                args = (sources, self.app.output_destination(), self.app.keep_data_uris.get(), options, variant)
                threading.Thread(target=self._worker, args=args, daemon=True).start()
            else:
                self.busy = False
                self._offset = 0
                self.status = ""
        self._draw_well()

    # ----- results --------------------------------------------------------- #
    def _add_card(self, result: PanelResult) -> None:
        card = ResultCard(self, result)
        card.motion.set("enter", 0.0)
        if self.cards:
            card.canvas.pack(fill="x", pady=(0, round(8 * self.scale)), before=self.cards[0].canvas)
        else:
            card.canvas.pack(fill="x", pady=(0, round(8 * self.scale)))
        self.cards.insert(0, card)
        for old in self.cards[MAX_CARDS:]:
            old.canvas.destroy()
        del self.cards[MAX_CARDS:]
        self.list_canvas.yview_moveto(0)
        if result.ok:
            card.canvas.focus_set()
        self._refresh()
        # Cards that arrive together come in one after another, never blocking the list.
        now = time.perf_counter()
        delay = max(0.0, self._next_enter - now)
        self._next_enter = now + delay + STAGGER_MS / 1000
        card.motion.to("enter", 1.0, self._ms(ENTER_MS, moves=False), EASE_OUT, card.draw, delay_ms=round(delay * 1000))

    def clear(self) -> None:
        for card in self.cards:
            card.canvas.destroy()
        self.cards = []
        self._refresh()

    def _copy_focused(self) -> None:
        focus = self.win.focus_get()
        for card in self.cards:
            if card.canvas is focus:
                card.copy()
                return
        if self.cards:
            next((c for c in self.cards if c.result.ok), self.cards[0]).copy()

    def copy_instructions(self) -> None:
        """The standing instructions for a Claude Project: index first, data files on request."""
        self.win.clipboard_clear()
        self.win.clipboard_append(PROJECT_INSTRUCTIONS)
        label = self.settings_label
        previous = label.cget("text")
        label.configure(text="✓ 已複製。貼到 Claude 專案的「指示」欄位",
                        foreground=self.colors["success"])
        self.win.after(4000, lambda: label.winfo_exists() and label.configure(text=previous, foreground=""))

    # ----- window ---------------------------------------------------------- #
    def _apply_pin(self) -> None:
        self.win.attributes("-topmost", bool(self.pinned.get()))

    def show(self) -> None:
        self._refresh()
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()

    def show_main_window(self) -> None:
        self.standalone = False
        self.app.root.deiconify()
        self.app.root.lift()
        try:
            self.app.options_button.focus_set()
        except Exception:
            pass

    def close(self) -> None:
        if self.standalone or self.app.root.state() == "withdrawn":
            self.app._on_close()  # the panel was the only window
        else:
            self.win.withdraw()


class ResultCard:
    """One result: a file card to drag into Claude, or an error with the fix."""

    def __init__(self, panel: QuickPanel, result: PanelResult):
        self.panel, self.result = panel, result
        self.hover = False
        self._width = 0
        tk = panel.tk
        self.canvas = tk.Canvas(panel.list_frame, height=round(92 * panel.scale), highlightthickness=0,
                                borderwidth=0, takefocus=1, cursor="hand2" if result.ok else "")
        c = self.canvas
        self.motion = Motion(c)
        c.bind("<Configure>", self._on_configure)
        c.bind("<Enter>", lambda _e: self._set_hover(True))
        c.bind("<Leave>", lambda _e: self._set_hover(False))
        c.bind("<FocusIn>", lambda _e: self.draw())
        c.bind("<FocusOut>", lambda _e: self.draw())
        c.bind("<ButtonPress-1>", lambda _e: (c.focus_set(), self._press(True)))
        c.bind("<ButtonRelease-1>", lambda _e: self._press(False))
        c.bind("<Double-1>", lambda _e: self.open())
        c.bind("<Return>", lambda _e: self.open())
        c.bind("<Button-3>", self._menu)
        c.bind("<Shift-F10>", lambda e: self._menu(e, keyboard=True))
        c.bind("<MouseWheel>", panel._on_wheel)
        c.bind("<Button-4>", lambda _e: panel.list_canvas.yview_scroll(-1, "units"))
        c.bind("<Button-5>", lambda _e: panel.list_canvas.yview_scroll(1, "units"))
        if panel.dnd is not None and result.ok:
            c.drag_source_register(1, panel.dnd.DND_FILES)
            c.dnd_bind("<<DragInitCmd>>", self._drag_init)
            c.dnd_bind("<<DragEndCmd>>", lambda e: (self._press(False), e.action)[1])

    # ----- drawing --------------------------------------------------------- #
    def _on_configure(self, event) -> None:
        if event.width != self._width:  # height changes come from draw() itself
            self._width = event.width
            self.draw()

    def _set_hover(self, value: bool) -> None:
        # Tens of times a day: an instant change, no transition.
        self.hover = value
        self.draw()

    def _press(self, down: bool) -> None:
        """Press feedback: the card gives a little under the pointer, then springs back."""
        if not self.result.ok:
            return
        duration = PRESS_IN_MS if down else PRESS_OUT_MS
        self.motion.to("press", 1.0 if down else 0.0, self.panel._ms(duration, moves=False), EASE_OUT, self.draw)

    def draw(self) -> None:
        c, colors, k, fonts = self.canvas, self.panel.colors, self.panel.scale, self.panel.app.fonts
        c.delete("all")
        bg = colors["bg"]
        c.configure(background=bg)
        w = c.winfo_width()
        if w < 40:
            return
        r = self.result
        moves = not self.panel.reduce_motion
        enter = self.motion.get("enter", 1.0)
        press = self.motion.get("press", 0.0)
        rise = (1 - enter) * ENTER_RISE * k if moves else 0.0

        def ink(color: str) -> str:  # fades in with the card
            return blend(bg, color, enter)

        focused = c.focus_displayof() is c if c.winfo_exists() else False
        raised = (self.hover or focused) and r.ok
        pad = round(14 * k)
        icon_w, icon_h = round(30 * k), round(38 * k)
        text_x = pad + icon_w + round(14 * k)
        text_w = w - text_x - pad
        items = []
        y = pad + rise
        if r.ok:
            # The token estimate is what people come for: it gets the text colour; the
            # verification mark gets the success colour; everything else stays quiet.
            items.append(c.create_text(text_x, y, text=fit_text(fonts["strong"], r.output.name, text_w),
                                       font=fonts["strong"], fill=ink(colors["text"]), anchor="nw"))
            y = c.bbox(items[-1])[3] + round(3 * k)
            items.append(c.create_text(text_x, y, text=describe_tokens(r.tokens), font=fonts["body"],
                                       fill=ink(colors["text"]), anchor="nw"))
            if r.excel:
                x = c.bbox(items[-1])[2] + round(8 * k)
                c.create_text(x, y, text="✓ 已逐格核對", font=fonts["body"], fill=ink(colors["success"]), anchor="nw")
            y = c.bbox(items[-1])[3] + round(3 * k)
            lines = []
            if r.data_files:  # a split workbook: the index is what to hand over first
                lines.append(("先拖這張就好，Claude 會說還要哪個檔", colors["text"]))
                lines.append((f"同資料夾另有 {r.data_files} 個資料檔，共{describe_tokens(r.data_tokens)}", colors["secondary"]))
            if r.total_tokens > CONVERSATION_BUDGET:
                lines.append((f"⚠ 全部{describe_tokens(r.total_tokens)}，太大，別整包上傳", colors["caution"]))
            for text, color in lines:
                items.append(c.create_text(text_x, y, text=fit_text(fonts["caption"], text, text_w),
                                           font=fonts["caption"], fill=ink(color), anchor="nw"))
                y = c.bbox(items[-1])[3] + round(3 * k)
            # One-click extracts for waterfall-style workbooks: the newest version(s) only.
            for n, estimate in sorted(r.offers.items()):
                label = ("只轉最新一版" if n == 1 else f"只轉最新 {n} 版") + f" · {describe_tokens(estimate)}"
                tag = f"offer{n}"
                items.append(c.create_text(text_x, y, text="→ " + fit_text(self.panel.link_font, label, text_w - round(16 * k)),
                                           font=self.panel.link_font, fill=ink(colors["accent"]), anchor="nw", tags=(tag,)))
                c.tag_bind(tag, "<ButtonRelease-1>", lambda _e, n=n: self.panel.convert_latest(r.source, n))
                y = c.bbox(items[-1])[3] + round(4 * k)
            source = f"來自 {Path(r.source).name}" + (f" · {r.variant}" if r.variant else "")
            items.append(c.create_text(text_x, y, text=fit_text(fonts["caption"], source, text_w),
                                       font=fonts["caption"], fill=ink(colors["secondary"]), anchor="nw"))
            y = c.bbox(items[-1])[3] + round(3 * k)
        else:
            from gui import friendly_error

            what, fix = friendly_error(r.error_type)
            items.append(c.create_text(text_x, y, text=fit_text(fonts["strong"], f"{Path(r.source).name}：{what}", text_w),
                                       font=fonts["strong"], fill=ink(colors["danger"]), anchor="nw"))
            y = c.bbox(items[-1])[3] + round(3 * k)
            items.append(c.create_text(text_x, y, text=fix, font=fonts["body"], fill=ink(colors["secondary"]),
                                       anchor="nw", width=text_w))
            y = c.bbox(items[-1])[3] + round(3 * k)
        natural = round(max(y - rise + pad - round(3 * k), pad * 2 + icon_h))
        # Entering cards open up their slot as they fade in, so the list below slides rather than jumps.
        height = max(1, round(natural * enter)) if moves else natural
        shadow_gap = round(2 * k)
        if int(c.cget("height")) != height + shadow_gap:
            c.configure(height=height + shadow_gap)
        # Card body: hover and focus raise it (a soft shadow) and give it the accent edge.
        surface = blend(colors["surface"], colors["border"], 0.45 * press)  # pressed: a touch darker
        body_bottom = natural - 1
        if raised:
            shade = blend(bg, "#000000", 0.35 if self.panel.app.appearance == "dark" else 0.07)
            shadow = rounded_rect(c, 2, 1 + shadow_gap + rise, w - 2, body_bottom + shadow_gap + rise, round(10 * k),
                                  fill=ink(shade), outline="")
            c.tag_lower(shadow)
        edge = colors["accent"] if raised else colors["border"]
        body = rounded_rect(c, 1, 1 + rise, w - 1, body_bottom + rise, round(10 * k), fill=ink(surface),
                            outline=ink(edge), width=max(1, round((2 if focused else 1) * k)))
        c.tag_lower(body)
        if raised:
            c.tag_lower(shadow)
        # The file icon: a page with a folded corner, labelled MD; or an alert mark.
        ix, iy = pad, (natural - icon_h) / 2 + rise
        icon = ink(colors["accent"] if r.ok else colors["danger"])
        fold, stroke = round(9 * k), max(1, round(1.5 * k))
        c.create_line(ix, iy, ix + icon_w - fold, iy, ix + icon_w, iy + fold, ix + icon_w, iy + icon_h, ix, iy + icon_h,
                      ix, iy, fill=icon, width=stroke, joinstyle="round")
        c.create_line(ix + icon_w - fold, iy, ix + icon_w - fold, iy + fold, ix + icon_w, iy + fold, fill=icon, width=stroke)
        c.create_text(ix + icon_w / 2, iy + icon_h * 0.62, text="MD" if r.ok else "!", fill=icon,
                      font=fonts["caption"] if r.ok else fonts["strong"])
        if press and moves:  # scale(0.97) about the card's centre
            factor = 1 - (1 - PRESS_SCALE) * press
            c.scale("all", w / 2, natural / 2 + rise, factor, factor)

    # ----- actions --------------------------------------------------------- #
    def _drag_init(self, _event):
        dnd = self.panel.dnd
        return (dnd.COPY, dnd.DND_FILES, (str(self.result.output),))

    def open(self) -> None:
        if self.result.ok:
            from gui import open_path

            open_path(self.result.output)

    def reveal(self) -> None:
        if self.result.ok:
            from gui import reveal_path

            reveal_path(self.result.output)

    def copy(self) -> None:
        if not self.result.ok:
            return
        text = _read(self.result.output)
        win = self.panel.win
        win.clipboard_clear()
        win.clipboard_append(text)
        # Often keyboard-initiated (Ctrl+Shift+C): an instant state change, no motion.
        label = self.panel.results_label
        label.configure(text=f"✓ 已複製 {self.result.output.name}", foreground=self.panel.colors["success"])
        win.after(2500, lambda: label.winfo_exists() and label.configure(text="轉換結果", foreground=""))

    def copy_path(self) -> None:
        if self.result.ok:
            win = self.panel.win
            win.clipboard_clear()
            win.clipboard_append(str(self.result.output))

    def _menu(self, event, keyboard: bool = False) -> None:
        tk = self.panel.tk
        menu = tk.Menu(self.canvas, tearoff=False)
        state = "normal" if self.result.ok else "disabled"
        menu.add_command(label="開啟", command=self.open, state=state)
        menu.add_command(label="開啟檔案位置", command=self.reveal, state=state)
        menu.add_separator()
        menu.add_command(label="複製內容", accelerator="Ctrl+Shift+C", command=self.copy, state=state)
        menu.add_command(label="複製檔案路徑", command=self.copy_path, state=state)
        if self.result.offers:
            menu.add_separator()
            for n in sorted(self.result.offers):
                menu.add_command(label="只轉最新一版" if n == 1 else f"只轉最新 {n} 版",
                                 command=lambda n=n: self.panel.convert_latest(self.result.source, n))
        if keyboard:
            x, y = self.canvas.winfo_rootx() + 20, self.canvas.winfo_rooty() + 20
        else:
            x, y = event.x_root, event.y_root
        menu.tk_popup(x, y)
