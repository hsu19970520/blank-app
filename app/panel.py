"""Quick Convert: a small floating panel between Explorer and Claude.

Drop files on the well; each result becomes a file card that can be dragged
straight into Claude Desktop (or anywhere that takes files). One job, three
states: waiting for a drop, converting, results to hand over. Every drag has a
non-drag alternative: "選擇檔案…" in, and copy / open / show-in-folder out.
"""

from __future__ import annotations

import queue
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import convert
from xlsx_compact import INDEX_NAME, estimate_tokens

PANEL_TITLE = "快速轉換"
MAX_CARDS = 30


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


def measure(result: "convert.Result") -> PanelResult:
    """Token estimates for a conversion result (runs on the worker thread)."""
    source = result.source.value
    excel = convert.is_excel_file(result.source)
    if result.error or result.output is None:
        return PanelResult(source, error=result.error or "沒有輸出", error_type=result.error_type, excel=excel)
    output = Path(result.output)
    text = result.markdown if result.markdown is not None else _read(output)
    item = PanelResult(source, output, estimate_tokens(text), excel=excel)
    if output.name == INDEX_NAME:
        for data in sorted(output.parent.glob("[0-9][0-9]_*.md")):
            item.data_files += 1
            item.data_tokens += estimate_tokens(_read(data))
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
        self.pending: List[str] = []  # dropped while busy: converted next
        self.busy = False
        self.hovering = False
        self.status = ""
        self.progress_value = (0, 0)  # (finished, total) across everything dropped since idle
        self._offset = 0  # files finished before the batch now running
        self.cards: List[ResultCard] = []
        self.pinned = tk.BooleanVar(value=True)

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
        footer.grid(row=3, column=0, sticky="ew", pady=(round(10 * k), 0))
        footer.columnconfigure(0, weight=1)
        self.settings_label = ttk.Label(footer, style="Secondary.TLabel")
        self.settings_label.grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="選項…", command=self.show_main_window).grid(row=0, column=1, sticky="e")
        self.pin_switch = ttk.Checkbutton(footer, text="浮在其他視窗上方", variable=self.pinned, command=self._apply_pin,
                                          style="Switch.TCheckbutton" if self.app.sv_ttk else "TCheckbutton")
        self.pin_switch.grid(row=1, column=0, columnspan=2, sticky="w", pady=(round(6 * k), 0))

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

    def _draw_well(self) -> None:
        canvas, c, k = self.well, self.colors, self.scale
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 20 or h < 20:
            return
        active = self.hovering
        inset = round(2 * k)
        rounded_rect(
            canvas, inset, inset, w - inset, h - inset, round(14 * k),
            fill=c["drop"] if active else c["surface"],
            outline=c["accent"] if active else c["border"], width=max(1, round((2 if active else 1.5) * k)),
            dash=() if active else (round(6 * k), round(4 * k)),
        )
        fonts = self.app.fonts
        if self.busy:
            done, total = self.progress_value
            count = f"第 {min(done + 1, total)} 個，共 {total} 個" + ("　·　可以繼續拖進來" if self.dnd else "")
            self.progress.configure(maximum=max(total, 1), value=done)
            if self.cards:  # one row: what is happening, then the bar
                pad = round(18 * k)
                canvas.create_text(pad, h / 2 - round(14 * k), anchor="w", font=fonts["strong"], fill=c["text"],
                                   text=fit_text(fonts["strong"], self.status or "正在轉換…", w - 2 * pad))
                canvas.create_window(pad, h / 2 + round(14 * k), anchor="w", window=self.progress,
                                     width=round((w - 2 * pad) * 0.45))
                canvas.create_text(pad + round((w - 2 * pad) * 0.45) + round(12 * k), h / 2 + round(14 * k),
                                   anchor="w", font=fonts["caption"], fill=c["secondary"], text=f"{min(done + 1, total)}/{total}")
            else:
                cx, cy = w / 2, h * 0.42
                canvas.create_text(cx, cy, text=self.status or "正在轉換…", font=fonts["strong"], fill=c["text"],
                                   width=w - round(40 * k), justify="center")
                canvas.create_text(cx, cy + round(30 * k), font=fonts["caption"], fill=c["secondary"], text=count)
                canvas.create_window(cx, cy + round(64 * k), window=self.progress, width=w - round(64 * k))
            return
        ink = c["accent"] if active else c["secondary"]
        if active:
            title, note = "放開就開始轉換", "轉好的檔案會出現在下方"
        elif self.dnd:
            title, note = ("把檔案拖到這裡", "轉好後，把檔案卡片直接拖進 Claude") if not self.cards \
                else ("再拖入檔案", "可以一次拖入好幾個")
        else:
            title, note = "選擇要轉換的檔案", "轉好後可以複製內容，或開啟檔案位置"
        if self.cards:  # one row: icon, words, button
            pad = round(18 * k)
            self._page_icon(canvas, pad + 13 * k, h / 2, 0.75, ink)
            text_x = pad + round(40 * k)
            canvas.create_text(text_x, h / 2 - round(10 * k), anchor="w", text=title, font=fonts["strong"], fill=c["text"])
            canvas.create_text(text_x, h / 2 + round(12 * k), anchor="w", text=note, font=fonts["caption"], fill=c["secondary"])
            if not active:
                canvas.create_window(w - pad, h / 2, anchor="e", window=self.choose_button)
            return
        cx, cy = w / 2, h / 2 - round(46 * k)
        self._page_icon(canvas, cx, cy, 1.0, ink)
        canvas.create_text(cx, cy + round(50 * k), text=title, font=fonts["strong"], fill=c["text"])
        canvas.create_text(cx, cy + round(76 * k), text=note, font=fonts["caption"], fill=c["secondary"],
                           width=w - round(32 * k), justify="center")
        if not active:
            canvas.create_window(cx, cy + round(114 * k), window=self.choose_button)

    def _page_icon(self, canvas, cx, cy, size: float, ink: str) -> None:
        """A page with an arrow dropping in, as on the main window's empty state."""
        k = self.scale * size
        stroke = max(2, round(2 * self.scale))
        left, top, right, bottom, fold = cx - 17 * k, cy - 21 * k, cx + 17 * k, cy + 23 * k, 10 * k
        canvas.create_line(left, top, right - fold, top, right, top + fold, right, bottom, left, bottom, left, top,
                           fill=ink, width=stroke, joinstyle="round")
        canvas.create_line(right - fold, top, right - fold, top + fold, right, top + fold, fill=ink, width=stroke)
        canvas.create_line(cx, cy - 9 * k, cx, cy + 13 * k, fill=ink, width=stroke, arrow="last",
                           arrowshape=(9 * k, 9 * k, 4 * k))

    def _refresh(self) -> None:
        options = self.app.conversion_options()
        excel = "Excel：給 Claude 的精簡版" if options["xlsx_format"] == "compact" else "Excel：完整表格"
        if options.get("xlsx_rows_containing"):
            excel += f"（只留含「{'、'.join(options['xlsx_rows_containing'][:2])}」的列）"
        where = "存到原檔旁" if self.app.output_destination() is None else f"存到 {self.app.output_destination().name}"
        self.settings_label.configure(text=f"{excel} · {where}")
        # Progressive disclosure: a big well until there is something to hand over,
        # then a one-row well above the results.
        k = self.scale
        if self.cards:
            self.frame.rowconfigure(0, weight=0)
            self.frame.rowconfigure(2, weight=1)
            self.well.configure(height=round(76 * k))
            self.header.grid(row=1, column=0, sticky="ew", pady=(round(14 * k), round(8 * k)))
            self.holder.grid(row=2, column=0, sticky="nsew")
        else:
            self.frame.rowconfigure(0, weight=1)
            self.frame.rowconfigure(2, weight=0)
            self.header.grid_remove()
            self.holder.grid_remove()
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
    def _on_drag_enter(self, event):
        self.hovering = True
        self._draw_well()
        return event.action

    def _on_drag_leave(self, event):
        self.hovering = False
        self._draw_well()
        return event.action

    def _on_drop(self, event):
        self.hovering = False
        self.add(self.win.tk.splitlist(event.data))
        self._draw_well()
        return event.action

    def choose_files(self) -> None:
        paths = self.app.filedialog.askopenfilenames(parent=self.win, title="選擇要轉換的檔案")
        if paths:
            self.add(paths)

    def add(self, paths) -> None:
        values = [str(p) for p in paths if str(p).strip()]
        if not values:
            return
        if self.busy:
            self.pending.extend(values)
            self.progress_value = (self.progress_value[0], self.progress_value[1] + len(convert.collect_sources(values, recursive=True)))
            self._draw_well()
            return
        self._start(values)

    def _start(self, values: List[str]) -> None:
        sources = convert.collect_sources(values, recursive=self.app.recursive.get())
        if not sources:
            self._add_card(PanelResult(values[0], error="沒有可以轉換的檔案", error_type="UnsupportedFormatException"))
            return
        self.busy = True
        self.status = "正在準備轉換引擎…" if self.app.converter is None else f"正在轉換 {Path(sources[0].value).name}"
        self.progress_value = (0, len(sources))
        self._draw_well()
        args = (sources, self.app.output_destination(), self.app.keep_data_uris.get(), self.app.conversion_options())
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, sources, out_dir, keep_data_uris, options) -> None:
        try:
            converter = self.app._load_converter()
            if converter is None:
                self.events.put(("failed", self.app.converter_error or "無法啟動轉換引擎"))
                return
            convert.convert_batch(
                converter, sources, out_dir, keep_data_uris=keep_data_uris, options=options,
                on_start=lambda n, t, s: self.events.put(("working", n, t, Path(s.value).name)),
                on_result=lambda n, t, r: self.events.put(("result", n, measure(r))),
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
        elif kind == "status":
            self.status = event[1]
        elif kind == "result":
            self.progress_value = (offset + event[1], self.progress_value[1])
            self._add_card(event[2])
        elif kind == "failed":
            self._add_card(PanelResult("轉換引擎", error=event[1]))
        elif kind == "finished":
            if self.pending:
                values, self.pending = self.pending, []
                self._offset = self.progress_value[0]
                sources = convert.collect_sources(values, recursive=self.app.recursive.get())
                args = (sources, self.app.output_destination(), self.app.keep_data_uris.get(), self.app.conversion_options())
                threading.Thread(target=self._worker, args=args, daemon=True).start()
            else:
                self.busy = False
                self._offset = 0
                self.status = ""
        self._draw_well()

    # ----- results --------------------------------------------------------- #
    def _add_card(self, result: PanelResult) -> None:
        card = ResultCard(self, result)
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
        tk = panel.tk
        self.canvas = tk.Canvas(panel.list_frame, height=round(92 * panel.scale), highlightthickness=0,
                                borderwidth=0, takefocus=1, cursor="hand2" if result.ok else "")
        c = self.canvas
        c.bind("<Configure>", lambda _e: self.draw())
        c.bind("<Enter>", lambda _e: self._set_hover(True))
        c.bind("<Leave>", lambda _e: self._set_hover(False))
        c.bind("<FocusIn>", lambda _e: self.draw())
        c.bind("<FocusOut>", lambda _e: self.draw())
        c.bind("<Button-1>", lambda _e: c.focus_set())
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

    # ----- drawing --------------------------------------------------------- #
    def _set_hover(self, value: bool) -> None:
        self.hover = value
        self.draw()

    def draw(self) -> None:
        c, colors, k, fonts = self.canvas, self.panel.colors, self.panel.scale, self.panel.app.fonts
        c.delete("all")
        c.configure(background=colors["bg"])
        w = c.winfo_width()
        if w < 40:
            return
        r = self.result
        focused = c.focus_displayof() is c if c.winfo_exists() else False
        pad = round(14 * k)
        icon_w, icon_h = round(30 * k), round(38 * k)
        text_x = pad + icon_w + round(14 * k)
        text_w = w - text_x - pad
        items = []
        if r.ok:
            detail = describe_tokens(r.tokens) + (" · 已逐格核對" if r.excel else "")
            lines = [(r.output.name, fonts["strong"], colors["text"]), (detail, fonts["body"], colors["secondary"])]
            if r.data_files:  # a split workbook: the index is what to hand over first
                lines.append((f"同資料夾另有 {r.data_files} 個資料檔，共{describe_tokens(r.data_tokens)}",
                              fonts["caption"], colors["secondary"]))
            lines.append((f"來自 {Path(r.source).name}", fonts["caption"], colors["secondary"]))
        else:
            from gui import friendly_error

            what, fix = friendly_error(r.error_type)
            lines = [(f"{Path(r.source).name}：{what}", fonts["strong"], colors["danger"]),
                     (fix, fonts["body"], colors["secondary"])]
        y = pad
        for i, (text, font, color) in enumerate(lines):
            if r.ok or i == 0:
                text = fit_text(font, text, text_w)
                items.append(c.create_text(text_x, y, text=text, font=font, fill=color, anchor="nw"))
            else:
                items.append(c.create_text(text_x, y, text=text, font=font, fill=color, anchor="nw", width=text_w))
            y = c.bbox(items[-1])[3] + round(3 * k)
        height = max(y + pad - round(3 * k), pad * 2 + icon_h)
        if int(c.cget("height")) != height:
            c.configure(height=height)
        # Card body under the text; the border carries hover, focus and the drag hint.
        edge = colors["accent"] if (self.hover or focused) and r.ok else colors["border"]
        body = rounded_rect(c, 1, 1, w - 1, height - 1, round(10 * k), fill=colors["surface"], outline=edge,
                            width=max(1, round((2 if focused else 1) * k)))
        c.tag_lower(body)
        # The file icon: a page with a folded corner, labelled MD; or an alert mark.
        ix, iy = pad, (height - icon_h) / 2
        ink = colors["accent"] if r.ok else colors["danger"]
        fold, stroke = round(9 * k), max(1, round(1.5 * k))
        c.create_line(ix, iy, ix + icon_w - fold, iy, ix + icon_w, iy + fold, ix + icon_w, iy + icon_h, ix, iy + icon_h,
                      ix, iy, fill=ink, width=stroke, joinstyle="round")
        c.create_line(ix + icon_w - fold, iy, ix + icon_w - fold, iy + fold, ix + icon_w, iy + fold, fill=ink, width=stroke)
        c.create_text(ix + icon_w / 2, iy + icon_h * 0.62, text="MD" if r.ok else "!", fill=ink,
                      font=fonts["caption"] if r.ok else fonts["strong"])

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
        self.panel.status = ""
        self.panel.results_label.configure(text=f"已複製 {self.result.output.name}")
        win.after(2500, lambda: self.panel.results_label.configure(text="轉換結果"))

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
        if keyboard:
            x, y = self.canvas.winfo_rootx() + 20, self.canvas.winfo_rooty() + 20
        else:
            x, y = event.x_root, event.y_root
        menu.tk_popup(x, y)
