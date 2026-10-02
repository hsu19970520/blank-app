"""Recalculate a workbook with the user's own Microsoft Excel before converting it.

Excel is the only engine that reproduces every function (XLOOKUP, OFFSET,
dynamic arrays, ...) and can refresh pivot tables, so instead of guessing we
drive it through COM:

    copy the file -> open the copy in a private, hidden Excel instance
    -> recalculate -> refresh every pivot cache -> recalculate again
       (formulas that read pivot output, e.g. Summary, see the new pivots)
    -> save the copy -> quit Excel

The original file is never opened by Excel and any Excel window the user has
open is left alone (DispatchEx starts a separate Excel process). Requires
Windows, Microsoft Excel and pywin32.
"""

from __future__ import annotations

import importlib.util
import re
import shutil
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional, Tuple

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # never run macros in the copy
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


class RecalcError(RuntimeError):
    """Recalculation failed. The message is meant for the user (Traditional Chinese)."""


@dataclass
class RecalcResult:
    path: Path  # the recalculated copy
    workdir: Path  # temporary folder holding the copy; delete it when done
    excel_version: str
    pivot_tables: int
    pivot_failures: List[str] = field(default_factory=list)
    seconds: float = 0.0
    new_name_errors: int = 0
    finished_at: str = ""

    def note(self) -> str:
        """One sentence for the top of the Markdown."""
        version = f" {self.excel_version}" if self.excel_version else ""
        text = f"已用 Microsoft Excel{version} 在轉換前重新計算所有公式（{self.finished_at}，耗時 {self.seconds:.0f} 秒）"
        if self.pivot_tables:
            text += f"，並重新整理 {self.pivot_tables} 個樞紐分析表"
        if self.pivot_failures:
            text += f"；{len(self.pivot_failures)} 個樞紐分析表資料無法重新整理，仍是舊的內容"
        if self.new_name_errors:
            text += (f"；⚠️ 重新計算後多出 {self.new_name_errors:,} 個 #NAME? 錯誤，"
                     "可能是這台電腦的 Excel 版本不支援檔案中的函數（例如 XLOOKUP），相關數字不可信")
        return text


def excel_status() -> Tuple[bool, str]:
    """(available, reason) without starting Excel."""
    if sys.platform != "win32":
        return False, "重新計算需要 Windows 與 Microsoft Excel"
    if importlib.util.find_spec("win32com") is None:
        return False, "缺少 pywin32 元件，無法控制 Excel"
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
    except OSError:
        return False, "這台電腦沒有安裝 Microsoft Excel"
    return True, "可以用 Microsoft Excel 重新計算"


def count_name_errors(path: Path) -> int:
    """Cells whose cached value is #NAME? (an unknown function or name)."""
    total = 0
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                with archive.open(name) as fh:
                    total += len(re.findall(rb"<v>#NAME\?</v>", fh.read()))
    return total


def _com_message(exc: BaseException) -> str:
    """The readable part of a pywintypes.com_error (or any exception)."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and isinstance(args[1], str):
        return args[1]
    return str(exc) or exc.__class__.__name__


def recalculate(
    path,
    refresh_pivots: bool = True,
    progress: Optional[Callable[[str], None]] = None,
    dispatch: Optional[Callable[[str], object]] = None,
) -> RecalcResult:
    """Recalculate a copy of ``path`` with Excel and return where the copy is.

    ``dispatch`` replaces win32com.client.DispatchEx (used by the tests)."""
    progress = progress or (lambda message: None)
    source = Path(path)
    if source.suffix.lower() not in EXCEL_EXTENSIONS:
        raise RecalcError(f"只有 .xlsx / .xlsm 可以重新計算（{source.name}）")
    com = None
    if dispatch is None:
        available, reason = excel_status()
        if not available:
            raise RecalcError(reason)
        import pythoncom
        import win32com.client

        com = pythoncom
        dispatch = win32com.client.DispatchEx

    workdir = Path(tempfile.mkdtemp(prefix="markitdown-recalc-"))
    copy = workdir / source.name
    shutil.copyfile(source, copy)
    started = time.time()
    excel = workbook = None
    pivot_tables, failures, version = 0, [], ""
    error: Optional[BaseException] = None
    if com is not None:
        com.CoInitialize()  # needed on every thread that talks to COM (the GUI converts on a worker thread)
    try:
        progress("正在啟動 Excel…")
        excel = dispatch("Excel.Application")
        for name, value in (
            ("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False), ("EnableEvents", False),
            ("AskToUpdateLinks", False), ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        ):
            try:
                setattr(excel, name, value)
            except Exception:
                pass
        version = str(getattr(excel, "Version", "") or "")
        progress("正在用 Excel 開啟檔案的副本…")
        workbook = excel.Workbooks.Open(
            str(copy), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False, Notify=False
        )
        progress("正在重新計算所有公式…")
        excel.CalculateFull()
        if refresh_pivots:
            caches = workbook.PivotCaches()
            for i in range(1, int(caches.Count) + 1):
                progress(f"正在重新整理樞紐分析表的資料（{i}/{int(caches.Count)}）…")
                try:
                    caches.Item(i).Refresh()
                except Exception as exc:  # e.g. a source range that no longer exists
                    failures.append(_com_message(exc))
            pivot_tables = sum(int(sheet.PivotTables().Count) for sheet in workbook.Worksheets)
            if pivot_tables:
                progress("正在用新的樞紐分析表結果再算一次…")
                excel.CalculateFull()
        progress("正在儲存重新計算後的副本…")
        workbook.Save()
    except Exception as exc:
        error = exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        workbook = excel = None
        if com is not None:
            com.CoUninitialize()
    if error is not None:  # only now is the copy no longer held open by Excel
        shutil.rmtree(workdir, ignore_errors=True)
        raise RecalcError(f"Excel 重新計算失敗：{_com_message(error)}") from error

    progress("正在檢查重新計算的結果…")
    try:
        new_errors = max(0, count_name_errors(copy) - count_name_errors(source))
    except (OSError, zipfile.BadZipFile):
        new_errors = 0
    return RecalcResult(
        path=copy,
        workdir=workdir,
        excel_version=version,
        pivot_tables=pivot_tables,
        pivot_failures=failures,
        seconds=time.time() - started,
        new_name_errors=new_errors,
        finished_at=time.strftime("%Y-%m-%d %H:%M"),
    )
