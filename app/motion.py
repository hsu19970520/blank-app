"""Motion for the Tk interface: easing curves, colour blending, retargetable tweens.

Tk has no compositor, opacity or transforms, so the usual web tools translate:
opacity becomes a colour blended with the background, transform becomes moved
or scaled canvas coordinates, and a CSS transition becomes a time-based tween
that restarts from the current value when it is retargeted (so a drag that
enters and leaves quickly never jumps).
"""

from __future__ import annotations

import os
import sys
import time
from typing import Callable, Dict, Optional

FRAME_MS = 15  # ~60 fps; tweens are time-based, so a slow frame never slows the motion


def cubic_bezier(x1: float, y1: float, x2: float, y2: float) -> Callable[[float], float]:
    """CSS cubic-bezier(x1, y1, x2, y2) as a function of progress 0..1."""

    def sample(a1, a2, t):
        return ((1 - 3 * a2 + 3 * a1) * t + (3 * a2 - 6 * a1)) * t * t + 3 * a1 * t

    def slope(a1, a2, t):
        return 3 * (1 - 3 * a2 + 3 * a1) * t * t + 2 * (3 * a2 - 6 * a1) * t + 3 * a1

    def ease(x: float) -> float:
        if x <= 0:
            return 0.0
        if x >= 1:
            return 1.0
        t = x
        for _ in range(8):  # Newton's method, then bisection if it stalls
            error = sample(x1, x2, t) - x
            if abs(error) < 1e-5:
                return sample(y1, y2, t)
            d = slope(x1, x2, t)
            if abs(d) < 1e-6:
                break
            t -= error / d
        lo, hi, t = 0.0, 1.0, x
        for _ in range(30):
            value = sample(x1, x2, t)
            if abs(value - x) < 1e-5:
                break
            lo, hi = (t, hi) if value < x else (lo, t)
            t = (lo + hi) / 2
        return sample(y1, y2, t)

    return ease


EASE_OUT = cubic_bezier(0.23, 1, 0.32, 1)  # entering / exiting, press feedback
EASE_IN_OUT = cubic_bezier(0.77, 0, 0.175, 1)  # something moving or resizing on screen
EASE = cubic_bezier(0.25, 0.1, 0.25, 1)  # colour changes


def blend(a: str, b: str, t: float) -> str:
    """Mix two #RRGGBB colours: t = 0 gives a, 1 gives b. Stands in for opacity."""
    t = max(0.0, min(1.0, t))
    ca = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    cb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def reduced_motion() -> bool:
    """Windows' "Animation effects" switch (Settings › Accessibility › Visual effects).

    When it is off, things stop moving and resizing; colour fades stay, shorter."""
    if os.environ.get("MARKITDOWN_REDUCE_MOTION") == "1":
        return True
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        enabled = ctypes.c_int(1)
        SPI_GETCLIENTAREAANIMATION = 0x1042
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETCLIENTAREAANIMATION, 0, ctypes.byref(enabled), 0)
        return not enabled.value
    except Exception:
        return False


class Motion:
    """Named values on one widget, each animated toward a target like a CSS transition."""

    def __init__(self, widget):
        self.widget = widget
        self.values: Dict[str, float] = {}
        self.jobs: Dict[str, str] = {}

    @property
    def active(self) -> bool:
        return bool(self.jobs)

    def get(self, key: str, default: float = 1.0) -> float:
        return self.values.get(key, default)

    def set(self, key: str, value: float) -> None:
        self._cancel(key)
        self.values[key] = value

    def to(self, key: str, target: float, duration_ms: int, ease: Callable[[float], float],
           on_frame: Callable[[], None], delay_ms: int = 0, start: Optional[float] = None,
           on_done: Optional[Callable[[], None]] = None) -> None:
        """Animate ``key`` from its current value (or ``start``) to ``target``."""
        self._cancel(key)
        begin = self.values.get(key, target) if start is None else start
        self.values[key] = begin
        if duration_ms <= 0 or begin == target:
            self.values[key] = target
            on_frame()
            if on_done:
                on_done()
            return
        t0 = time.perf_counter() + delay_ms / 1000

        def step():
            try:
                progress = (time.perf_counter() - t0) / (duration_ms / 1000)
                if progress >= 1:  # last frame: no longer running when it is drawn
                    self.jobs.pop(key, None)
                    self.values[key] = target
                    on_frame()
                    if on_done:
                        on_done()
                    return
                if progress >= 0:
                    self.values[key] = begin + (target - begin) * ease(progress)
                    on_frame()
                self.jobs[key] = self.widget.after(FRAME_MS, step)
            except Exception:  # the widget went away mid-animation
                self.jobs.pop(key, None)

        self.jobs[key] = self.widget.after(max(0, delay_ms) if delay_ms else 0, step)

    def _cancel(self, key: str) -> None:
        job = self.jobs.pop(key, None)
        if job is not None:
            try:
                self.widget.after_cancel(job)
            except Exception:
                pass
