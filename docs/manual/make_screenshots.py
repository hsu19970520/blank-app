"""Screenshots for the user guide, made from synthetic sample workbooks (no real company data).

    xvfb-run python docs/manual/make_screenshots.py docs/manual/shots <scratch-dir-for-samples>
"""
import random
import sys
import time
import tkinter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path[:0] = [str(ROOT / "app"), str(ROOT / "skills" / "markitdown" / "scripts")]
import gui
out, sample = Path(sys.argv[1]), Path(sys.argv[2])
gui.system_appearance = lambda: "light"
class FakeDnd: DND_FILES = "DND_Files"; COPY = "copy"
tkinter.Misc.drag_source_register = lambda self, *a: None
tkinter.Misc.drop_target_register = lambda self, *a: None
tkinter.Misc.dnd_bind = lambda self, *a, **k: None

# Synthetic sample workbooks (no real company data): a forecast waterfall and a small report.
from openpyxl import Workbook
rng = random.Random(3)
wb = Workbook(); ws = wb.active; ws.title = "By MPN x Site"
weeks = [f"W{i:02d}" for i in range(1, 27)]
ws.append(["MPN", "Site", "Brand", "Version", *weeks, "TTL"])
versions = [f"2026/{m:02d}/{d:02d}" for m, d in [(7, 6), (7, 13), (7, 20), (7, 27), (8, 3), (8, 10), (8, 17), (8, 24), (8, 31), (9, 7), (9, 14), (9, 21), (9, 28), (10, 5)]]
parts = [(f"PN-{i:04d}", rng.choice(["TW", "VN"]), rng.choice(["Brand A", "Brand B", "Brand C"])) for i in range(160)]
for v in versions:
    for mpn, site, brand in parts:
        q = [rng.choice([0, 0, rng.randint(100, 90000)]) for _ in weeks]
        ws.append([mpn, site, brand, v, *q, sum(q)])
wb.save(sample / "範例 FCST Waterfall.xlsx")
wb = Workbook(); ws = wb.active; ws.title = "週報"
ws.append(["料號", "客戶", "本週需求", "上週需求", "差異"])
for i in range(40):
    a, b = rng.randint(0, 5000), rng.randint(0, 5000); ws.append([f"PN-{i:04d}", rng.choice(["客戶甲", "客戶乙"]), a, b, a - b])
wb.save(sample / "範例週報.xlsx")

root, app = gui.create_window()
app.dnd = FakeDnd
def pump(t):
    end = time.time() + t
    while time.time() < end: root.update(); time.sleep(0.01)
def grab(win, name):
    pump(0.5)
    from PIL import ImageGrab
    x, y = win.winfo_rootx(), win.winfo_rooty()
    ImageGrab.grab(bbox=(x, y, x + win.winfo_width(), y + win.winfo_height())).save(out / f"{name}.png")
root.geometry("1100x700+20+20"); pump(1)
while app.converter is None: pump(0.2)
grab(root, "main_empty")
root.withdraw()
p = app.show_panel(standalone=True)
p.win.geometry("400x640+40+40"); pump(0.8)
grab(p.win, "panel_empty")
p.add([str(sample / "範例週報.xlsx"), str(sample / "範例 FCST Waterfall.xlsx")])
while p.busy or p.animating: pump(0.1)
grab(p.win, "panel_results")
big = next(c for c in p.cards if c.result.offers)
print("offers", big.result.offers, "total", big.result.total_tokens, "files", big.result.data_files)
p.convert_latest(big.result.source, 1)
while p.busy or p.animating: pump(0.1)
grab(p.win, "panel_latest")
print([(c.result.output.name, c.result.tokens) for c in p.cards])
root.destroy()
