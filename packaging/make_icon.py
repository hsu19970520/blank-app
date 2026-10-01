#!/usr/bin/env python3
"""Draw the app icon: the Markdown mark (CC0) on an accent-blue rounded square.

    python packaging/make_icon.py   # writes app/assets/icon.png and app/assets/icon.ico

Requires Pillow (installed with markitdown[all] via python-pptx).
"""

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "app" / "assets"

SIZE = 1024
TOP, BOTTOM = (0x2B, 0x7F, 0xD4), (0x05, 0x60, 0xB6)  # accent, light to dark

# The Markdown mark (https://github.com/dcurtis/markdown-mark, CC0), 208x128 viewBox.
M_SHAPE = [(30, 98), (30, 30), (50, 30), (70, 55), (90, 30), (110, 30), (110, 98),
           (90, 98), (90, 59), (70, 84), (50, 59), (50, 98)]
ARROW = [(155, 98), (125, 65), (145, 65), (145, 30), (165, 30), (165, 65), (185, 65)]


def draw() -> Image.Image:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

    gradient = Image.new("RGBA", (SIZE, SIZE))
    gd = ImageDraw.Draw(gradient)
    for y in range(SIZE):
        t = y / (SIZE - 1)
        gd.line([(0, y), (SIZE, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(TOP, BOTTOM)) + (255,))

    margin = round(SIZE * 0.06)
    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [margin, margin, SIZE - margin, SIZE - margin], radius=round(SIZE * 0.2), fill=255
    )
    img.paste(gradient, (0, 0), mask)

    # Fit the mark (x 30..185, y 30..98) to ~64% of the icon width, centred.
    scale = SIZE * 0.64 / 155
    ox = (SIZE - 155 * scale) / 2 - 30 * scale
    oy = (SIZE - 68 * scale) / 2 - 30 * scale
    d = ImageDraw.Draw(img)
    for shape in (M_SHAPE, ARROW):
        d.polygon([(ox + x * scale, oy + y * scale) for x, y in shape], fill=(255, 255, 255, 255))
    return img


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    icon = draw()
    icon.resize((256, 256), Image.LANCZOS).save(ASSETS / "icon.png")
    icon.save(ASSETS / "icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"wrote {ASSETS / 'icon.png'} and {ASSETS / 'icon.ico'}")


if __name__ == "__main__":
    main()
