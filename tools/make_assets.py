"""Generate assets/icon.ico and assets/splash.png using PyMuPDF only."""
import struct
from pathlib import Path

import pymupdf

ASSETS = Path(__file__).resolve().parent.parent / "assets"
BLUE = (0, 0.443, 0.89)  # #0071e3
INK = (0.114, 0.114, 0.122)  # #1d1d1f
GREY = (0.525, 0.525, 0.545)  # #86868b


def draw_document(page, x, y, w, h, stroke):
    """Rounded document outline with a few text lines."""
    shape = page.new_shape()
    shape.draw_rect(pymupdf.Rect(x, y, x + w, y + h), radius=0.12)
    shape.finish(color=BLUE, fill=(1, 1, 1), width=stroke)
    pad = w * 0.22
    for i, frac in enumerate((0.3, 0.45, 0.6, 0.75)):
        ly = y + h * frac
        end = x + w - pad if i < 3 else x + w * 0.55
        shape.draw_line((x + pad, ly), (end, ly))
    shape.finish(color=BLUE, width=stroke, closePath=False)
    shape.commit()


def render(page, path):
    page.get_pixmap(dpi=72).save(str(path))


def make_icon():
    doc = pymupdf.open()
    page = doc.new_page(width=256, height=256)
    draw_document(page, 48, 24, 160, 208, stroke=12)
    png_path = ASSETS / "_icon.png"
    render(page, png_path)
    png = png_path.read_bytes()
    png_path.unlink()
    header = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", 0, 0, 0, 0, 1, 32, len(png), 22)
    (ASSETS / "icon.ico").write_bytes(header + entry + png)


def make_splash():
    doc = pymupdf.open()
    page = doc.new_page(width=420, height=240)
    draw_document(page, 40, 60, 90, 120, stroke=5)
    page.insert_text((160, 125), "PDF Converter", fontsize=28, fontname="helv", color=INK)
    page.insert_text((160, 155), "Starting...", fontsize=14, fontname="helv", color=GREY)
    render(page, ASSETS / "splash.png")


if __name__ == "__main__":
    ASSETS.mkdir(exist_ok=True)
    make_icon()
    make_splash()
