#!/usr/bin/env python3
"""Lenta tugmalari uchun ikonkalar (32x32 PNG) — 4-Ornatuvchi-EXE/src/icons/ ga yoziladi."""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "4-Ornatuvchi-EXE", "src", "icons")
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
S = 256                      # chizish o'lchami, keyin 32 ga kichraytiriladi
TEAL, BLUE, PURPLE, AMBER, SLATE, GREEN = (30, 110, 90), (37, 99, 168), (123, 63, 160), (183, 112, 28), (85, 96, 110), (46, 125, 50)
WHITE = (255, 255, 255, 255)


def base(color):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([8, 8, S - 9, S - 9], radius=48, fill=color + (255,))
    return im, d


def text(d, s, size, cx, cy, fill=WHITE):
    f = ImageFont.truetype(BOLD, size)
    bb = d.textbbox((0, 0), s, font=f)
    d.text((cx - (bb[0] + bb[2]) / 2, cy - (bb[1] + bb[3]) / 2), s, font=f, fill=fill)


def arrow(d, x1, x2, y, w=14, fill=WHITE):
    d.line([x1, y, x2 - 18, y], fill=fill, width=w)
    d.polygon([(x2, y), (x2 - 34, y - 24), (x2 - 34, y + 24)], fill=fill)


def frame(d):                 # belgilangan matn: uzuq ramka
    for x in range(30, S - 40, 34):
        d.line([x, 28, x + 18, 28], fill=WHITE, width=10)
        d.line([x, S - 29, x + 18, S - 29], fill=WHITE, width=10)
    for y in range(30, S - 40, 34):
        d.line([28, y, 28, y + 18], fill=WHITE, width=10)
        d.line([S - 29, y, S - 29, y + 18], fill=WHITE, width=10)


def pair(color, a, b, sel=False):
    im, d = base(color)
    if sel: frame(d)
    text(d, a, 104, 70, 108)
    text(d, b, 104, 186, 108)
    arrow(d, 60, 196, 196, w=16)
    return im


def shield(color, inner, plus=False):
    im, d = base(color)
    d.polygon([(128, 34), (210, 66), (200, 150), (128, 222), (56, 150), (46, 66)], fill=WHITE)
    text(d, inner, 92, 128, 120, fill=color + (255,))
    if plus:
        d.ellipse([162, 158, 238, 234], fill=GREEN + (255,), outline=WHITE, width=8)
        text(d, "+", 66, 200, 194)
    return im


def make():
    os.makedirs(OUT, exist_ok=True)
    icons = {}
    # yangi alifbo
    im, d = base(TEAL); text(d, "Ö", 170, 128, 122); icons["iYangi"] = im
    im, d = base(TEAL); frame(d); text(d, "Ö", 140, 128, 124); icons["iTanlangan"] = im
    # eski lotinga
    im, d = base(AMBER); text(d, "Oʻ", 132, 128, 110); arrow(d, 196, 60, 206, w=16)
    d.polygon([(60, 206), (94, 182), (94, 230)], fill=WHITE); icons["iEski"] = im
    im, d = base(AMBER); frame(d); text(d, "Oʻ", 118, 128, 124); icons["iEskiSel"] = im
    # yozishda (klaviatura)
    im, d = base(TEAL)
    d.rounded_rectangle([30, 120, 226, 214], radius=18, outline=WHITE, width=10)
    for row, y in enumerate([138, 166]):
        for x in range(46 + row * 12, 206, 30):
            d.rectangle([x, y, x + 18, y + 16], fill=WHITE)
    d.rectangle([80, 192, 176, 202], fill=WHITE)
    text(d, "ş", 104, 128, 70); icons["iYozish"] = im
    # kirill <-> lotin
    icons["iKirill"] = pair(BLUE, "Ш", "Ş")
    icons["iKirillSel"] = pair(BLUE, "Ш", "Ş", sel=True)
    icons["iLotinKirill"] = pair(PURPLE, "Ş", "Ш")
    icons["iLotinKirillSel"] = pair(PURPLE, "Ş", "Ш", sel=True)
    # istisnolar
    icons["iIstisnolar"] = shield(SLATE, "A")
    icons["iIstisnoQosh"] = shield(SLATE, "A", plus=True)
    icons["iKirillIstisno"] = shield(BLUE, "Ж")
    icons["iKirillIstisnoQosh"] = shield(BLUE, "Ж", plus=True)
    im, d = base(SLATE)
    for i, y in enumerate([70, 118, 166]):
        d.ellipse([48, y - 12, 72, y + 12], fill=WHITE)
        d.rounded_rectangle([88, y - 9, 208 - (i % 2) * 30, y + 9], radius=8, fill=WHITE)
    icons["iRoyxat"] = im
    im, d = base(GREEN); d.ellipse([40, 40, 216, 216], outline=WHITE, width=16); text(d, "?", 120, 128, 126)
    icons["iYordam"] = im
    # tezkor tugmalar: ikki klavish "Alt" + "K"
    im, d = base(SLATE)
    d.rounded_rectangle([22, 70, 140, 186], radius=18, outline=WHITE, width=10); text(d, "Alt", 50, 81, 128)
    d.rounded_rectangle([156, 70, 234, 186], radius=18, outline=WHITE, width=10); text(d, "K", 64, 195, 128)
    icons["iTugmalar"] = im
    for name, im in icons.items():
        im.resize((32, 32), Image.LANCZOS).save(os.path.join(OUT, name + ".png"))
    return icons


if __name__ == "__main__":
    icons = make()
    # ko'rib chiqish uchun varaq
    names = list(icons)
    sheet = Image.new("RGBA", (len(names) * 72 + 16, 140), (245, 245, 245, 255))
    d = ImageDraw.Draw(sheet)
    for i, n in enumerate(names):
        big = icons[n].resize((48, 48), Image.LANCZOS)
        small = Image.open(os.path.join(OUT, n + ".png")).resize((16, 16), Image.LANCZOS)
        sheet.alpha_composite(big, (16 + i * 72, 12))
        sheet.alpha_composite(Image.open(os.path.join(OUT, n + ".png")), (24 + i * 72, 70))
        sheet.alpha_composite(small, (32 + i * 72, 112))
    sheet.save(os.path.join(OUT, "..", "icons_preview.png"))
    print(len(icons), "ta ikonka")
