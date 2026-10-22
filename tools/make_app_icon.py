#!/usr/bin/env python3
"""Özbekça.exe ikonkasi (icon.ico: 16..256 px) va oldindan ko'rish (app_icon_preview.png)."""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "4-Ornatuvchi-EXE", "src")
SERIF = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"   # kichik o'lchamlarda yo'g'onroq
TEAL_D, TEAL_L, AMBER = (18, 84, 68), (40, 135, 110), (240, 176, 60)
S = 1024


def draw(small=False):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    # vertikal gradient fon, yumaloq burchak
    grad = Image.new("RGBA", (S, S))
    gd = ImageDraw.Draw(grad)
    for y in range(S):
        t = y / (S - 1)
        c = tuple(int(TEAL_L[i] * (1 - t) + TEAL_D[i] * t) for i in range(3)) + (255,)
        gd.line([(0, y), (S, y)], fill=c)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([24, 24, S - 25, S - 25], radius=210, fill=255)
    im.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(im)
    # Ö va Ç bir xil kattalikda, bir-biriga kirishib turadi (har qanday o'lchamda ham)
    f = ImageFont.truetype(SANS, 640) if small else ImageFont.truetype(SERIF, 560)
    bo, bc = d.textbbox((0, 0), "Ö", font=f), d.textbbox((0, 0), "Ç", font=f)
    wo, wc = bo[2] - bo[0], bc[2] - bc[0]
    overlap = int(wo * (0.36 if small else 0.30))
    total = wo + wc - overlap
    x0 = (S - total) / 2
    top = min(bo[1], bc[1]); bottom = max(bo[3], bc[3])
    y = (S - (bottom - top)) / 2 - top
    xo, xc = x0 - bo[0], x0 + wo - overlap - bc[0]
    sh = 16 if not small else 0
    d.text((xc + sh, y + sh), "Ç", font=f, fill=(0, 0, 0, 80))
    d.text((xc, y), "Ç", font=f, fill=(AMBER if not small else (255, 196, 70)) + (255,))
    d.text((xo + sh, y + sh), "Ö", font=f, fill=(0, 0, 0, 80))       # Ö oldinda, soyasi Ç ustiga tushadi
    d.text((xo, y), "Ö", font=f, fill=(255, 255, 255, 255))
    return im


big, small = draw(False), draw(True)
imgs = [(small if n <= 32 else big).resize((n, n), Image.LANCZOS) for n in (16, 24, 32, 48, 64, 128, 256)]
imgs[-1].save(os.path.join(SRC, "icon.ico"), sizes=[(i.width, i.height) for i in imgs], append_images=imgs[:-1])
prev = Image.new("RGBA", (16 + 24 + 32 + 48 + 64 + 128 + 256 + 8 * 20, 276), (255, 255, 255, 255))
x = 10
for i in imgs:
    prev.paste(i, (x, 266 - i.height), i); x += i.width + 20
prev.save(os.path.join(SRC, "app_icon_preview.png"))
print("icon.ico tayyor")
