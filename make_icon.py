"""
make_icon.py — generates assets/icon.ico for the app / .exe build.

Run once:  python make_icon.py
"""

import os

from PIL import Image, ImageDraw, ImageFont

SIZE = 256
INDIGO_TOP = (99, 102, 241)     # #6366F1
INDIGO_BOT = (67, 56, 202)      # #4338CA
PAGE = (255, 255, 255)
FOLD = (224, 226, 240)
TEXTLINE = (203, 208, 220)
PDF_RED = (225, 76, 76)


def _font(size, bold=True):
    candidates = [
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if os.path.isfile(p):
            try:
                return ImageFont.truetype(p, size)
            except OSError:
                pass
    return ImageFont.load_default()


def build(path="assets/icon.ico"):
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # rounded-square vertical gradient background
    grad = Image.new("RGBA", (SIZE, SIZE))
    gd = ImageDraw.Draw(grad)
    for y in range(SIZE):
        t = y / (SIZE - 1)
        c = tuple(int(a + (b - a) * t) for a, b in zip(INDIGO_TOP, INDIGO_BOT))
        gd.line([(0, y), (SIZE, y)], fill=c + (255,))
    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle([8, 8, SIZE - 8, SIZE - 8],
                                           radius=52, fill=255)
    img.paste(grad, (0, 0), mask)

    # white page with folded corner
    px0, py0, px1, py1 = 64, 40, 192, 216
    fold = 34
    d.rounded_rectangle([px0, py0, px1, py1], radius=14, fill=PAGE)
    d.polygon([(px1 - fold, py0), (px1, py0 + fold), (px1 - fold, py0 + fold)],
              fill=FOLD)
    d.rectangle([px1 - fold, py0, px1, py0 + 2], fill=INDIGO_TOP)

    # "M" mark + text lines
    f_m = _font(64)
    d.text(((px0 + px1) / 2, 78), "M", font=f_m, fill=INDIGO_BOT, anchor="mm")
    for i, w in enumerate((84, 84, 60)):
        y = 118 + i * 16
        d.rounded_rectangle([px0 + 22, y, px0 + 22 + w, y + 7], radius=3,
                            fill=TEXTLINE)

    # red PDF pill
    pw, ph = 92, 34
    cx = (px0 + px1) / 2
    d.rounded_rectangle([cx - pw / 2, py1 - 52, cx + pw / 2, py1 - 52 + ph],
                        radius=10, fill=PDF_RED)
    d.text((cx, py1 - 52 + ph / 2), "PDF", font=_font(24),
           fill=(255, 255, 255), anchor="mm")

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    img.save(path, format="ICO",
             sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    img.save(os.path.splitext(path)[0] + ".png")
    print("wrote", path)


if __name__ == "__main__":
    build()
