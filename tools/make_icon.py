# -*- coding: utf-8 -*-
"""FSN.ico 생성 (청록→파랑 그라데이션 + 흰 봉투 + 체크 배지).  python tools/make_icon.py"""
from pathlib import Path

from PIL import Image, ImageDraw

S = 1024  # 크게 그린 뒤 축소 (안티앨리어싱)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def make():
    top, bottom = (22, 160, 160), (30, 88, 190)
    grad = Image.new("RGB", (S, S))
    gd = ImageDraw.Draw(grad)
    for y in range(S):
        gd.line([(0, y), (S, y)], fill=lerp(top, bottom, y / S))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([24, 24, S - 24, S - 24], radius=200, fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(img)
    # 봉투
    x0, y0, x1, y1 = 170, 290, 854, 760
    w = 44
    d.rounded_rectangle([x0, y0, x1, y1], radius=60, fill=(255, 255, 255, 255))
    flap = (30, 110, 175, 255)
    d.line([(x0 + 40, y0 + 40), ((x0 + x1) // 2, y0 + 300), (x1 - 40, y0 + 40)], fill=flap, width=w, joint="curve")
    # 체크 배지
    cx, cy, r = 790, 740, 175
    d.ellipse([cx - r - 22, cy - r - 22, cx + r + 22, cy + r + 22], fill=(255, 255, 255, 255))
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(12, 163, 12, 255))
    d.line([(cx - 85, cy + 5), (cx - 20, cy + 70), (cx + 95, cy - 65)], fill=(255, 255, 255, 255), width=52,
           joint="curve")
    return img


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "FSN.ico"
    big = make()
    sizes = [16, 24, 32, 48, 64, 128, 256]
    big.resize((256, 256), Image.LANCZOS).save(out, sizes=[(s, s) for s in sizes])
    big.resize((256, 256), Image.LANCZOS).save(out.with_suffix(".png"))
    print("생성:", out)
