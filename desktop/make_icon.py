r"""Generate desktop/app_icon.ico — a pair of glasses with an amber pencil badge.

Deliberately the same drawing as the Glasses Validator Desktop icon
(bgal-png/Glasses-Validator-Desktop/make_icon.py) so the two apps read as one
family: identical slate-blue glasses, identical badge position and size. Only
the badge differs — green disc + white check there, amber disc + white pencil
here (validate vs fill).

Drawn at high resolution and downsampled so the small sizes stay smooth.
Re-run only when the icon should change:
    "C:\gv\Scripts\python.exe" desktop\make_icon.py
"""
import math
import os

from PIL import Image, ImageDraw

S = 1024                       # working canvas
FRAME = (43, 90, 158, 255)     # slate blue   — same as the validator
PENCIL = (232, 148, 26, 255)   # amber        — "fill/edit" instead of green
LENS = (120, 170, 225, 70)     # faint glass tint
WHITE = (255, 255, 255, 255)

HERE = os.path.dirname(os.path.abspath(__file__))


def rounded(draw, box, r, **kw):
    draw.rounded_rectangle(box, radius=r, **kw)


def build(size=S):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    u = size / 1024.0                       # scale helper
    w = int(58 * u)                         # stroke width

    # ---- glasses (unchanged from the validator) ----
    top, h = int(340 * u), int(310 * u)
    lens_w = int(350 * u)
    inset = int(96 * u)                     # room for the temple stubs
    left = (inset, top, inset + lens_w, top + h)
    right = (size - inset - lens_w, top, size - inset, top + h)
    for box in (left, right):
        rounded(d, box, int(85 * u), fill=LENS, outline=FRAME, width=w)

    # bridge
    d.arc((left[2] - int(20 * u), top + int(30 * u),
           right[0] + int(20 * u), top + int(180 * u)),
          start=200, end=340, fill=FRAME, width=w)

    # temple stubs — short, near-horizontal, hugging the frame
    ty = top + int(70 * u)
    d.line((left[0], ty, int(10 * u), ty - int(34 * u)), fill=FRAME, width=w)
    d.line((right[2], ty, size - int(10 * u), ty - int(34 * u)), fill=FRAME, width=w)

    # ---- pencil badge (same disc as the validator's check) ----
    cx, cy, r = int(720 * u), int(730 * u), int(250 * u)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PENCIL)

    # Pencil drawn along a diagonal: tip lower-left (as if writing), butt upper-right.
    ax, ay = cx - 138 * u, cy + 138 * u     # tip apex
    bx, by = cx + 142 * u, cy - 142 * u     # butt end
    dx, dy = bx - ax, by - ay
    length = math.hypot(dx, dy)
    ux, uy = dx / length, dy / length       # along the pencil
    px, py = -uy, ux                        # across it
    hw = 50 * u                             # half-width of the body
    tip_len = 96 * u

    # shoulder: where the tapered tip meets the body
    sx, sy = ax + ux * tip_len, ay + uy * tip_len
    s1 = (sx + px * hw, sy + py * hw)
    s2 = (sx - px * hw, sy - py * hw)
    b1 = (bx + px * hw, by + py * hw)
    b2 = (bx - px * hw, by - py * hw)

    d.polygon([s1, b1, b2, s2], fill=WHITE)        # body
    d.polygon([(ax, ay), s1, s2], fill=WHITE)      # tapered tip

    # One amber band just above the tip, so it reads as a pencil and not a
    # blob. Everything finer than this disappears at 16 px anyway.
    band_at = tip_len + 34 * u
    band_th = 24 * u
    gx, gy = ax + ux * band_at, ay + uy * band_at
    g1 = (gx + px * hw, gy + py * hw)
    g2 = (gx - px * hw, gy - py * hw)
    hx, hy = ax + ux * (band_at + band_th), ay + uy * (band_at + band_th)
    g3 = (hx - px * hw, hy - py * hw)
    g4 = (hx + px * hw, hy + py * hw)
    d.polygon([g1, g2, g3, g4], fill=PENCIL)

    return img


def main():
    base = build()
    sizes = [16, 24, 32, 48, 64, 128, 256]
    frames = [base.resize((s, s), Image.LANCZOS) for s in sizes]
    ico = os.path.join(HERE, "app_icon.ico")
    png = os.path.join(HERE, "app_icon.png")
    frames[-1].save(ico, format="ICO", sizes=[(s, s) for s in sizes])
    base.resize((256, 256), Image.LANCZOS).save(png)
    print(f"wrote {ico} and {png}")


if __name__ == "__main__":
    main()
