"""Logo (500x500) dan banner besar (3360x840) toko Etsy AbhesonicTemplates -> out/etsy-brand/"""
import os
from PIL import Image, ImageDraw, ImageFont, ImageFilter
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "out", "etsy-brand"); os.makedirs(OUT, exist_ok=True)
F = "/usr/share/fonts/opentype/inter/Inter-%s.otf"
TEAL, TEAL_D, AMBER, WHITE = "#0F766E", "#115E59", "#F59E0B", "#FFFFFF"


def f(w, s): return ImageFont.truetype(F % w, s)


def grid_icon(d, x, y, size, line=WHITE, accent=AMBER):
    """Ikon spreadsheet: kotak 4x4 dengan baris judul dan satu sel kuning, plus garis grafik naik."""
    cell = size / 4; r = int(size * 0.08)
    d.rounded_rectangle((x, y, x + size, y + size), radius=r, outline=line, width=max(4, int(size * 0.035)))
    d.rectangle((x + 2, y + 2, x + size - 2, y + cell), fill=line)
    for i in range(1, 4):
        d.line((x + cell * i, y + cell, x + cell * i, y + size), fill=line, width=max(3, int(size * 0.02)))
        d.line((x, y + cell * i, x + size, y + cell * i), fill=line, width=max(3, int(size * 0.02)))
    d.rectangle((x + cell * 2 + 4, y + cell * 2 + 4, x + cell * 3 - 4, y + cell * 3 - 4), fill=accent)


def logo():
    im = Image.new("RGB", (500, 500), TEAL); d = ImageDraw.Draw(im)
    grid_icon(d, 110, 70, 280)
    pts = [(120, 330), (200, 270), (270, 300), (380, 190)]
    d.line(pts, fill=AMBER, width=16, joint="curve")
    for p in pts: d.ellipse((p[0] - 12, p[1] - 12, p[0] + 12, p[1] + 12), fill=AMBER)
    t = "ABHESONIC"; fo = f("ExtraBold", 54); w = d.textlength(t, font=fo); d.text(((500 - w) / 2, 380), t, font=fo, fill=WHITE)
    t = "TEMPLATES"; fo = f("Medium", 30); w = d.textlength(t, font=fo); d.text(((500 - w) / 2, 442), t, font=fo, fill="#CCFBF1")
    im.save(os.path.join(OUT, "logo.png"))


def banner(covers):
    W, H = 3360, 840
    im = Image.new("RGB", (W, H), TEAL); d = ImageDraw.Draw(im)
    for i in range(0, W, 120): d.line((i, 0, i, H), fill=TEAL_D, width=2)
    for j in range(0, H, 120): d.line((0, j, W, j), fill=TEAL_D, width=2)

    def card(path, cx, cy, w, ang):
        c = Image.open(path).convert("RGB"); h = int(c.size[1] * w / c.size[0]); c = c.resize((w, h), Image.LANCZOS)
        fr = Image.new("RGBA", (w + 24, h + 24), (255, 255, 255, 255)); fr.paste(c, (12, 12))
        sh = Image.new("RGBA", (w + 64, h + 64), (0, 0, 0, 0)); ImageDraw.Draw(sh).rectangle((20, 28, w + 44, h + 52), fill=(0, 0, 0, 110))
        sh = sh.filter(ImageFilter.GaussianBlur(14)); sh.paste(fr, (20, 20), fr); sh = sh.rotate(ang, expand=True, resample=Image.BICUBIC)
        im.paste(sh, (int(cx - sh.size[0] / 2), int(cy - sh.size[1] / 2)), sh)
    card(covers[0], 300, 420, 470, 6); card(covers[1], 640, 460, 460, -4)
    card(covers[2], W - 640, 460, 460, 4); card(covers[3], W - 300, 420, 470, -6)
    cx = W / 2
    t = "Abhesonic Templates"; fo = f("ExtraBold", 132); w = d.textlength(t, font=fo); d.text((cx - w / 2, 190), t, font=fo, fill=WHITE)
    t = "Smart spreadsheets for Excel & Google Sheets"; fo = f("Medium", 62); w = d.textlength(t, font=fo); d.text((cx - w / 2, 375), t, font=fo, fill="#CCFBF1")
    pills = ["Budget & Finance", "Small Business", "Planners"]; fo = f("SemiBold", 46)
    ws = [d.textlength(p, font=fo) + 80 for p in pills]; x = cx - (sum(ws) + 30 * (len(pills) - 1)) / 2
    for p, pw in zip(pills, ws):
        d.rounded_rectangle((x, 500, x + pw, 590), radius=45, fill=AMBER); d.text((x + 40, 518), p, font=fo, fill="#1F2937"); x += pw + 30
    t = "Instant download  ·  Blank + sample copy  ·  No macros"; fo = f("Regular", 44); w = d.textlength(t, font=fo); d.text((cx - w / 2, 645), t, font=fo, fill=WHITE)
    im.save(os.path.join(OUT, "banner.jpg"), quality=90)


if __name__ == "__main__":
    o = os.path.join(HERE, "out")
    covers = [os.path.join(o, n, "cover.png") for n in ["Budget-Expense-Planner", "Small-Business-Bookkeeping", "Wedding-Planner-Budget", "Habit-Tracker-Heatmap"]]
    logo(); banner(covers); print("ok", OUT)
