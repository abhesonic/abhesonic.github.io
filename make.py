"""Build one template: python3 make.py <module_name>  ->  out/<file_base>/ (xlsx x2, zip, cover.png, cover.jpg, listing.md)"""
import sys, os, glob, json, subprocess, shutil, importlib, zipfile, tempfile
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, "templates"))

def recalc(path):
    cands = glob.glob("/root/.claude/skills/**/xlsx/scripts/recalc.py", recursive=True)
    if not cands: raise SystemExit("recalc.py from the xlsx skill not found")
    out = subprocess.run([sys.executable, cands[0], path, "180"], capture_output=True, text=True).stdout
    res = json.loads(out[out.index("{"):])
    if res.get("status") != "success" or res.get("total_errors"): raise SystemExit(f"Formula errors in {path}: {res}")
    return res["total_formulas"]

FB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; FR = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

def _fit(text, path, size, maxw):
    while ImageFont.truetype(path, size).getlength(text) > maxw and size > 10: size -= 1
    return ImageFont.truetype(path, size)

def _compose(crop, meta, W, H, title_size, tag_size, top):
    """Coloured background, title + tagline, then the dashboard screenshot on a white card."""
    bg = Image.new("RGB", (W, H), meta["color"]); d = ImageDraw.Draw(bg); m = int(W * 0.044)
    tf = _fit(meta["title"], FB, title_size, W - 2 * m); d.text((m, int(top * 0.27)), meta["title"], font=tf, fill="white")
    gf = _fit(meta["tagline"], FR, tag_size, W - 2 * m); d.text((m + 2, int(top * 0.67)), meta["tagline"], font=gf, fill="#F1F5F9")
    cw = W - 2 * m + 20; ch = int(crop.size[1] * cw / crop.size[0]); img = crop.resize((cw, ch), Image.LANCZOS)
    ch = min(ch, H - top - 10); card = Image.new("RGB", (cw + 16, ch + 16), "white"); card.paste(img.crop((0, 0, cw, ch)), (8, 8))
    bg.paste(card, (m - 18, top)); return bg

def cover(sample_xlsx, meta, dest):
    """cover.png/jpg 1600x1200 (4:3), gumroad-cover.png 1280x720 (16:9) and thumb.png 600x600 (square)."""
    tmp = tempfile.mkdtemp()
    shutil.copy(sample_xlsx, os.path.join(tmp, "s.xlsx"))
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, os.path.join(tmp, "s.xlsx")], capture_output=True, timeout=240)
    subprocess.run(["pdftoppm", "-r", "170", "-png", "-f", "2", "-l", "2", os.path.join(tmp, "s.pdf"), os.path.join(tmp, "p")], check=True)
    im = Image.open(glob.glob(os.path.join(tmp, "p*.png"))[0]).convert("RGB"); w, h = im.size
    crop = im.crop((int(w*0.06), int(h*0.10), int(w*0.93), int(h*0.80)))
    big = _compose(crop, meta, 1600, 1200, 66, 31, 230)
    big.save(os.path.join(dest, "cover.png")); big.save(os.path.join(dest, "cover.jpg"), quality=85)
    wide = _compose(crop, meta, 1280, 720, 52, 25, 180)
    wide.save(os.path.join(dest, "gumroad-cover.png"))
    sq = Image.new("RGB", (600, 600), meta["color"]); d = ImageDraw.Draw(sq)
    words = meta["title"].split(); lines = []; cur = ""
    for wd in words:
        t = (cur + " " + wd).strip()
        if ImageFont.truetype(FB, 54).getlength(t) <= 520: cur = t
        else: lines.append(cur); cur = wd
    lines.append(cur)
    size = 54 if len(lines) <= 3 else 44
    y = 48
    for ln in lines[:4]:
        f = _fit(ln, FB, size, 520); d.text((40, y), ln, font=f, fill="white"); y += int(size * 1.18)
    d.text((42, y + 6), "Excel & Google Sheets", font=ImageFont.truetype(FR, 26), fill="#F1F5F9")
    top = y + 60; cw = 560; ch = int(crop.size[1] * cw / crop.size[0]); img = crop.resize((cw, ch), Image.LANCZOS)
    ch = min(ch, 600 - top - 20); card = Image.new("RGB", (cw + 10, ch + 10), "white"); card.paste(img.crop((0, 0, cw, ch)), (5, 5))
    sq.paste(card, (15, top)); sq.save(os.path.join(dest, "thumb.png"))

def listing(meta):
    fb = meta["file_base"]
    note = (f'Thanks for your purchase! Download {fb}.zip, open {fb}.xlsx and start on the "Start Here" tab. '
            'For Google Sheets: upload the file to Google Drive and choose Open with > Google Sheets.')
    desc = (meta["intro"] + "\n\nWHAT'S INSIDE\n" + "\n".join("• " + f for f in meta["features"]) +
            "\n\nWORKS WITH\nMicrosoft Excel (2010 or newer, Windows and Mac) and Google Sheets. No macros.\n\n"
            "YOU GET\nTwo files: a blank copy ready for your data and a copy with sample data so you can see how it works.\n\n"
            "One-time payment, instant download. Questions: muhakbaryusuf.ppp@gmail.com. Not happy? Ask for a refund within 14 days.")
    html = (f"<p>{meta['intro']}</p><p><strong>What's inside</strong></p><ul>" + "".join(f"<li>{f.replace('&','&amp;')}</li>" for f in meta["features"]) +
            "</ul><p><strong>Works with</strong> Microsoft Excel (2010 or newer, Windows and Mac) and Google Sheets. No macros.</p>"
            "<p><strong>You get</strong> two files: a blank copy ready for your data and a copy with sample data so you can see how it works.</p>"
            "<p>One-time payment, instant download. Questions: muhakbaryusuf.ppp@gmail.com. Not happy? Ask for a refund within 14 days.</p>")
    return desc, html, note

def main(mod_name):
    mod = importlib.import_module(mod_name); meta = mod.META; fb = meta["file_base"]
    dest = os.path.join(HERE, "out", fb); os.makedirs(dest, exist_ok=True)
    blank = os.path.join(dest, f"{fb}.xlsx"); sample = os.path.join(dest, f"{fb}-SAMPLE.xlsx")
    mod.build(True, sample); mod.build(False, blank)
    n = recalc(sample); recalc(blank)
    with zipfile.ZipFile(os.path.join(dest, f"{fb}.zip"), "w", zipfile.ZIP_DEFLATED) as z:
        z.write(blank, f"{fb}.xlsx"); z.write(sample, f"{fb}-SAMPLE.xlsx")
    cover(sample, meta, dest)
    desc, html, note = listing(meta)
    with open(os.path.join(dest, "listing.md"), "w") as f:
        f.write(f"# {meta['name']}\n\nPrice: US${meta['price']} single payment · Tax category: Digital Graphic / Template · License keys: off\n\n"
                f"## Description\n\n{desc}\n\n## Thank-you / receipt text\n\n{note}\n\n## Etsy tags\n\n{', '.join(meta['tags'])}\n")
    json.dump({"name": meta["name"], "price": meta["price"], "description_html": html, "description_text": desc, "thank_you": note, "tags": meta["tags"],
               "intro": meta["intro"], "features": meta["features"], "summary": meta.get("summary", meta["tagline"])},
              open(os.path.join(dest, "listing.json"), "w"), indent=1)
    print(json.dumps({"ok": True, "dest": dest, "formulas": n}))

if __name__ == "__main__": main(sys.argv[1])
