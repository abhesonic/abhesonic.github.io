"""Etsy listing assets: python3 tools/etsy_assets.py <module>  ->  out/<file_base>/etsy/ (img1-3.jpg, etsy.json)
Butuh out/<file_base>/ hasil make.py lebih dulu."""
import sys, os, glob, json, subprocess, tempfile, shutil, importlib, re
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "templates")); sys.path.insert(0, HERE)
FB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; FR = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
W, H = 2000, 1500  # 4:3, ukuran yang disarankan Etsy


def font(p, s): return ImageFont.truetype(p, s)


def wrap(d, text, f, maxw):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if d.textlength(t, font=f) <= maxw: cur = t
        else: lines.append(cur); cur = w
    lines.append(cur); return lines


def features_card(meta, dest):
    im = Image.new("RGB", (W, H), "white"); d = ImageDraw.Draw(im)
    d.rectangle((0, 0, W, 260), fill=meta["color"])
    d.text((90, 70), "What's inside", font=font(FB, 92), fill="white")
    y = 330; fb = font(FR, 46)
    for ft in meta["features"][:7]:
        lines = wrap(d, ft, fb, W - 300)
        d.ellipse((95, y + 14, 125, y + 44), fill=meta["color"])
        for ln in lines[:2]:
            d.text((160, y), ln, font=fb, fill="#1F2937"); y += 60
        y += 34
        if y > H - 230: break
    d.rectangle((0, H - 170, W, H), fill="#F1F5F9")
    d.text((90, H - 125), "Excel + Google Sheets  ·  Blank copy + sample copy  ·  Instant download", font=font(FB, 44), fill="#334155")
    im.save(os.path.join(dest, "img2.jpg"), quality=88)


def sheet_shot(sample, meta, dest):
    tmp = tempfile.mkdtemp(); shutil.copy(sample, os.path.join(tmp, "s.xlsx"))
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, os.path.join(tmp, "s.xlsx")], capture_output=True, timeout=240)
    pdf = os.path.join(tmp, "s.pdf")
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    page = 3 if n >= 3 else n
    subprocess.run(["pdftoppm", "-r", "170", "-png", "-f", str(page), "-l", str(page), pdf, os.path.join(tmp, "p")], check=True)
    shot = Image.open(glob.glob(os.path.join(tmp, "p*.png"))[0]).convert("RGB"); w, h = shot.size
    shot = shot.crop((int(w * .05), int(h * .06), int(w * .95), int(h * .72)))
    im = Image.new("RGB", (W, H), meta["color"]); d = ImageDraw.Draw(im)
    d.text((90, 60), "Fill in the yellow cells — formulas do the rest", font=font(FB, 64), fill="white")
    cw = W - 160; ch = int(shot.size[1] * cw / shot.size[0]); shot = shot.resize((cw, ch), Image.LANCZOS)
    ch = min(ch, H - 220); card = Image.new("RGB", (cw + 20, ch + 20), "white"); card.paste(shot.crop((0, 0, cw, ch)), (10, 10))
    im.paste(card, (70, 190)); im.save(os.path.join(dest, "img3.jpg"), quality=88)
    shutil.rmtree(tmp, ignore_errors=True)


def clean_tags(tags):
    out = []
    for t in tags:
        t = re.sub(r"[^A-Za-z0-9 '\-]", " ", t).strip(); t = re.sub(r"\s+", " ", t).lower()
        if 0 < len(t) <= 20 and t not in out: out.append(t)
    for extra in ["excel template", "google sheets", "spreadsheet", "digital download", "editable template"]:
        if len(out) >= 13: break
        if extra not in out: out.append(extra)
    return out[:13]


def etsy_title(meta):
    base = meta["title"].replace("&", "and")
    ACR = {"okr": "OKR", "tbr": "TBR", "crm": "CRM", "gpa": "GPA", "gci": "GCI", "rsvp": "RSVP", "fire": "FIRE", "pto": "PTO", "diy": "DIY"}
    skip = ("google sheets", "excel", "spreadsheet", "template", "digital download")
    kw = [k for k in meta["tags"] if k.lower() not in base.lower() and not any(x in k.lower() for x in skip)][:3]
    t = f"{base} Spreadsheet, Excel & Google Sheets Template"
    for k in kw:
        k = " ".join(ACR.get(w, w.capitalize()) for w in k.lower().split())
        cand = t + ", " + k
        if len(cand) <= 140: t = cand
    return t[:140]


def description(meta):
    return (meta["intro"] + "\n\n"
            "WHAT'S INSIDE\n" + "\n".join("• " + f for f in meta["features"]) + "\n\n"
            "WORKS WITH\n• Microsoft Excel 2010 or newer (Windows and Mac)\n• Google Sheets (free) – upload the file to Google Drive and choose Open with > Google Sheets\n"
            "• No macros, no add-ons\n\n"
            "YOU GET\n• 1 ZIP file with two spreadsheets: a blank copy ready for your data and a copy filled with sample data so you can see how everything works\n"
            "• Start Here tab with step-by-step instructions\n\n"
            "HOW IT WORKS\n1. Buy and download the file instantly from your Etsy Purchases page\n2. Unzip it and open the .xlsx file\n3. Type in the yellow cells – the grey formula cells and the dashboard update by themselves\n\n"
            "PLEASE NOTE\n• This is a digital download. No physical item will be shipped.\n"
            "• Designed by our shop with the help of AI tools. Every formula is checked before release.\n"
            "• For personal or single-business use. Please do not resell or share the file.\n\n"
            "Questions or something not working? Send us a message on Etsy and we will help.")


def main(mod_name):
    meta = importlib.import_module(mod_name).META; fb = meta["file_base"]
    src = os.path.join(HERE, "out", fb); dest = os.path.join(src, "etsy"); os.makedirs(dest, exist_ok=True)
    if "--json-only" not in sys.argv:
        build_images(meta, src, dest, fb)
    data = {"module": mod_name, "file_base": fb, "title": etsy_title(meta), "price": meta["price"], "tags": clean_tags(meta["tags"]),
            "description": description(meta), "images": ["img1.jpg", "img2.jpg", "img3.jpg"], "file": f"{fb}.zip"}
    json.dump(data, open(os.path.join(dest, "etsy.json"), "w"), indent=1)
    print(json.dumps({"ok": True, "title": data["title"], "len": len(data["title"]), "tags": len(data["tags"])}))


def build_images(meta, src, dest, fb):
    Image.open(os.path.join(src, "cover.png")).convert("RGB").resize((W, H), Image.LANCZOS).save(os.path.join(dest, "img1.jpg"), quality=88)
    features_card(meta, dest); sheet_shot(os.path.join(src, f"{fb}-SAMPLE.xlsx"), meta, dest)
    shutil.copy(os.path.join(src, f"{fb}.zip"), os.path.join(dest, f"{fb}.zip"))


if __name__ == "__main__": main(sys.argv[1])
