"""Generate landing pages for templates that have no page yet on the public site.

Usage: python3 gen_pages.py <factory_dir> <site_dir>
Reads catalog.json + templates/<module>.py META from the factory branch checkout,
writes templates/<slug>.html, templates/img/<slug>.jpg, adds cards to templates/index.html
and URLs to sitemap.xml. Pages without a built cover (out/<file_base>/cover.jpg) are skipped.
"""
import sys, os, json, html, importlib, datetime, re
from PIL import Image

factory, site = sys.argv[1], sys.argv[2]
sys.path.insert(0, os.path.join(factory, "templates"))
today = datetime.date.today().isoformat()
e = lambda s: html.escape(s, quote=True)

CSS = """:root{color-scheme:light dark;--bg:#f6f5f2;--surface:#fff;--border:#e3e0d8;--text:#1c1b19;--muted:#6e6a61;--accent:#15803d;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;line-height:1.6}
@media(prefers-color-scheme:dark){:root{--bg:#141413;--surface:#1d1c1a;--border:#34322e;--text:#edebe6;--muted:#a29e94;--accent:#4ade80}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text)}
main{max-width:860px;margin:0 auto;padding:40px 16px 60px}
nav{font-size:14px;margin-bottom:18px}nav a{color:var(--muted)}
h1{font-size:clamp(26px,4.5vw,36px);line-height:1.2;letter-spacing:-.02em;margin:0 0 10px}
.lead{font-size:18px;color:var(--muted);margin:0 0 20px}
img.cover{width:100%;border-radius:14px;border:1px solid var(--border);margin:8px 0 22px}
.btn{display:inline-block;background:var(--accent);color:#fff;text-decoration:none;font-weight:600;padding:12px 18px;border-radius:10px;margin:6px 0 18px}
h2{font-size:21px;margin:28px 0 8px}ul{padding-left:20px}li{margin:4px 0}
.faq p{margin:2px 0 12px}.faq b{display:block}
.meta{color:var(--muted);font-size:14px}
footer{max-width:860px;margin:0 auto;padding:0 16px 40px;color:var(--muted);font-size:14px}footer a{color:inherit}"""

def faqs(module_src):
    q = [("Does it work in Google Sheets?", "Yes. Upload the .xlsx file to Google Drive, open it with Google Sheets, and all formulas, drop-down lists and conditional formatting keep working. No macros are used."),
         ("Which Excel versions are supported?", "Excel 2010 or newer on Windows and Mac, including Microsoft 365. Numbers on iPad/Mac and LibreOffice also open it, with small cosmetic differences."),
         ("What do I get after paying?", "An instant download of a ZIP with two files: a blank copy ready for your numbers and a copy filled with sample data so you can see how every sheet works. Start on the \"Start Here\" tab.")]
    if re.search(r"currency", module_src, re.I):
        q.append(("Can I use my own currency?", "Yes. Pick the currency symbol once in Settings and every sheet follows."))
    q.append(("What if it is not for me?", "Email within 14 days and you get a refund."))
    return q

def page(meta, slug, url, price, q, related):
    name, intro = meta["name"], meta["intro"]
    short = name.split(" – ")[0]
    short_nav = re.sub(r"\s*\(.*?\)", "", short)
    title = f"{name} – US${price}"
    ld_prod = {"@context": "https://schema.org", "@type": "Product", "name": name,
               "image": f"https://abhesonic.github.io/templates/img/{slug}.jpg", "description": intro,
               "brand": {"@type": "Brand", "name": "abhesonic"},
               "offers": {"@type": "Offer", "url": url, "priceCurrency": "USD", "price": str(price), "availability": "https://schema.org/InStock"}}
    ld_faq = {"@context": "https://schema.org", "@type": "FAQPage",
              "mainEntity": [{"@type": "Question", "name": a, "acceptedAnswer": {"@type": "Answer", "text": b}} for a, b in q]}
    feats = "".join(f"<li>{e(f)}</li>" for f in meta["features"])
    faq_html = "".join(f"<p><b>{e(a)}</b>{e(b)}</p>" for a, b in q)
    rel = ""
    if related:
        rel = "<h2>Related guide</h2><p>" + " · ".join(f'<a href="../guides/{g}.html">{e(t)}</a>' for g, t in related) + "</p>"
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(title)}</title>
<meta name="description" content="{e(intro)}">
<link rel="canonical" href="https://abhesonic.github.io/templates/{slug}.html">
<meta property="og:title" content="{e(name)}"><meta property="og:description" content="{e(intro)}"><meta property="og:image" content="https://abhesonic.github.io/templates/img/{slug}.jpg">
<script type="application/ld+json">{json.dumps(ld_prod)}</script><script type="application/ld+json">{json.dumps(ld_faq)}</script>
<style>
{CSS}
</style></head><body><main>
<nav><a href="../">abhesonic</a> · <a href="./">Templates</a> · {e(short_nav)}</nav>
<h1>{e(name)}</h1>
<p class="lead">{e(intro)}</p>
<a class="btn" href="{url}">Get it on Gumroad – US${price}, one-time</a>
<p class="meta">Instant download · blank copy + sample copy · Excel 2010+ and Google Sheets · no macros · 14-day refund</p>
<img class="cover" src="img/{slug}.jpg" alt="{e(short)} dashboard preview">
<h2>What's inside</h2><ul>{feats}</ul>
<h2>How it works</h2>
<ol><li>Open the file in Excel, or upload it to Google Drive and open with Google Sheets.</li><li>Read the "Start Here" tab (2 minutes). Yellow cells are yours to fill in, grey cells calculate by themselves.</li><li>Type your own numbers; the dashboard, charts and alerts update as you go.</li></ol>
<a class="btn" href="{url}">Get {e(short_nav)} – US${price}</a>
{rel}
<h2>Questions</h2><div class="faq">{faq_html}</div>
<p class="meta">Keywords: {e(", ".join(meta.get("tags", [])))}</p>
</main><footer>More: <a href="./">all templates</a> · <a href="../guides/">guides</a> · <a href="../tablegrab/">TableGrab</a> · <a href="../batchgrab/">BatchGrab</a> · <a href="mailto:muhakbaryusuf.ppp@gmail.com">support</a> · <a href="https://abhesonic.gumroad.com/affiliates">affiliates (30%)</a></footer></body></html>
"""

def card(meta, slug, url, price):
    name = meta["name"]
    feats = "".join(f"<li>{e(f)}</li>" for f in meta["features"][:4])
    return f"""
    <article class="card" id="{slug}">
      <a href="{slug}.html"><img src="img/{slug}.jpg" alt="{e(name)} cover" loading="lazy"></a>
      <div class="body">
        <h2><a href="{slug}.html" style="color:inherit;text-decoration:none">{e(name)}</a></h2>
        <p class="sum">{e(meta["tagline"])}</p>
        <ul>{feats}</ul>
        <a class="btn" href="{url}">Get it on Gumroad – US${price}</a>
      </div>
    </article>"""

related_map = json.load(open(os.path.join(os.path.dirname(__file__), "related.json"))) if os.path.exists(os.path.join(os.path.dirname(__file__), "related.json")) else {}
cat = json.load(open(os.path.join(factory, "catalog.json")))
tdir = os.path.join(site, "templates"); os.makedirs(os.path.join(tdir, "img"), exist_ok=True)
have = {f[:-5] for f in os.listdir(tdir) if f.endswith(".html")}
idx_path = os.path.join(tdir, "index.html"); idx = open(idx_path, encoding="utf-8").read()
smap_path = os.path.join(site, "sitemap.xml"); smap = open(smap_path, encoding="utf-8").read()
made = []
for m in cat["made"]:
    if m.get("status") != "published": continue
    url = m["gumroad_url"]; slug = url.rsplit("/", 1)[1]
    if slug in have: continue
    cov = os.path.join(factory, "out", m["file_base"], "cover.jpg")
    if not os.path.exists(cov): print("skip (no cover):", slug); continue
    mod = importlib.import_module(m["module"]); meta = mod.META
    src = open(mod.__file__, encoding="utf-8").read()
    im = Image.open(cov).convert("RGB"); im.thumbnail((1200, 900)); im.save(os.path.join(tdir, "img", f"{slug}.jpg"), quality=82, optimize=True)
    rel = [(g, t) for g, t in related_map.get(slug, [])]
    open(os.path.join(tdir, f"{slug}.html"), "w", encoding="utf-8").write(page(meta, slug, url, m["price"], faqs(src), rel))
    marker = "</div>\n<p class=\"note\">"
    assert marker in idx, "index marker not found"
    idx = idx.replace(marker, card(meta, slug, url, m["price"]) + marker, 1)
    loc = f"https://abhesonic.github.io/templates/{slug}.html"
    if loc not in smap:
        smap = smap.replace("</urlset>", f"<url><loc>{loc}</loc><lastmod>{today}</lastmod></url>\n</urlset>")
    made.append(slug)
open(idx_path, "w", encoding="utf-8").write(idx)
open(smap_path, "w", encoding="utf-8").write(smap)
print("pages:", len(made), made)
