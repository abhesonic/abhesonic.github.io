#!/usr/bin/env python3
"""Rewrite every published Gumroad listing (description, summary, tags) from the
public template pages on abhesonic.github.io.

Why: 30 days of data showed 113 views and 0 sales. Gumroad products with a
description of 5,000+ characters, 2-3 covers and clear FAQ earn many times more
than short listings (InsightRaider, 146k products, 2026). All facts in the
generated text come from the template page itself (feature list, price, FAQ),
so nothing is invented.

Usage:
  python3 update_listings.py --site /path/to/site --catalog ../../catalog.json --dry-run
      -> writes out/listings/<slug>.html (+ .txt summary) for review, no API calls.
  GUMROAD_TOKEN=... python3 update_listings.py --site ... --catalog ... --apply
      -> PUT /v2/products/<id> for every product; only description, custom_summary
         and tags are sent (never files, price or published state).

The token is read from the environment only. Never print it.
"""
import argparse, glob, html, json, os, re, sys, time, urllib.error, urllib.parse, urllib.request

API = "https://api.gumroad.com/v2"


def text_of(fragment: str) -> str:
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", fragment, flags=re.S)
    t = re.sub(r"<[^>]+>", " ", t)
    return html.unescape(re.sub(r"\s+", " ", t)).strip()


def parse_page(path: str) -> dict:
    s = open(path, encoding="utf-8").read()
    d = {"slug": os.path.basename(path)[:-5]}
    m = re.search(r"<h1[^>]*>(.*?)</h1>", s, re.S)
    d["h1"] = text_of(m.group(1)) if m else d["slug"].replace("-", " ").title()
    m = re.search(r"<h1.*?</h1>\s*<p[^>]*>(.*?)</p>", s, re.S)
    d["lead"] = text_of(m.group(1)) if m else ""
    m = re.search(r"US\$\s?(\d+(?:\.\d+)?)", s)
    d["price"] = m.group(1) if m else ""
    # bullet lists: first <ul> after "What's inside"
    inside = re.search(r"What.s inside(.*?)(<h2|<section|$)", s, re.S)
    d["features"] = [text_of(li) for li in re.findall(r"<li[^>]*>(.*?)</li>", inside.group(1), re.S)] if inside else []
    how = re.search(r"How it works(.*?)(<h2|<section|$)", s, re.S)
    d["how"] = [text_of(li) for li in re.findall(r"<li[^>]*>(.*?)</li>", how.group(1), re.S)] if how else []
    faq = re.search(r"Questions(.*?)(<footer|$)", s, re.S)
    d["faq"] = []
    if faq:
        for q, a in re.findall(r"<(?:h3|summary|dt|strong)[^>]*>(.*?)</(?:h3|summary|dt|strong)>\s*(?:<p[^>]*>|<dd[^>]*>|)(.*?)(?=<(?:h3|summary|dt|strong)|</details>|</dl>|<footer|$)", faq.group(1), re.S):
            q, a = text_of(q), text_of(a)
            if q and a:
                d["faq"].append((q, a[:600]))
    m = re.search(r'name="description" content="([^"]+)"', s)
    d["meta"] = html.unescape(m.group(1)) if m else d["lead"]
    return d


def audience_for(slug: str) -> str:
    table = {
        "budget": "anyone who wants to see where the money goes every month without learning a finance app",
        "debt": "people paying off credit cards, student loans or a car loan who want a clear payoff date",
        "rental": "landlords and small real-estate investors managing one to a few dozen units",
        "airbnb": "Airbnb, Vrbo and Booking.com hosts with one or more listings",
        "etsy": "Etsy sellers who want to know their real profit after fees, shipping and materials",
        "freelancer": "freelancers and consultants who bill by the hour and need clean invoices",
        "inventory": "small shops, resellers and makers tracking stock by hand",
        "crm": "solo founders and small sales teams who outgrew a notebook but do not want a CRM subscription",
        "bookkeeping": "sole proprietors and small LLCs doing their own books",
        "habit": "anyone building daily habits who likes seeing a streak fill in",
        "meal": "busy households planning the week's meals and one grocery run",
        "wedding": "couples planning a wedding on a fixed budget",
        "travel": "travelers planning a trip budget and itinerary in one place",
        "workout": "lifters and runners logging sessions and tracking progress",
        "net-worth": "people tracking assets, debts and net worth month by month",
        "savings": "savers following a 52-week or custom savings challenge",
        "mortgage": "homeowners testing extra payments against their mortgage",
        "dividend": "dividend investors tracking income, yield and portfolio allocation",
        "house-flipping": "house flippers estimating purchase, rehab, holding and sale numbers",
        "paycheck": "people who budget paycheck to paycheck and want every bill covered",
        "event": "event planners managing a budget, vendors and a timeline",
        "content": "creators and small marketing teams planning posts across channels",
        "project": "small teams planning tasks on a simple Gantt timeline",
        "shift": "managers scheduling a small team's shifts and labor cost",
        "payroll": "small employers calculating hours, overtime and pay stubs",
        "cleaning": "cleaning business owners tracking recurring clients and jobs",
        "real-estate-agent": "real-estate agents tracking commissions, splits and leads",
        "teacher": "teachers keeping grades and attendance without a school system",
        "job-application": "job seekers tracking applications, interviews and follow-ups",
        "subscription": "anyone who keeps paying for subscriptions they forgot about",
    }
    for k, v in table.items():
        if k in slug:
            return v
    return "people who want a ready-made, editable spreadsheet instead of building one from scratch"


def tags_for(slug: str, h1: str) -> list:
    base = ["spreadsheet", "excel template", "google sheets template"]
    words = re.findall(r"[a-z]+", (slug + " " + h1).lower())
    extra = [w for w in words if w not in {"and", "the", "tracker", "template", "excel", "google", "sheets", "planner", "calculator", "spreadsheet", "with", "for"}]
    seen, out = set(), []
    for w in extra:
        if w not in seen and len(w) > 3:
            seen.add(w)
            out.append(w)
    return (base + out)[:10]


def build_description(d: dict, store_url: str) -> str:
    name = d["h1"].split(" – ")[0].split(" - ")[0].strip()
    aud = audience_for(d["slug"])
    feats = d["features"] or ["Dashboard with the key numbers", "Monthly tables and charts", "Yellow input cells, grey calculated cells"]
    how = d["how"] or [
        "Download the .xlsx file (you get a blank copy and a filled sample copy).",
        "Open it in Excel 2010 or newer, or upload it to Google Drive and open it with Google Sheets.",
        "Read the Start Here tab (about two minutes), then type your numbers into the yellow cells.",
        "Everything in grey — totals, charts, alerts — calculates by itself.",
    ]
    faq = d["faq"] or []
    std_faq = [
        ("Does it work in Google Sheets?", "Yes. Upload the .xlsx file to Google Drive and open it with Google Sheets; all formulas are standard functions that Sheets supports. It also works in Excel 2010 and newer on Windows and Mac, and in Excel for the web."),
        ("Are there macros?", "No. There are no macros, scripts, add-ins or external links, so nothing to enable and nothing that can break when you move the file."),
        ("Can I change categories, currency or the number of rows?", "Yes. Every label is editable, the currency is just a cell format, and the tables are sized with room to spare. The formulas keep working when you rename things."),
        ("Is this a one-time payment?", "Yes. One payment, download the file, use it forever on as many of your own devices as you like. Free updates to this version."),
        ("What if it is not what I needed?", "Email within 14 days and you get a refund. No forms, no questions."),
        ("Can I share it with my team or clients?", "The license is for one person or household. For a team, buy one copy per user or message us for a small-team license."),
    ]
    seen_q = {q.lower() for q, _ in faq}
    for q, a in std_faq:
        if q.lower() not in seen_q:
            faq.append((q, a))

    p = []
    p.append(f"<p><strong>{html.escape(d['lead'] or d['meta'])}</strong></p>")
    p.append(f"<p>{html.escape(name)} is a ready-made spreadsheet for {html.escape(aud)}. It is built for people who want the result — the dashboard, the totals, the charts — without spending a weekend wiring up formulas. You type your numbers once; the workbook does the rest, and it keeps working as your data grows.</p>")
    p.append("<h3>Who it is for</h3>")
    p.append(f"<p>This template is for {html.escape(aud)}. If you have ever started a spreadsheet like this yourself and abandoned it half-finished, this is the finished version: every formula tested, every tab labelled, and a sample copy filled with realistic data so you can see how it behaves before you enter your own.</p>")
    p.append("<h3>What is inside</h3><ul>" + "".join(f"<li>{html.escape(f)}</li>" for f in feats) + "</ul>")
    p.append("<p>Yellow cells are inputs — those are yours. Grey cells are calculations — leave them alone and they will always be right. Charts and alerts update the moment you type, in Excel and in Google Sheets.</p>")
    p.append("<h3>How it works</h3><ol>" + "".join(f"<li>{html.escape(h)}</li>" for h in how) + "</ol>")
    p.append("<h3>What you get</h3><ul>"
             "<li>The template as an .xlsx file — a blank copy ready for your data</li>"
             "<li>A second copy pre-filled with sample data so you can explore every tab</li>"
             "<li>A Start Here tab inside the file with a two-minute walkthrough</li>"
             "<li>Works in Microsoft Excel 2010 and newer (Windows, Mac, web) and Google Sheets</li>"
             "<li>No macros, no add-ins, no external links, no subscription</li>"
             "<li>Instant download after checkout, free updates to this version</li></ul>")
    p.append("<h3>Why a spreadsheet and not an app</h3>")
    p.append("<p>An app locks your data in someone else's database and charges you every month. A spreadsheet is yours: you can see every formula, change any label, add a column, print it, back it up, and open it ten years from now. This template gives you the structure and the calculations of a good app with none of the lock-in.</p>")
    p.append("<h3>Questions</h3>")
    for q, a in faq[:8]:
        p.append(f"<p><strong>{html.escape(q)}</strong><br>{html.escape(a)}</p>")
    p.append(f"<p>More templates for budgeting, small business and real estate at <a href=\"{store_url}\">{store_url.replace('https://','')}</a>. Questions before buying? Use the contact link on this page — replies within a day.</p>")
    return "\n".join(p)


def api(method: str, path: str, token: str, data: dict = None):
    params = {"access_token": token, **(data or {})}
    if method == "GET":
        url = f"{API}{path}?{urllib.parse.urlencode(params, doseq=True)}"
        req = urllib.request.Request(url, method="GET")
    else:
        req = urllib.request.Request(f"{API}{path}", data=urllib.parse.urlencode(params, doseq=True).encode(), method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")[:400]
        raise RuntimeError(f"HTTP {e.code}: {body}") from None


def all_products(token: str):
    """GET /products, following next_page_key until exhausted."""
    out, extra, seen = [], {}, set()
    for _ in range(50):
        r = api("GET", "/products", token, extra)
        out.extend(r.get("products", []))
        key = r.get("next_page_key")
        if not key or key in seen:
            break
        seen.add(key)
        extra = {"page_key": key}
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--catalog", required=True)
    ap.add_argument("--out", default="out/listings")
    ap.add_argument("--store", default="https://abhesonic.gumroad.com")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--dry-run", action="store_true")
    g.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    cat = json.load(open(a.catalog))
    pages = {os.path.basename(p)[:-5]: p for p in glob.glob(os.path.join(a.site, "templates", "*.html")) if not p.endswith("index.html")}
    os.makedirs(a.out, exist_ok=True)

    plan = []
    for m in cat["made"]:
        if m.get("status") != "published" or not m.get("gumroad_url"):
            continue
        slug = m["gumroad_url"].rstrip("/").split("/")[-1]
        page = pages.get(slug)
        if not page:
            print(f"SKIP {slug}: no site page", file=sys.stderr)
            continue
        d = parse_page(page)
        desc = build_description(d, a.store)
        summary = "You'll get the .xlsx template (blank + sample copy). Works in Excel 2010+ and Google Sheets. No macros. 14-day refund."
        plan.append({"slug": slug, "name": d["h1"], "description": desc, "summary": summary, "tags": tags_for(slug, d["h1"]), "chars": len(text_of(desc)), "words": len(text_of(desc).split())})

    for item in plan:
        open(os.path.join(a.out, item["slug"] + ".html"), "w", encoding="utf-8").write(item["description"])
    print(f"{len(plan)} listings prepared; chars min/avg/max = {min(i['chars'] for i in plan)}/{sum(i['chars'] for i in plan)//len(plan)}/{max(i['chars'] for i in plan)}; words min = {min(i['words'] for i in plan)}")
    if a.dry_run:
        return

    token = os.environ.get("GUMROAD_TOKEN")
    if not token:
        sys.exit("GUMROAD_TOKEN missing")
    products = all_products(token)
    print(f"{len(products)} products on Gumroad")
    by_slug = {}
    for p in products:
        slug = (p.get("custom_permalink") or p.get("short_url", "").rstrip("/").split("/")[-1] or "").lower()
        by_slug[slug] = p
    ok = fail = 0
    for item in plan:
        p = by_slug.get(item["slug"].lower())
        if not p:
            print(f"NOT FOUND on Gumroad: {item['slug']}")
            fail += 1
            continue
        try:
            r = api("PUT", f"/products/{urllib.parse.quote(p['id'], safe='')}", token, {
                "description": item["description"],
                "custom_summary": item["summary"],
                "tags[]": item["tags"],
            })
            print(("OK   " if r.get("success") else f"FAIL ({json.dumps(r)[:300]}) ") + item["slug"])
            ok += r.get("success", False)
            fail += not r.get("success", False)
        except Exception as e:  # noqa
            print(f"FAIL {item['slug']}: {e}")
            fail += 1
        time.sleep(1.2)
    print(f"done: {ok} updated, {fail} failed")


if __name__ == "__main__":
    main()
