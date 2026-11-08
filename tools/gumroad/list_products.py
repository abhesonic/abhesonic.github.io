#!/usr/bin/env python3
"""Print id / name / short_url / custom_permalink of every product (debug, read-only), following pagination."""
import json, os, urllib.parse, urllib.request
tok = os.environ["GUMROAD_TOKEN"]
extra, total = {}, 0
for _ in range(50):
    q = urllib.parse.urlencode({"access_token": tok, **extra})
    r = json.loads(urllib.request.urlopen(f"https://api.gumroad.com/v2/products?{q}", timeout=60).read())
    print("keys:", sorted(k for k in r if k != "products"))
    for p in r.get("products", []):
        total += 1
        print(p.get("id"), "|", p.get("name", "")[:60], "|", p.get("short_url"), "|", p.get("custom_permalink"), "|", p.get("published"))
    key = r.get("next_page_key")
    if not key:
        break
    extra = {"page_key": key}
print("TOTAL", total)
