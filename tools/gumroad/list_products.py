#!/usr/bin/env python3
"""Print id / name / short_url / custom_permalink of every product (debug, read-only)."""
import json, os, urllib.parse, urllib.request
tok = os.environ["GUMROAD_TOKEN"]
r = urllib.request.urlopen(urllib.request.Request("https://api.gumroad.com/v2/products", data=urllib.parse.urlencode({"access_token": tok}).encode(), method="GET"), timeout=60)
for p in json.loads(r.read())["products"]:
    print(p.get("id"), "|", p.get("name", "")[:60], "|", p.get("short_url"), "|", p.get("custom_permalink"), "|", p.get("published"))
