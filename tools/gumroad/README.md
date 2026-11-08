# Gumroad listing updater

`update_listings.py` rewrites description / summary / tags of every published template from its page on the public site
(branch main, templates/<slug>.html). Facts only from the page; ~700 words per listing. Never touches price, files or
published state. Extensions (TableGrab/BatchGrab) are not included — their listings are edited by hand.

Run from GitHub Actions ("Gumroad listings update", input apply=true) — the token lives only in the repo secret
`GUMROAD_TOKEN` (Settings → Secrets and variables → Actions). Create the token at
https://gumroad.com/settings/advanced → Applications → "Kantor AI listing updater" → Generate access token.
Locally: `python3 tools/gumroad/update_listings.py --site <main checkout> --catalog catalog.json --dry-run`.
