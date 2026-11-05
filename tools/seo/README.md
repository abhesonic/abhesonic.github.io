# SEO tools for the public site (branch main)

- `gen_pages.py <factory_dir> <site_dir>`: writes a landing page `templates/<slug>.html` (+ `templates/img/<slug>.jpg`) for every
  published template in `catalog.json` that has no page yet, adds a card to `templates/index.html` and a URL to `sitemap.xml`.
  Needs the cover first: run `python3 -W ignore make.py <module>` in the factory dir (creates `out/<file_base>/cover.jpg`).
- `related.json`: template slug -> list of [guide slug, link text]. Shown as "Related guide" on the template page.
  Add an entry when a new guide is written, and add the guide link by hand to the template page if the page already exists.
- New guides: copy the structure of `guides/habit-tracker-spreadsheet.html` on main (same head, CSS, Article JSON-LD,
  nav, footer). Add the guide at the top of the list in `guides/index.html` and to `sitemap.xml`.

Published guides (do not duplicate topics):
- extract-table-from-website-to-excel, download-all-images-from-a-webpage, free-vs-paid-budget-spreadsheet,
  rental-property-spreadsheet, debt-snowball-vs-avalanche (29 Oct 2026)
- 52-week-savings-challenge-spreadsheet, rental-property-analysis-cash-flow-cap-rate, habit-tracker-spreadsheet,
  net-worth-tracker-spreadsheet (5 Nov 2026)
