# Template factory – daily run (2 templates per run, until 50)

Owner: Akbar (Bahasa Indonesia, beginner, wants autopilot). Store: Lemon Squeezy (LS) "BatchGrab", USD, test mode until LS activates the store.
This folder lives on the `factory` branch of github.com/abhesonic/abhesonic.github.io (never merge into main: main is the public website).

## Each run
1. Clone: `git clone -b factory https://github.com/abhesonic/abhesonic.github.io factory` (attach the repo with add_repo, access push, if needed).
2. Read `catalog.json`. Stop and report if `made` already has 50 entries.
3. Pick the next 2 ideas from `next` (skip any already made). A quick web search to confirm demand and sharpen features is fine.
4. For each: write `templates/<module>.py` with `build(sample, path)` and `META` (copy the structure of `small_business_bookkeeping.py`, or reuse the shared helpers in `templates/_common.py` as `airbnb_host_tracker.py` does).
   Rules (xlsx skill): openpyxl, Arial, yellow input cells / grey formula cells, a Start Here sheet, Dashboard as the 2nd sheet,
   sample data in the SAMPLE copy only, Excel-2007 functions only (SUMIFS, COUNTIFS, INDEX/MATCH, IFERROR, SUMPRODUCT – never XLOOKUP/FILTER/UNIQUE/SORT/SEQUENCE),
   quote sheet names with spaces, no hardcoded results.
5. `python3 -W ignore make.py <module>` → must print ok with zero formula errors. Check a few computed values with openpyxl data_only and look at out/<file_base>/cover.jpg.
6. Upload to Gumroad (store: https://abhesonic.gumroad.com). Lemon Squeezy rejected the store on 2026-10-29: never upload there.
   Only when catalog.json `store_ready` is true and Akbar's Chrome (Claude in Chrome extension) is online and logged in to gumroad.com.
   JavaScript is NOT allowed on gumroad.com, so use find / click / type / file_upload only. file_upload accepts files under the session's working directory
   (e.g. /home/<you>/factory/out/<file_base>/...); if a path is rejected, try /mnt/user-data/outputs/<file_base>/ after copying the files there.
   a. Open https://gumroad.com/products/new. Click the Name field (it is the first text box), type META name. `find` "Price input", triple-click it, type the price.
      Click "Next: Customize" (top right). The product opens at /products/<id>/edit.
   b. Description: click inside the description editor and type the paragraphs one by one with an Enter key press between them:
      intro, "WHAT'S INSIDE", one "• feature" line per feature, then
      "WORKS WITH: Microsoft Excel (2010 or newer, Windows and Mac) and Google Sheets. No macros, no add-ons.",
      "YOU GET: two files – a blank copy ready for your data and a copy with sample data so you can see how everything works.",
      "One-time payment, instant download. Questions? Email muhakbaryusuf.ppp@gmail.com – not happy within 14 days? Ask for a refund."
   c. `find` the URL slug box, the "Upload images or videos" button (Cover), the Thumbnail file input and the Summary box.
      URL slug: click it and type the file_base in lower case (e.g. debt-payoff-planner).
      Thumbnail: file_upload out/<file_base>/thumb.png to the Thumbnail file input.
      Cover: click "Upload images or videos" (it opens a small menu, not a file dialog), `find` "Computer files file input under Cover",
      file_upload out/<file_base>/gumroad-cover.png to it.
      Summary: "Instant download: Excel + Google Sheets template, blank copy + sample copy."
   d. Click "Save and continue" (top right) → Content tab. `find` the file input (type=file) and file_upload out/<file_base>/<file_base>.zip.
      Wait ~7 s, click in the empty content area below the file and type:
      'How to start: unzip the download, open <file_base>.xlsx and begin on the "Start Here" tab. The -SAMPLE file shows the same template filled with example data.'
      Enter, then 'Google Sheets: go to drive.google.com, upload the .xlsx file, then right-click it and choose Open with > Google Sheets.'
      Click "Save changes".
   e. If the yellow banner "You haven't connected a payout method yet" is gone, click "Publish and continue" so the product goes live; otherwise leave it
      unpublished and tell Akbar to add his bank details (Settings > Payments). Then also publish the earlier unpublished products.
   Note: Gumroad allows at most 10 new products per day ("Sorry, you can only create 10 products per day"). If you see that, keep the template with status "not uploaded" and retry next run.
   f. Open https://gumroad.com/products and check the new product shows the right price and URL.
7. Update `catalog.json` (move the idea from `next` to `made` with date, price, gumroad_id, gumroad_url and status), commit, push to `factory`.
8. Message Akbar in Bahasa Indonesia, short: which 2 templates were made, Gumroad status (published or waiting), anything he must click.
