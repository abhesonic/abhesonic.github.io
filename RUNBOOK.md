# Template factory – daily run (2 templates per run, until 50)

Owner: Akbar (Bahasa Indonesia, beginner, wants autopilot). Store: Lemon Squeezy (LS) "BatchGrab", USD, test mode until LS activates the store.
This folder lives on the `factory` branch of github.com/abhesonic/abhesonic.github.io (never merge into main: main is the public website).

## Each run
1. Clone: `git clone -b factory https://github.com/abhesonic/abhesonic.github.io factory` (attach the repo with add_repo, access push, if needed).
2. Read `catalog.json`. Stop and report if `made` already has 50 entries.
3. Pick the next 2 ideas from `next` (skip any already made). A quick web search to confirm demand and sharpen features is fine.
4. For each: write `templates/<module>.py` with `build(sample, path)` and `META` (copy the structure of `small_business_bookkeeping.py`).
   Rules (xlsx skill): openpyxl, Arial, yellow input cells / grey formula cells, a Start Here sheet, Dashboard as the 2nd sheet,
   sample data in the SAMPLE copy only, Excel-2007 functions only (SUMIFS, COUNTIFS, INDEX/MATCH, IFERROR, SUMPRODUCT – never XLOOKUP/FILTER/UNIQUE/SORT/SEQUENCE),
   quote sheet names with spaces, no hardcoded results.
5. `python3 -W ignore make.py <module>` → must print ok with zero formula errors. Check a few computed values with openpyxl data_only and look at out/<file_base>/cover.jpg.
6. Upload to LS (needs Akbar's computer linked and online):
   - Copy `out/<file_base>/<file_base>.zip` and `cover.png` to `C:\Users\LENOVO\Downloads\Template-Factory\<file_base>\` (device_commit_files after SendUserFile, or device bridge).
   - In the browser (Lemon Squeezy must be logged in), Products → New Product: name, description (paste `description_html` from listing.json as a text/html paste event into the tiptap editor),
     Single payment, price, Tax category "Digital Graphic / Template", license keys off, Confirmation modal message and Email receipt "Thank you note" = `thank_you`.
     Upload cover.png to Media and the zip to Files (Chrome extension `file_upload` on the file inputs if available; otherwise leave as draft and list it for Akbar).
     Then click "Publish product". KNOWN QUIRK: the price typed into the new-product form is often saved as the default $9.99.
     After publishing, re-open the product (it stays open), triple-click the Price field, type the price (e.g. 29.00), press Tab, click "Save changes",
     then open /products and confirm the list shows the right price. Never enter passwords; if logged out, stop and tell Akbar.
     Chrome extension notes: file_upload works with container paths under /home/claude/... (e.g. /home/claude/factory/out/<file_base>/cover.png);
     the description is a tiptap editor: focus it and dispatch a paste ClipboardEvent with text/html via javascript_tool; log in state lives in Akbar's Chrome.
7. Update `catalog.json` (move the idea from `next` to `made` with date, price, LS id and status), commit, push to `factory`.
8. Message Akbar in Bahasa Indonesia, short: which 2 templates were made, LS status, anything he must click.
