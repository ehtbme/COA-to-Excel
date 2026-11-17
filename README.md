# COA → Excel, browser version

Converts COA PDFs into the inventory Excel sheet **inside the browser**, with plain JavaScript.
COA files are processed in the browser tab and are **never uploaded anywhere**. No server, no install,
no internet needed once the page is open.

It gives the same results as `coa_to_excel.py`: same Excel rows and the same COAs renamed `<lot>_<expiry>.pdf`
(tested on 583 real COAs, see *How it stays identical* below).

## Use it

Either:

- **Online:** open **https://ehtbme.github.io/COA-to-Excel/** in Edge or Chrome.
- **Offline, one file:** download `dist/COA_to_Excel.html` (2 MB) and **double-click it**. Send it to
  colleagues by email or Teams, or put it on SharePoint. It opens in Edge or Chrome and needs nothing else:
  no internet, no server, no install.

Both are the same page. Choose or drag in COA PDFs or zips of PDFs → **Convert to Excel** → download
`COA_inventory_files.zip` (Excel sheet(s) + the renamed PDFs) or just the Excel sheet.

## Files

| Path | What it is |
|---|---|
| `index.html` | The page served by GitHub Pages (built by `build_single.py`) |
| `dist/COA_to_Excel.html` | The same page, to download and double-click |
| `src/page.html` | The page design (edit this, not `index.html`) |
| `js/coa_core.js` | The converter: a JavaScript port of `coa_to_excel.py` plus the web app's upload checks |
| `js/std_fonts.js` | pdfminer's metrics for the 14 standard PDF fonts (keeps results identical) |
| `js/app_local.js` | Connects the page to the converter |
| `vendor/` | pdf.js 4.10.38 (Mozilla, Apache-2.0) and fflate 0.8.2 (MIT), with their licences |
| `py/coa_to_excel.py` | Copy of the Python script, the reference for `tests/parity_test.py` |
| `build_single.py` | Builds `index.html` and `dist/COA_to_Excel.html` (checks the `vendor/` files' SHA-256 first) |
| `tests/parity_test.py` | Compares the JavaScript converter with `coa_to_excel.py` on a folder of PDFs |
| `tests/page_test.py` | Opens the single file offline in headless Chromium and converts files |
| `.nojekyll` | Tells GitHub Pages to serve the files as they are |

## How it stays identical to the Python script

`js/coa_core.js` repeats `coa_to_excel.py` step by step: the same character positions as pdfplumber
(including its removal of overprinted "fake bold" text and pdfminer's font measurements), the same line
building, regular expressions, date formats, NA rules, file renaming and Excel layout. pdf.js gets two small,
checked patches at build time (see `build_single.py`).

Tested on 583 real COAs (`COA_downloads` + the Cayman test zips, 139 of them encrypted):

- Excel rows and text lines: **identical for all 583**; renamed PDFs byte-for-byte identical.
- Character positions: identical for 99.9 %. The rest are atom labels inside chemical structure drawings in
  69 files (pdf.js and pdfminer disagree on one glyph width); no text line or field changes.
- Upload errors (bad PDF/zip, no PDFs, wrong file type, unreadable PDFs) give the same messages as the
  web and desktop apps.

## After changing the converter

**If you change `coa_to_excel.py`, make the same change in `js/coa_core.js`**, then:

```bash
cp ../coa_to_excel/coa_to_excel.py py/          # keep the reference copy up to date
python3 build_single.py                         # rebuild index.html + dist/COA_to_Excel.html
python tests/parity_test.py <folder of COA PDFs>  # must say "Excel rows identical: N/N"
# then commit and push; GitHub Pages updates within a minute or two
```

To change the page's look, edit `src/page.html` and run `python3 build_single.py`.

## Publish on GitHub Pages

On GitHub: **Settings → Pages → Build and deployment → Deploy from a branch** → `main` / `(root)`.
On a free account the repository and site are **public** (the page contains code only, no COA data).

## Requirements on the user's PC

A modern Edge or Chrome. Nothing else.
