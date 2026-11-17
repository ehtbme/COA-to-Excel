# COA → Excel, browser version

Converts COA PDFs to the inventory Excel sheet **inside the user's browser**. COA files are processed
in the browser tab and are **never uploaded anywhere**. There is no server, nothing to install, and no
`.exe`. Two forms:

- **`dist/COA_to_Excel.html`: one file, plain JavaScript, works offline.** The converter is a JavaScript
  port of `coa_to_excel.py`, tested to give identical results (see *Single file* below).
- **The hosted page** (`index.html`, GitHub Pages): runs the real `coa_to_excel.py` with
  [Pyodide](https://pyodide.org) (Python in the browser), downloaded from `cdn.jsdelivr.net` when it opens.

## Open the page

Pick **one** of these. For B and C, keep the window that is serving the page open while you use it.

### Easiest: one HTML file, double-click it
`dist/COA_to_Excel.html` (2 MB) is the whole tool in a single file. Send it to colleagues by
email or Teams, or put it on SharePoint. They save it anywhere and **double-click it**; it opens in
Edge or Chrome and is ready in under a second. **No internet needed**, no server, no folder, no install:
it makes no network requests at all.

### A. Online (nothing to start)
Open **https://ehtbme.github.io/COA-to-Excel/** in Edge or Chrome.
If it says *"Your connection isn't private"*, the network is blocking GitHub; use B instead.

### B. On your Windows laptop (Anaconda)
1. Copy this whole folder to the laptop, e.g. `C:\Users\you\COA-to-Excel\`
   (or on GitHub: **Code → Download ZIP**, then unzip).
2. Open **Anaconda Prompt** and run:
   ```
   cd C:\Users\you\COA-to-Excel
   python -m http.server 8000 --bind 127.0.0.1
   ```
   It prints `Serving HTTP on 127.0.0.1 port 8000 ...`. Leave this window open.
3. In Edge or Chrome, open **http://127.0.0.1:8000**
4. Wait until the page says **Ready.** (about 10-30 s the first time), choose your COA files,
   click **Convert to Excel**, then click a download button.
5. To stop: go back to Anaconda Prompt and press **Ctrl+C**.

### C. On the Linux VM, viewed from VS Code
1. In the VS Code terminal (connected to the VM):
   ```bash
   cd ~/Projects/ad-hoc/inventory/coa_browser
   python3 -m http.server 8000 --bind 127.0.0.1
   ```
2. VS Code shows *"Your application running on port 8000 is available"* → **Open in Browser**.
   If it doesn't, open the **Ports** tab, add port `8000`, and click the globe icon.
3. Stop with **Ctrl+C** in that terminal.

### Notes
- **Don't double-click `index.html`** (double-click `dist/COA_to_Excel.html` instead). Opened as a file
  (`file://...`), the browser blocks its background worker and the page stays on *Starting…*.
- The page fetches Python (Pyodide) from `cdn.jsdelivr.net` when it loads, so B and C still need
  internet access to that site. If the page stops at *Starting…* or shows *Could not start*,
  that site is probably blocked on your network.
- `--bind 127.0.0.1` means only your own computer can open the page.

## Files

| Path | What it is |
|---|---|
| `index.html` | The page: choose or drop files → Convert to Excel → Download |
| `worker.js` | Loads Python (Pyodide 0.26.4 from cdn.jsdelivr.net) in a background thread |
| `py/coa_to_excel.py` | Your converter, copied unchanged by `build.sh` |
| `py/browser_glue.py` | Same checks as the web/desktop apps (file types, zip limits, failures) |
| `wheels/` | pdfplumber, pdfminer.six, openpyxl, et_xmlfile (pure Python). Pyodide provides cryptography, Pillow and charset-normalizer. |
| `.nojekyll` | Tells GitHub Pages to serve the files as they are |
| `build.sh` | Refreshes `py/coa_to_excel.py` and `wheels/` |
| `dist/COA_to_Excel.html` | **The single-file version** (built by `build_single.py`) |
| `js/coa_core.js` | The JavaScript converter: a port of `coa_to_excel.py` + `browser_glue.py` |
| `js/std_fonts.js` | pdfminer's metrics for the 14 standard PDF fonts (keeps results identical) |
| `js/app_local.js` | Connects the page to the JavaScript converter in the single file |
| `vendor/` | pdf.js 4.10.38 (Mozilla, Apache-2.0) and fflate 0.8.2 (MIT), with their licences |
| `build_single.py` | Builds `dist/COA_to_Excel.html` (checks the `vendor/` files' SHA-256 first) |
| `tests/parity_test.py` | Compares the JavaScript converter with `coa_to_excel.py` on a folder of PDFs |
| `tests/browser_test.py` | Headless-Chromium end-to-end test of the hosted page |

## Publish on GitHub Pages

1. Create a repository and push **only this folder's contents** (no COA PDFs or Excel files).
2. On GitHub: **Settings → Pages → Build and deployment → Deploy from a branch** → `main` / `(root)`.
3. The site appears at `https://<user>.github.io/<repo>/` after a minute or two.

On a free account, the site and its code are **public**.

## Single file: how it stays identical to the Python script

`js/coa_core.js` repeats `coa_to_excel.py` step by step: the same character positions as pdfplumber
(including its de-duplication of overprinted bold text and pdfminer's font measurements), the same
line building, regular expressions, date formats, NA rules, file renaming and Excel layout.
pdf.js gets two small, checked patches at build time (see `build_single.py`).

Tested on 583 real COAs (`COA_downloads` + the Cayman test zips, 139 of them encrypted):

- Excel rows and text lines: **identical for all 583**; renamed PDFs byte-for-byte identical.
- Character positions: identical for 99.9 %. The rest are atom labels inside chemical structure
  drawings in 69 files (pdf.js and pdfminer disagree on one glyph width); no text line or field changes.
- Upload errors (bad PDF/zip, no PDFs, wrong file type, unreadable PDFs) give the same messages.

**If you change `coa_to_excel.py`, make the same change in `js/coa_core.js`**, then:

```bash
bash build.sh                                   # refresh py/coa_to_excel.py for the hosted page
python3 build_single.py                         # rebuild dist/COA_to_Excel.html
python tests/parity_test.py <folder of COA PDFs>  # must say "Excel rows identical: N/N"
```

## Requirements on the user's PC

- A modern browser (Edge or Chrome; tested in Chromium).
- Single file: nothing else. Hosted page: access to `github.io` and `cdn.jsdelivr.net`.
