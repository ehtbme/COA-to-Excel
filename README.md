# COA → Excel, browser version (GitHub Pages)

A static web page that runs `coa_to_excel.py` **inside the user's browser** using
[Pyodide](https://pyodide.org) (Python compiled to WebAssembly). COA files are processed in the
browser tab and are **never uploaded anywhere**. There is no server, nothing to install, and no
`.exe`.

## Files

| Path | What it is |
|---|---|
| `index.html` | The page: choose or drop files → Process → Download |
| `worker.js` | Loads Python (Pyodide 0.26.4 from cdn.jsdelivr.net) in a background thread |
| `py/coa_to_excel.py` | Your converter, copied unchanged by `build.sh` |
| `py/browser_glue.py` | Same checks as the web/desktop apps (file types, zip limits, failures) |
| `wheels/` | pdfplumber, pdfminer.six, openpyxl, et_xmlfile (pure Python). Pyodide provides cryptography, Pillow and charset-normalizer. |
| `.nojekyll` | Tells GitHub Pages to serve the files as they are |
| `build.sh` | Refreshes `py/coa_to_excel.py` and `wheels/` |
| `tests/browser_test.py` | Headless-Chromium end-to-end test (not needed on the site) |

## Publish on GitHub Pages

1. Create a repository and push **only this folder's contents** (no COA PDFs or Excel files).
2. On GitHub: **Settings → Pages → Build and deployment → Deploy from a branch** → `main` / `(root)`.
3. The site appears at `https://<user>.github.io/<repo>/` after a minute or two.

On a free account, the site and its code are **public**.

## Update after changing the converter

```bash
bash build.sh                 # copies ../coa_to_excel/coa_to_excel.py, re-downloads wheels
# then commit and push
```

## Test locally

```bash
python3 -m http.server 8000 --bind 127.0.0.1    # in this folder, then open http://127.0.0.1:8000
```

You can't open `index.html` directly as a file (`file://`): browsers block the background worker
there. Serve it over http as above, or use GitHub Pages.

## Requirements on the user's PC

- A modern browser (Edge, Chrome or Firefox).
- Access to `github.io` and `cdn.jsdelivr.net`. If the CDN is blocked, host the Pyodide files
  in this repository and change `PYODIDE` in `worker.js`.
