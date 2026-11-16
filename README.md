# COA → Excel, browser version (GitHub Pages)

A static web page that runs `coa_to_excel.py` **inside the user's browser** using
[Pyodide](https://pyodide.org) (Python compiled to WebAssembly). COA files are processed in the
browser tab and are **never uploaded anywhere**. There is no server, nothing to install, and no
`.exe`.

## Open the page

Pick **one** of these. In each case, keep the window that is serving the page open while you use it.

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
   click **Process**, then click the download button.
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
- **Don't double-click `index.html`.** Opened as a file (`file://...`), the browser blocks the
  background worker and the page stays on *Starting…*. Always open it through `http://` as above.
- The page fetches Python (Pyodide) from `cdn.jsdelivr.net` when it loads, so B and C still need
  internet access to that site. If the page stops at *Starting…* or shows *Could not start*,
  that site is probably blocked on your network.
- `--bind 127.0.0.1` means only your own computer can open the page.

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

## Requirements on the user's PC

- A modern browser (Edge, Chrome or Firefox).
- Access to `github.io` and `cdn.jsdelivr.net`. If the CDN is blocked, host the Pyodide files
  in this repository and change `PYODIDE` in `worker.js`.
