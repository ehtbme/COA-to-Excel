"""End-to-end test of the static site in a real (headless) Chromium.

Serves the site on 127.0.0.1, uploads files through the page, and compares the downloaded
Excel files with what coa_to_excel.py produces natively.

Usage: python browser_test.py SITE_DIR REFERENCE_PYTHON FILE [FILE ...]
(needs: pip install playwright openpyxl; playwright install chromium-headless-shell)
"""
import functools
import http.server
import io
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

site, ref_python, inputs = Path(sys.argv[1]), sys.argv[2], [Path(p) for p in sys.argv[3:]]
script = site / "py" / "coa_to_excel.py"


def rows(xlsx_bytes):
    return [r for r in openpyxl.load_workbook(io.BytesIO(xlsx_bytes)).active.iter_rows(values_only=True)]


def native(path):
    """Reference: run the script with normal Python on the same input."""
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(path)
        if src.suffix.lower() == ".pdf":
            z = Path(tmp) / "in.zip"
            with zipfile.ZipFile(z, "w") as zf:
                zf.write(src, arcname=src.name)
            src = z
        subprocess.run([ref_python, str(script), "-o", f"{tmp}/out", "--", str(src)],
                       check=True, capture_output=True)
        return rows(Path(f"{tmp}/out/COA_inventory.xlsx").read_bytes())


handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(site))
handler.log_message = lambda *a: None
server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
threading.Thread(target=server.serve_forever, daemon=True).start()
url = f"http://127.0.0.1:{server.server_port}/index.html"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(accept_downloads=True)
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    t0 = time.time()
    page.goto(url)
    page.wait_for_function("document.getElementById('status').textContent.startsWith('Ready')", timeout=300000)
    print(f"page ready (Python loaded) in {time.time() - t0:.1f}s")
    page.set_input_files("#files", [str(f) for f in inputs])
    page.click("#go")
    page.wait_for_function("/ready \\(|No files could/.test(document.getElementById('status').textContent)",
                           timeout=900000)
    print("status:", page.inner_text("#status"))
    for line in page.inner_text("#results table").splitlines():
        print("  |", line)
    downloads = {}
    for a in page.locator("#results a").all():
        with page.expect_download() as d:
            a.click()
        downloads[d.value.suggested_filename] = Path(d.value.path()).read_bytes()
    browser.close()
server.shutdown()

print("downloads:", sorted(downloads))
failures = 0
for src in inputs:
    name = f"{src.stem}_COA_inventory.xlsx"
    if name not in downloads:
        continue
    got, want = rows(downloads[name]), native(src)
    same = got == want
    failures += not same
    print(f"{name}: {len(got) - 1} rows, {'IDENTICAL to native run' if same else 'DIFFERENT from native run'}")
if "COA_inventory_files.zip" in downloads:
    with zipfile.ZipFile(io.BytesIO(downloads["COA_inventory_files.zip"])) as z:
        print("all-zip contains:", z.namelist())
print("page errors:", errors or "none")
sys.exit(1 if failures or errors else 0)
