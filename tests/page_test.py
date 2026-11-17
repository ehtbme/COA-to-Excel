"""End-to-end test of the page, opened the way colleagues open it: as a double-clicked file, offline.

    python page_test.py OUTDIR FILE [FILE ...]

Opens dist/COA_to_Excel.html from disk in headless Chromium with ALL network access blocked, converts the
given PDFs/zips, saves the downloads to OUTDIR and reports any network request or page error.
Needs: pip install playwright; playwright install chromium-headless-shell.
"""
import sys, time
from pathlib import Path
from playwright.sync_api import sync_playwright
url = (Path(__file__).resolve().parent.parent / "dist" / "COA_to_Excel.html").as_uri()
out, files = Path(sys.argv[1]), sys.argv[2:]
out.mkdir(exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(accept_downloads=True, viewport={"width": 1280, "height": 900})
    net = []
    ctx.route("**/*", lambda route: (net.append(route.request.url), route.abort()) if not route.request.url.startswith(("file:", "blob:", "data:")) else route.continue_())
    pg = ctx.new_page()
    errs = []; pg.on("pageerror", lambda e: errs.append(str(e))); pg.on("console", lambda m: m.type == "error" and errs.append(m.text))
    t0 = time.time(); pg.goto(url)
    pg.wait_for_function("/Ready|Could not/.test(document.getElementById('status').textContent)", timeout=120000)
    print(f"load: {pg.inner_text('#status')!r} in {time.time()-t0:.1f}s")
    pg.set_input_files("#files", files); pg.click("#go")
    seen = set()
    t1 = time.time()
    while True:
        s = pg.inner_text("#status")
        if time.time() - t1 > float(__import__("os").environ.get("LIMIT", "240")):
            print("TIMEOUT; status:", s, "| errors:", errs, "| net:", net); break
        if "COA" in s and len(seen) < 3: seen.add(s)
        if any(k in s for k in ("Done in", "No files could", "Error")): break
        pg.wait_for_timeout(200)
    print("progress samples:", sorted(seen)[:3])
    print("run:", pg.inner_text("#status").replace("\n", " "))
    print(pg.inner_text("#results table"))
    pg.screenshot(path=str(out / "results.png"), full_page=True)
    for a in pg.locator("#results a").all():
        with pg.expect_download() as d: a.click()
        d.value.save_as(str(out / d.value.suggested_filename)); print("  download", d.value.suggested_filename)
    print("network requests attempted:", net or "none")
    print("page errors:", errs or "none")
    b.close()
