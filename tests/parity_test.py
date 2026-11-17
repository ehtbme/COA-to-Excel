"""Check that the JavaScript converter (js/coa_core.js) reads COAs exactly like coa_to_excel.py.

    python parity_test.py PDF_FOLDER [--chars]

For every PDF in the folder, compares the Python script (pdfplumber) with the JavaScript port (pdf.js, run in
headless Chromium): the text lines it builds and the final Excel row. --chars also compares every character
position. Needs: pip install pdfplumber openpyxl playwright; playwright install chromium-headless-shell.
Run it after changing coa_to_excel.py or js/coa_core.js, on as many real COAs as you have.

Known, harmless difference: in a few PDFs the positions of atom labels inside chemical structure drawings differ
slightly (pdf.js and pdfminer disagree on one glyph width); text lines and Excel rows are not affected.
"""
import functools
import http.server
import importlib.util
import json
import shutil
import sys
import tempfile
import threading
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
import build_single  # noqa: E402

HARNESS = """<!doctype html><meta charset="utf-8">
<script src="fflate.js"></script><script src="std_fonts.js"></script><script src="coa_core.js"></script>
<script type="module">
import * as pdfjs from "./pdf.mjs";
pdfjs.GlobalWorkerOptions.workerPort = new Worker("./pdf.worker.js");
window.runBatch = async (names, withChars) => {
  const out = [];
  for (const name of names) {
    const r = { name };
    const bytes = new Uint8Array(await (await fetch("pdfs/" + encodeURIComponent(name))).arrayBuffer());
    try {
      const { chars, width } = await COA.pdfChars(pdfjs, bytes);
      if (withChars) r.chars = chars.map((c) => [c.text, c.x0, c.x1, c.top, c.bottom, c.size, c.upright]);
      const lines = COA.pageLines(chars);
      r.lines = lines.map((l) => l.text);
      try { r.rec = COA.parseCoa(lines, width, name); } catch (e) { r.rec = null; }
    } catch (e) { r.rec = null; r.lines = null; }
    out.push(r);
  }
  return out;
};
window.harnessReady = true;
</script>"""


def python_side(pdfs, with_chars):
    spec = importlib.util.spec_from_file_location("coa", HERE / "py" / "coa_to_excel.py")
    coa = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(coa)
    import pdfplumber
    out = {}
    for p in pdfs:
        r = {"lines": None, "rec": None}
        try:
            with pdfplumber.open(p) as pdf:
                pg = pdf.pages[0]
                if with_chars:
                    r["chars"] = [[c["text"], c["x0"], c["x1"], c["top"], c["bottom"], c["size"], c["upright"]] for c in pg.chars]
                r["lines"] = [ln["text"] for ln in coa.page_lines(pg)]
        except Exception:
            pass
        try:
            r["rec"] = coa.parse_coa(str(p), p.name)
        except Exception:
            pass
        out[p.name] = r
    return out


def js_side(pdfs, with_chars):
    from playwright.sync_api import sync_playwright
    with tempfile.TemporaryDirectory() as tmp:
        site = Path(tmp)
        (site / "pdfs").mkdir()
        for p in pdfs:
            shutil.copyfile(p, site / "pdfs" / p.name)
        (site / "index.html").write_text(HARNESS)
        (site / "fflate.js").write_text(build_single.vendor("fflate.umd.js"))
        (site / "pdf.mjs").write_text(build_single.vendor("pdf.min.mjs"))
        (site / "pdf.worker.js").write_text(build_single.patch_pdfjs_worker(build_single.vendor("pdf.worker.min.mjs")))
        for f in ("std_fonts.js", "coa_core.js"):
            shutil.copyfile(HERE / "js" / f, site / f)
        class Quiet(http.server.SimpleHTTPRequestHandler):
            def log_message(self, *args):
                pass
        handler = functools.partial(Quiet, directory=str(site))
        server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        out = {}
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(f"http://127.0.0.1:{server.server_port}/index.html")
            page.wait_for_function("window.harnessReady === true")
            names = [p.name for p in pdfs]
            for i in range(0, len(names), 40):
                for r in page.evaluate("([n, c]) => runBatch(n, c)", [names[i:i + 40], with_chars]):
                    out[r.pop("name")] = r
            browser.close()
        server.shutdown()
    return out


def main():
    folder = Path(sys.argv[1])
    with_chars = "--chars" in sys.argv
    pdfs = sorted(p for p in folder.iterdir() if p.suffix.lower() == ".pdf")
    print(f"{len(pdfs)} PDFs: running Python…")
    py = python_side(pdfs, with_chars)
    print("running JavaScript in Chromium…")
    js = js_side(pdfs, with_chars)
    bad_lines = [n for n in py if py[n]["lines"] != js[n]["lines"]]
    bad_rows = [n for n in py if py[n]["rec"] != js[n]["rec"]]
    for n in bad_rows[:10]:
        p, j = py[n]["rec"], js[n]["rec"]
        diff = {k: (p[k], j.get(k)) for k in p if p[k] != j.get(k)} if p and j else (p is not None, j is not None)
        print(f"  ROW DIFFERS  {n}: {json.dumps(diff, ensure_ascii=False)[:300]}")
    print(f"text lines identical: {len(pdfs) - len(bad_lines)}/{len(pdfs)}")
    print(f"Excel rows identical: {len(pdfs) - len(bad_rows)}/{len(pdfs)}")
    if with_chars:
        n = moved = 0
        for name in py:
            for a, b in zip(py[name].get("chars") or [], js[name].get("chars") or []):
                n += 1
                moved += a[6] != b[6] or any(abs(x - y) > 1e-6 for x, y in zip(a[1:6], b[1:6]))
        print(f"character positions identical: {n - moved}/{n}")
    sys.exit(1 if bad_rows else 0)


if __name__ == "__main__":
    main()
