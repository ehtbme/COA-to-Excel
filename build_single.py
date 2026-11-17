"""Build dist/COA_to_Excel.html: the whole tool in ONE file, plain JavaScript, nothing downloaded.

    python3 build_single.py

Double-click the result: it opens in Edge/Chrome and works with no internet, no server and no install.
It packs index.html's page with:
  js/coa_core.js      the converter (a JavaScript port of ../coa_to_excel/coa_to_excel.py)
  js/std_fonts.js     pdfminer's metrics for the 14 standard PDF fonts (keeps results identical to Python)
  js/app_local.js     connects the page to the converter
  vendor/             pdf.js 4.10.38 (Mozilla, Apache-2.0) and fflate 0.8.2 (MIT), checked by SHA-256 below

pdf.js's background script gets two small patches (patch_pdfjs_worker): it also passes on each font's
/Descent value from the PDF (pdfminer uses it for character positions; pdf.js normally replaces it with the
font file's value), and its final `export` line is removed so it runs from a double-clicked file.
After changing coa_to_excel.py or js/coa_core.js, run tests/parity_test.py to compare the two on real COAs.
"""
import hashlib
import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "dist" / "COA_to_Excel.html"
VENDOR = {
    "pdf.min.mjs": "44ec6f011027ee77791386b66c14876a5fc29e20bf0433c07c6726fff7212b72",
    "pdf.worker.min.mjs": "bd88805178a26c729db8c0107a5b630cb900ec070f4d8c7529a3e45530afd41d",
    "fflate.umd.js": "c3b34f2e9f5e74d4d7d64e01cac7a0c01954c6c406414d42185c7b53d6875ddf",
}


def vendor(name: str) -> str:
    data = (HERE / "vendor" / name).read_bytes()
    if hashlib.sha256(data).hexdigest() != VENDOR[name]:
        raise SystemExit(f"vendor/{name} does not match its checksum; refusing to build")
    return data.decode("utf-8")


def _sub_once(pattern: str, repl, text: str) -> str:
    new, n = re.subn(pattern, repl, text)
    if n != 1:
        raise SystemExit(f"pdf.js patch: expected 1 match for {pattern!r}, found {n}")
    return new


def patch_pdfjs_worker(src: str) -> str:
    """Make pdf.js keep the PDF's own FontDescriptor /Descent (0 if absent) as font.pdfDescent."""
    # 1. when reading the font descriptor, store /Descent on the font properties
    src = _sub_once(r'(let (\w+)=\w+\.get\("Descent"\);"number"!=typeof \2&&\(\2=void 0\);.{0,400}?const \w+=\{)(type:)',
                    lambda m: f"{m.group(1)}pdfDescent:{m.group(2)}??0,{m.group(3)}", src)
    # 2. copy it onto the Font object
    src = _sub_once(r"(this\.descent=(\w+)\.descent/\w+;)", lambda m: f"{m.group(1)}this.pdfDescent={m.group(2)}.pdfDescent??0;", src)
    # 3. include it in the data sent from pdf.js's background thread to the page
    src = _sub_once(r'\["ascent","bbox",', '["pdfDescent","ascent","bbox",', src)
    # 4. make it a classic script: a double-clicked file may start classic background threads but not
    #    "module" ones. The worker has no imports and starts itself; only its final export line goes.
    src = _sub_once(r"export\{\w+ as WorkerMessageHandler\};\s*$", "", src)
    return src


def inline_script(code: str) -> str:
    if re.search(r"</script", code, re.I):
        raise SystemExit("a script contains '</script' and cannot be inlined")
    return f"<script>\n{code}\n</script>\n"


def js_string(text: str) -> str:
    """A JS string literal that is safe inside an inline <script>."""
    return json.dumps(text).replace("</", "<\\/").replace("<!--", "<\\!--")


def main() -> None:
    page = (HERE / "index.html").read_text(encoding="utf-8")
    replacements = [
        ("<footer>Python runs inside this browser tab (Pyodide). After the page has loaded, conversion works offline.</footer>",
         "<footer>Everything runs inside this one file: no internet needed, nothing is installed or uploaded.</footer>"),
        ('<div id="status"><span class="spin"></span>Starting…</div>',
         '<div id="status"><span class="spin"></span>Loading…</div>'),
    ]
    for old, new in replacements:
        if page.count(old) != 1:
            raise SystemExit(f"index.html changed: can't find {old[:60]!r}")
        page = page.replace(old, new)

    marker = "<script>\n  const COLUMNS"
    if page.count(marker) != 1:
        raise SystemExit("index.html changed: can't find the page script")
    libs = (inline_script(vendor("fflate.umd.js"))
            + inline_script((HERE / "js" / "std_fonts.js").read_text(encoding="utf-8"))
            + inline_script((HERE / "js" / "coa_core.js").read_text(encoding="utf-8"))
            + "<script>\nwindow.COA_PDFJS_SRC = " + js_string(vendor("pdf.min.mjs")) + ";\n"
            + "window.COA_PDFJS_WORKER_SRC = " + js_string(patch_pdfjs_worker(vendor("pdf.worker.min.mjs"))) + ";\n</script>\n"
            + inline_script((HERE / "js" / "app_local.js").read_text(encoding="utf-8")))
    page = page.replace(marker, libs + marker)
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(page, encoding="utf-8")
    print(f"wrote {OUT}  ({OUT.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
