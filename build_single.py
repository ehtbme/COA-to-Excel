"""Build dist/COA_to_Excel.html: the whole browser version in ONE file that works when double-clicked.

    python3 build_single.py

Packs index.html + worker.js + py/*.py + the pure-Python libraries (from wheels/) into one HTML file.
Only Pyodide (Python for the browser) is still loaded from cdn.jsdelivr.net when the page opens.
pdfminer's CJK character-map data is left out, the same as in the Windows desktop build: no COA uses it.
Run `bash build.sh` first if py/coa_to_excel.py or the wheels need refreshing.
"""
import base64
import io
import json
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "dist" / "COA_to_Excel.html"
SKIP = ("pdfminer/cmap/",)   # CJK character maps: ~4 MB, never used by the COAs


def libs_zip() -> bytes:
    """All wheels merged into one zip that is extracted straight into site-packages."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as out:
        for name in json.loads((HERE / "wheels" / "manifest.json").read_text()):
            with zipfile.ZipFile(HERE / "wheels" / name) as whl:
                for info in whl.infolist():
                    if info.is_dir() or info.filename.startswith(SKIP) or "__pycache__" in info.filename:
                        continue
                    out.writestr(info.filename, whl.read(info))
    return buf.getvalue()


def script_literal(text: str) -> str:
    """A JS string literal that is safe inside an inline <script>."""
    return json.dumps(text).replace("</", "<\\/").replace("<!--", "<\\!--")


def main() -> None:
    embedded = {
        "libs": base64.b64encode(libs_zip()).decode("ascii"),
        "py": {f: (HERE / "py" / f).read_text(encoding="utf-8") for f in ("coa_to_excel.py", "browser_glue.py")},
    }
    worker_src = "self.COA_EMBEDDED = " + json.dumps(embedded) + ";\n" + (HERE / "worker.js").read_text(encoding="utf-8")
    page = (HERE / "index.html").read_text(encoding="utf-8")
    marker = "<script>\n  const COLUMNS"
    if page.count(marker) != 1:
        raise SystemExit("index.html layout changed: can't find where to insert the embedded worker")
    inject = f"<script>window.COA_WORKER_SRC = {script_literal(worker_src)};</script>\n"
    page = page.replace(marker, inject + marker)
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(page, encoding="utf-8")
    print(f"wrote {OUT}  ({OUT.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
