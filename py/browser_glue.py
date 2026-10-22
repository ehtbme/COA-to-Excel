"""Runs coa_to_excel.py inside the browser (Pyodide). Called from worker.js.

Each uploaded file -> one Excel sheet, same rules as the web and desktop apps.
Nothing here touches the network: files live in the browser's in-memory filesystem.
"""
import contextlib
import io
import json
import re
import runpy
import shutil
import sys
import zipfile
from pathlib import Path, PurePosixPath

SCRIPT = "/app/coa_to_excel.py"
MAX_PDFS_PER_ZIP = 300
MAX_ZIP_UNCOMPRESSED_MB = 500


def safe_basename(name):
    base = PurePosixPath(str(name).replace("\\", "/")).name
    base = re.sub(r"[^A-Za-z0-9._ ()-]", "_", base).strip(" .").lstrip("-.")
    return base[:150] or "upload"


def check_zip(path):
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
    except zipfile.BadZipFile:
        return "not a valid zip file"
    pdfs = [i for i in infos if i.filename.lower().endswith(".pdf")]
    if not pdfs:
        return "zip contains no PDF files"
    if len(pdfs) > MAX_PDFS_PER_ZIP:
        return f"zip contains {len(pdfs)} PDFs (limit {MAX_PDFS_PER_ZIP})"
    if sum(i.file_size for i in infos) > MAX_ZIP_UNCOMPRESSED_MB * 1024 * 1024:
        return f"zip expands to more than {MAX_ZIP_UNCOMPRESSED_MB} MB"
    return None


def run_script(script_input, out):
    """Run coa_to_excel.py's command line in-process; return (exit_code, stdout, stderr)."""
    stdout, stderr = io.StringIO(), io.StringIO()
    argv = sys.argv
    sys.argv = [SCRIPT, "-o", str(out), "--", str(script_input)]
    code = 0
    try:
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
            runpy.run_path(SCRIPT, run_name="__main__")
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else 1
    except Exception as e:
        code = 1
        stderr.write(f"{type(e).__name__}: {e}\n")
    finally:
        sys.argv = argv
    return code, stdout.getvalue(), stderr.getvalue()


def convert(name, src, job):
    res = {"name": name, "ok": False, "reason": "", "coas": 0, "xlsx_name": "", "note": ""}
    base = safe_basename(name)
    ext = Path(base).suffix.lower()
    if ext not in (".pdf", ".zip"):
        res["reason"] = f"unsupported file type '{ext or 'none'}' (use .pdf or .zip)"
        return res, None
    job.mkdir(parents=True)
    upload = job / base
    shutil.copyfile(src, upload)
    if ext == ".pdf":
        if not upload.read_bytes()[:4] == b"%PDF":
            res["reason"] = "not a valid PDF file"
            return res, None
        script_input = job / "input.zip"              # coa_to_excel.py only accepts zips
        with zipfile.ZipFile(script_input, "w") as zf:
            zf.write(upload, arcname=base)
    else:
        problem = check_zip(upload)
        if problem:
            res["reason"] = problem
            return res, None
        script_input = upload

    out = job / "out"
    code, stdout, stderr = run_script(script_input, out)
    if code != 0:
        res["reason"] = "conversion failed: " + ((stderr.strip().splitlines() or ["no details"])[-1])[:300]
        return res, None
    xlsx = out / "COA_inventory.xlsx"
    if not xlsx.is_file():
        res["reason"] = "no Excel file was produced"
        return res, None
    m = re.search(r"(\d+) COA\(s\) processed", stdout)
    coas = int(m.group(1)) if m else 0
    m = re.search(r"Unreadable PDFs \((\d+)\)", stderr)
    unreadable = int(m.group(1)) if m else 0
    if coas == 0:
        res["reason"] = "no PDFs found"
        return res, None
    if unreadable >= coas:
        res["reason"] = "none of the PDFs could be read (corrupt or scanned without text?)"
        return res, None
    res.update(ok=True, coas=coas, xlsx_name=f"{Path(base).stem}_COA_inventory.xlsx",
               note=f"{unreadable} of {coas} PDFs could not be read (rows marked NA)" if unreadable else "")
    return res, xlsx


def convert_all(names_json):
    """Inputs are /in/0, /in/1, ...; outputs are written to /out/<i>.xlsx and /out/all.zip."""
    names = json.loads(names_json)
    work, outdir = Path("/work"), Path("/out")
    for d in (work, outdir):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir()
    results, used = [], set()
    for i, name in enumerate(names):
        res, xlsx = convert(name, Path(f"/in/{i}"), work / str(i))
        if res["ok"]:
            stem, n = res["xlsx_name"][:-5], 2
            while res["xlsx_name"] in used:
                res["xlsx_name"], n = f"{stem}_{n}.xlsx", n + 1
            used.add(res["xlsx_name"])
            shutil.copyfile(xlsx, outdir / f"{i}.xlsx")
            res["file"] = f"/out/{i}.xlsx"
        results.append(res)
    ok = [r for r in results if r["ok"]]
    if len(ok) > 1:
        with zipfile.ZipFile(outdir / "all.zip", "w", zipfile.ZIP_DEFLATED) as zf:
            for r in ok:
                zf.write(r["file"], arcname=r["xlsx_name"])
    shutil.rmtree(work, ignore_errors=True)
    shutil.rmtree("/in", ignore_errors=True)
    return json.dumps(results)
