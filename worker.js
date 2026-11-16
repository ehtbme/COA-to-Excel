// Runs Python (Pyodide) in a background thread so the page stays responsive.
// COA files are processed in this browser tab only; nothing is uploaded anywhere.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
importScripts(PYODIDE + "pyodide.js");

// The single-file build (build_single.py) sets self.COA_EMBEDDED = {libs: <base64 zip>, py: {name: source}}
// so nothing has to be fetched except Pyodide itself.
const EMB = self.COA_EMBEDDED || null;
const base = EMB ? "" : new URL(".", self.location.href).href;
let py;

async function init() {
  postMessage({ type: "status", text: "Loading Python in your browser (first visit downloads ~20 MB, then it's cached)…" });
  py = await loadPyodide({ indexURL: PYODIDE });
  if (EMB) {
    await py.loadPackage(["cryptography", "pillow", "charset-normalizer"]);
    py.FS.writeFile("/tmp/libs.zip", Uint8Array.from(atob(EMB.libs), (c) => c.charCodeAt(0)));
    py.runPython("import site, zipfile; zipfile.ZipFile('/tmp/libs.zip').extractall(site.getsitepackages()[0])");
  } else {
    await py.loadPackage(["micropip", "cryptography", "pillow", "charset-normalizer"]);
    const wheels = await (await fetch(base + "wheels/manifest.json", { cache: "no-cache" })).json();
    py.globals.set("wheel_urls", wheels.map((w) => base + "wheels/" + w));
    await py.runPythonAsync("import micropip\nawait micropip.install(list(wheel_urls), deps=False)");
  }
  py.FS.mkdirTree("/app");
  for (const f of ["coa_to_excel.py", "browser_glue.py"]) {
    py.FS.writeFile("/app/" + f, EMB ? EMB.py[f] : await (await fetch(base + "py/" + f, { cache: "no-cache" })).text());
  }
  py.runPython("import sys; sys.path.insert(0, '/app'); import browser_glue");
  postMessage({ type: "ready" });
}
const ready = init().catch((e) => postMessage({ type: "fatal", text: String(e) }));

onmessage = async (ev) => {
  await ready;
  const files = ev.data.files; // [{name, data: ArrayBuffer}]
  try {
    py.FS.mkdirTree("/in");
    files.forEach((f, i) => py.FS.writeFile("/in/" + i, new Uint8Array(f.data)));
    postMessage({ type: "status", text: `Processing ${files.length} file(s)…` });
    const t0 = performance.now();
    const results = JSON.parse(
      py.globals.get("browser_glue").convert_all(JSON.stringify(files.map((f) => f.name)))
    );
    const outputs = [];
    for (const r of results) {
      if (r.ok) outputs.push({ name: r.xlsx_name, data: py.FS.readFile(r.file) });
    }
    let all = null;
    if (outputs.length) all = py.FS.readFile("/out/all.zip");
    postMessage({ type: "done", results, outputs, all, seconds: (performance.now() - t0) / 1000 });
  } catch (e) {
    postMessage({ type: "error", text: String(e) });
  }
};
