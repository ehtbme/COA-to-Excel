/* Runs the converter in this page with plain JavaScript (single-file build, nothing downloaded).
 *
 * Gives src/page.html a worker-style object (postMessage in; status / ready / fatal / error / done
 * messages out) while pdf.js does the PDF reading in a background thread. build_single.py puts the pdf.js sources in COA_PDFJS_SRC and
 * COA_PDFJS_WORKER_SRC; fflate (window.fflate) and coa_core.js (window.COA) are inlined before this.
 */
(function () {
  "use strict";
  const blobUrl = (src) => URL.createObjectURL(new Blob([src], { type: "text/javascript" }));

  async function loadPdfjs() {
    const pdfjs = await import(blobUrl(window.COA_PDFJS_SRC));
    // pdf.js reads PDFs in its own background thread (a classic one: a double-clicked file can't start
    // "module" threads), so the page stays responsive.
    pdfjs.GlobalWorkerOptions.workerPort = new Worker(blobUrl(window.COA_PDFJS_WORKER_SRC));
    return pdfjs;
  }

  window.COA_LOCAL_WORKER = function () {
    const w = { onmessage: null };
    const send = (m) => setTimeout(() => w.onmessage && w.onmessage({ data: m }), 0);
    const libs = loadPdfjs().then((pdfjs) => ({ pdfjs, fflate: window.fflate }));
    libs.then(() => send({ type: "ready" }), (e) => send({ type: "fatal", text: String((e && e.message) || e) }));

    w.postMessage = async (msg) => {
      try {
        const l = await libs;
        const t0 = performance.now();
        const files = msg.files.map((f) => ({ name: f.name, data: new Uint8Array(f.data) }));
        const r = await window.COA.convertAll(l, files, (name, i, n, k, m) =>
          send({ type: "status", text: `File ${i} of ${n}: reading COA ${k} of ${m}…` }));
        send({ type: "done", results: r.results, outputs: r.outputs, all: r.all,
          seconds: (performance.now() - t0) / 1000 });
      } catch (e) {
        send({ type: "error", text: String((e && e.message) || e) });
      }
    };
    return w;
  };
})();
