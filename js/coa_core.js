/* COA -> Excel in plain JavaScript (no Python, nothing downloaded).
 *
 * A line-by-line port of coa_to_excel.py (copy in py/) plus the upload checks of the web and desktop apps,
 * so the single-file page gives the same results as the Python apps. Uses pdf.js (Mozilla's PDF reader)
 * for the PDFs and fflate for zip files. Checked against the Python script on all test COAs
 * (see tests/parity_test.py); if you change one, change the other and re-run that test.
 *
 * Python details that are copied on purpose: regex whitespace/digit classes, str.strip(),
 * str.isdigit(), pdfminer's character boxes and pdfplumber's dedupe_chars().
 */
(function (root) {
  "use strict";

  // --------------------------------------------------------------------------- constants
  const COLUMNS = [
    "Bottle no.", "Date Received", "Chemical", "Synonyms", "Lot No.", "Product No.",
    "Expiry Date", "CAS No.", "Manufacturer", "Container", "Cap Color", "Amount", "Form",
    "Concentration", "Location", "Chemical Formula", "Mol. Wt.",
    "Project", "Remarks",
  ];
  const NA = "NA";
  const COA_FIELDS = ["Chemical", "Synonyms", "Lot No.", "Product No.", "Expiry Date", "CAS No.",
    "Manufacturer", "Form", "Concentration", "Chemical Formula", "Mol. Wt."];
  const MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const KNOWN_MANUFACTURERS = [
    ["cayman", "Cayman Chemical"],
    ["lgc", "LGC Standards"],
    ["cerilliant", "Cerilliant"],
    ["sigma-aldrich", "Sigma-Aldrich"],
    ["lipomed", "Lipomed"],
    ["chiron", "Chiron"],
    ["toronto research", "Toronto Research Chemicals"],
    ["restek", "Restek"],
  ];
  const STD_FONTS = root.COA_STD_FONTS || {};   // js/std_fonts.js
  const MAX_PDFS_PER_ZIP = 300;
  const MAX_ZIP_UNCOMPRESSED_MB = 500;

  // --------------------------------------------------------------------------- Python-compatible text helpers
  // Python's \s / str.isspace() set (differs from JavaScript's \s).
  const WSC = "\\t\\n\\v\\f\\r\\x1c-\\x20\\x85\\xa0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000";
  const WS = `[${WSC}]`, NWS = `[^${WSC}]`, D = "\\p{Nd}";
  const W = "[\\p{L}\\p{N}_]";                                   // Python's \w
  const B = `(?:(?<=${W})(?!${W})|(?<!${W})(?=${W}))`;           // Python's \b
  const EOL = "(?=\\n|(?![^]))";                                  // Python's $ with re.M
  const rx = (src, flags = "") => new RegExp(src, flags + "u");
  const WS_RUN = rx(`${WS}+`, "g");
  const STRIP_RE = rx(`^${WS}+|${WS}+$`, "g");
  const pyStrip = (s) => s.replace(STRIP_RE, "");
  const pySplitWs = (s) => s.split(WS_RUN).filter((p) => p !== "");
  const DIGITS_RE = /^[\p{Nd}\u00b2\u00b3\u00b9\u2070\u2074-\u2079\u2080-\u2089\u2460-\u2468\u2474-\u247c\u2488-\u2490\u24ea\u24f5-\u24fd\u24ff\u2776-\u277e\u2780-\u2788\u278a-\u2792]+$/u;
  const isDigit = (s) => DIGITS_RE.test(s);                      // Python str.isdigit()
  const DASHES = /[\u00ad\u2010\u2011\u2012\u2013]/g;            // soft hyphen / unicode dashes -> "-"

  // re.search(pattern, text, re.I | re.M) -> group(1).strip() or null
  function search(re, text) {
    const m = re.exec(text);
    return m ? pyStrip(m[1]) : null;
  }

  // --------------------------------------------------------------------------- PDF characters (pdfminer/pdfplumber geometry)
  const IDENTITY = [1, 0, 0, 1, 0, 0];
  function mult(m1, m0) {                                         // pdfminer utils.mult_matrix
    const [a1, b1, c1, d1, e1, f1] = m1, [a0, b0, c0, d0, e0, f0] = m0;
    return [a0 * a1 + c0 * b1, b0 * a1 + d0 * b1, a0 * c1 + c0 * d1, b0 * c1 + d0 * d1,
      a0 * e1 + c0 * f1 + e0, b0 * e1 + d0 * f1 + f0];
  }
  const applyPt = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

  /** First page's characters with pdfplumber's attributes: text, x0, x1, top, bottom, size, fontname, upright. */
  async function pdfChars(pdfjs, bytes) {
    const task = pdfjs.getDocument({
      data: bytes.slice(), isEvalSupported: false, disableFontFace: true, useSystemFonts: false,
      stopAtErrors: false, verbosity: 0,
    });
    const doc = await task.promise;
    try {
      const page = await doc.getPage(1);
      const ops = await page.getOperatorList({ annotationMode: pdfjs.AnnotationMode.DISABLE });
      const OPS = pdfjs.OPS, objs = page.commonObjs;
      const fonts = new Map();
      for (let i = 0; i < ops.fnArray.length; i++) {
        if (ops.fnArray[i] !== OPS.setFont) continue;
        const name = ops.argsArray[i][0];
        if (!fonts.has(name)) fonts.set(name, objs.has(name) ? objs.get(name) : await new Promise((r) => objs.get(name, r)));
      }

      // pdfminer's page CTM and pdfplumber's page box
      const [vx0, vy0, vx1, vy1] = page.view, rot = page.rotate;
      let ctm;
      if (rot === 90) ctm = [0, -1, 1, 0, -vy0, vx1];
      else if (rot === 180) ctm = [-1, 0, 0, 1, vx1, vy1];
      else if (rot === 270) ctm = [0, 1, -1, 0, vy1, -vx0];
      else ctm = [1, 0, 0, 1, -vx0, -vy0];
      let [bx0, by0, bx1, by1] = [Math.min(vx0, vx1), Math.min(vy0, vy1), Math.max(vx0, vx1), Math.max(vy0, vy1)];
      if (rot === 90 || rot === 270) [bx0, by0, bx1, by1] = [by0, bx0, by1, bx1];
      const height = by1 - by0, width = bx1 - bx0, mbTop = height - by1, mbX0 = bx0;

      const fresh = () => ({ tm: IDENTITY, lx: 0, ly: 0, font: null, fontsize: 0, scaling: 100,
        charspace: 0, wordspace: 0, rise: 0, leading: 0 });
      let ts = fresh();
      const stack = [];
      const chars = [];

      const nextLine = () => {
        const [a, b, c, d, e, f] = ts.tm;
        ts.tm = [a, b, c, d, ts.leading * c + e, ts.leading * d + f];
        ts.lx = 0; ts.ly = 0;
      };
      const showText = (seq) => {
        const font = ts.font;
        if (!font) return;
        const matrix = mult(ts.tm, ctm);
        const [a, b, c, d, e, f] = matrix;
        const fontsize = ts.fontsize, scaling = ts.scaling * 0.01;
        const charspace = ts.charspace * scaling;
        const wordspace = font.composite ? 0 : ts.wordspace * scaling;
        const dxscale = 0.001 * fontsize * scaling;
        const fm0 = (font.fontMatrix || [0.001])[0];
        // pdfminer: the 14 standard fonts use built-in metrics; other fonts use the PDF's own /Descent
        const std = !font.composite && !font.isType3Font ? STD_FONTS[font.name] : undefined;
        let descentEm;
        if (std) descentEm = std[0] * 0.001;
        else if (font.isType3Font || font.pdfDescent === undefined) descentEm = font.descent || 0;
        else descentEm = -Math.abs(font.pdfDescent) * 0.001;
        const descent = descentEm * fontsize;
        const upright = 0 < a * d * scaling && b * c <= 0;
        const fontname = font.name || "unknown";
        let x = ts.lx;
        const y = ts.ly;
        let needcharspace = false;
        for (const g of seq) {
          if (typeof g === "number") { x -= g * dxscale; needcharspace = true; continue; }
          if (!g) continue;
          if (needcharspace) x += charspace;
          const w = std ? (std[2][g.unicode] ?? std[1]) * 0.001 : (g.width || 0) * fm0;
          const adv = w * fontsize * scaling;
          const m2 = [a, b, c, d, x * a + y * c + e, x * b + y * d + f];
          let [X0, Y0] = applyPt(m2, 0, descent + ts.rise);
          let [X1, Y1] = applyPt(m2, adv, descent + ts.rise + fontsize);
          if (X1 < X0) [X0, X1] = [X1, X0];
          if (Y1 < Y0) [Y0, Y1] = [Y1, Y0];
          chars.push({
            text: g.unicode || "", x0: X0 + (mbX0 !== 0 ? mbX0 : 0), x1: X1 + (mbX0 !== 0 ? mbX0 : 0),
            top: height - Y1 + mbTop, bottom: height - Y0 + mbTop, size: Y1 - Y0, fontname, upright,
          });
          x += adv;
          if (wordspace && g.isSpace) x += wordspace;
          needcharspace = true;
        }
        ts.lx = x;
      };

      for (let i = 0; i < ops.fnArray.length; i++) {
        const fn = ops.fnArray[i], args = ops.argsArray[i];
        switch (fn) {
          case OPS.save: stack.push({ ctm, ts: { ...ts } }); break;
          case OPS.restore: if (stack.length) ({ ctm, ts } = stack.pop()); break;
          case OPS.transform: ctm = mult(args, ctm); break;
          case OPS.paintFormXObjectBegin:
            stack.push({ ctm, ts: { ...ts } });
            if (args[0]) ctm = mult(args[0], ctm);
            ts = fresh();
            break;
          case OPS.paintFormXObjectEnd: if (stack.length) ({ ctm, ts } = stack.pop()); break;
          case OPS.beginText: ts.tm = IDENTITY; ts.lx = 0; ts.ly = 0; break;
          case OPS.setCharSpacing: ts.charspace = args[0]; break;
          case OPS.setWordSpacing: ts.wordspace = args[0]; break;
          case OPS.setHScale: ts.scaling = args[0]; break;
          case OPS.setLeading: ts.leading = -args[0]; break;
          case OPS.setFont: ts.font = fonts.get(args[0]) || null; ts.fontsize = args[1]; break;
          case OPS.setTextRise: ts.rise = args[0]; break;
          case OPS.moveText:
          case OPS.setLeadingMoveText: {
            const [tx, ty] = args, [a, b, c, d, e, f] = ts.tm;
            ts.tm = [a, b, c, d, tx * a + ty * c + e, tx * b + ty * d + f];
            if (fn === OPS.setLeadingMoveText) ts.leading = ty;
            ts.lx = 0; ts.ly = 0;
            break;
          }
          case OPS.setTextMatrix: ts.tm = args.slice(0, 6); ts.lx = 0; ts.ly = 0; break;
          case OPS.nextLine: nextLine(); break;
          case OPS.showText: case OPS.showSpacedText: showText(args[0]); break;
          case OPS.nextLineShowText: nextLine(); showText(args[0]); break;
          case OPS.nextLineSetSpacingShowText:
            ts.wordspace = args[0]; ts.charspace = args[1]; nextLine(); showText(args[2]); break;
          default: break;
        }
      }
      return { chars, width };
    } finally {
      await doc.destroy();
    }
  }

  // pdfplumber cluster_list / make_cluster_dict / cluster_objects
  function clusterObjects(xs, keyFn, tol) {
    const values = [...new Set(xs.map(keyFn))].sort((p, q) => p - q);
    const ids = new Map();
    let id = -1, last = null;
    for (const v of values) {
      if (last === null || !(v <= last + tol)) id++;
      ids.set(v, id);
      last = v;
    }
    const groups = new Map();
    for (const x of xs) {
      const k = ids.get(keyFn(x));
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(x);
    }
    return [...groups.keys()].sort((p, q) => p - q).map((k) => groups.get(k));
  }

  // pdfplumber Page.dedupe_chars(tolerance=1, extra_attrs=("fontname", "size"))
  function dedupeChars(chars, tol = 1) {
    const groups = new Map();
    for (const c of chars) {
      const k = JSON.stringify([c.upright, c.text, c.fontname, c.size]);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(c);
    }
    const keep = new Set();
    for (const grp of groups.values()) {
      for (const yc of clusterObjects(grp, (c) => c.top, tol)) {
        for (const xc of clusterObjects(yc, (c) => c.x0, tol)) {
          let best = xc[0];
          for (const c of xc) if (c.top < best.top || (c.top === best.top && c.x0 < best.x0)) best = c;
          keep.add(best);
        }
      }
    }
    return chars.filter((c) => keep.has(c));
  }

  /** coa_to_excel.page_lines: text lines with sub/superscripts kept on their parent line. */
  function pageLines(rawChars) {
    const chars = dedupeChars(rawChars).filter((c) => pyStrip(c.text) !== "").map((c) => ({ ...c }));
    chars.sort((p, q) => p.top - q.top || p.x0 - q.x0);

    // 1. group characters with (almost) the same top into lines
    const lines = [];
    for (const c of chars) {
      if (lines.length && Math.abs(c.top - lines[lines.length - 1][0].top) <= 2) lines[lines.length - 1].push(c);
      else lines.push([c]);
    }
    const box = (line) => {
      let top = Infinity, bottom = -Infinity, size = -Infinity;
      for (const c of line) { top = Math.min(top, c.top); bottom = Math.max(bottom, c.bottom); size = Math.max(size, c.size); }
      return [top, bottom, size];
    };

    // 2. move smaller digits onto an adjacent larger-font line they overlap with
    for (let i = 0; i < lines.length; i++) {
      const src = lines[i];
      for (const c of [...src].sort((p, q) => p.x0 - q.x0)) {
        if (!isDigit(c.text)) continue;
        for (let j = 0; j < lines.length; j++) {
          const host = lines[j];
          if (i === j || !host.length) continue;
          const [top, bottom, size] = box(host);
          if (size < c.size * 1.15 || !(c.top < bottom && c.bottom > top)) continue;
          let gap = Infinity;
          for (const h of host) gap = Math.min(gap, Math.max(0, Math.max(h.x0, c.x0) - Math.min(h.x1, c.x1)));
          if (gap <= 0.5 * size) {
            host.push(c);
            src.splice(src.indexOf(c), 1);
            break;
          }
        }
      }
    }

    // 3. build strings; a wide horizontal gap becomes 3 spaces (column separator)
    const out = [];
    for (const line of lines) {
      if (!line.length) continue;
      line.sort((p, q) => p.x0 - q.x0);
      let size = -Infinity;
      for (const c of line) size = Math.max(size, c.size);
      const main = line.filter((c) => c.size === size);
      let mid = 0;
      for (const c of main) mid += (c.top + c.bottom) / 2;
      mid /= main.length;
      for (const c of line) {
        c.sup = isDigit(c.text) && c.size < 0.87 * size && (c.top + c.bottom) / 2 < mid - 0.1 * size;
      }
      let text = "", prev = null;
      for (const c of line) {
        if (prev !== null) {
          const gap = c.x0 - prev.x1, big = Math.max(c.size, prev.size);
          if (gap > 2 * big) text += "   ";
          else if (gap > 0.15 * big) text += " ";
        }
        text += (c.sup ? "^" : "") + c.text;
        prev = c;
      }
      out.push({ text: text.replace(DASHES, "-"), x1: line[line.length - 1].x1 });
    }
    return out;
  }

  // --------------------------------------------------------------------------- parsing helpers
  const plain = (s) => (s ? s.split("^").join("") : s);
  const SEG_SPLIT = rx(`${WS}{3,}`);
  const firstSegment = (s) => pyStrip(pyStrip(s).split(SEG_SPLIT)[0]);

  const DATE_PATTERNS = [
    [rx(`(${D}{1,2})${WS}*[-/ .]?${WS}*([A-Za-z]{3,9})\\.?${WS}*[-/ ,.]?${WS}*(${D}{4})`), "dmy"],
    [rx(`([A-Za-z]{3,9})\\.?${WS}+(${D}{1,2}),?${WS}+(${D}{4})`), "mdy"],
    [rx(`(${D}{4})-(${D}{1,2})-(${D}{1,2})`), "ymd"],
    [rx(`(${D}{1,2})[/.](${D}{1,2})[/.](${D}{4})`), "dmy"],
    [rx(`([A-Za-z]{3,9})\\.?${WS}*[-/ ,]?${WS}*(${D}{4})`), "my"],
    [rx(`(${D}{4})-(${D}{1,2})${B}`), "ym"],
    [rx(`${B}(${D}{1,2})[/.](${D}{4})`), "my"],
  ];
  /** Date string -> DDMMMYYYY; day defaults to 01 when not given. */
  function formatDate(text) {
    if (!text) return null;
    const t = pyStrip(text);
    for (const [re, order] of DATE_PATTERNS) {
      const m = re.exec(t);
      if (!m) continue;
      const parts = {};
      [...order].forEach((k, i) => { parts[k] = m[i + 1]; });
      let month = parts.m;
      month = isDigit(month) ? parseInt(month, 10) : MONTH_NAMES.indexOf(month.slice(0, 3).toUpperCase()) + 1 || null;
      const day = parts.d !== undefined ? parseInt(parts.d, 10) : 1;
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return String(day).padStart(2, "0") + MONTH_NAMES[month - 1] + parts.y;
      }
    }
    return null;
  }

  const DEUTERIUM = rx(`\\^2${WS}*H`, "g");
  const BULLET = /[•·]/;
  const BULLET_SPLIT = rx(`${WS}*[•·]${WS}*`);
  function cleanFormula(s) {
    if (!s) return null;
    s = s.replace(DEUTERIUM, "D").split("^").join("");
    const parts = BULLET.test(s) ? s.split(BULLET_SPLIT) : pySplitWs(s);
    return parts.filter((p) => pyStrip(p) !== "").map((p) => p.replace(WS_RUN, "")).join(" • ");
  }

  const CONC_RE = rx(`([${D}.]+)${WS}*(µg|μg|ug|mg|ng|g)${WS}*/${WS}*m?l`, "i");
  const SOLVENT_RE = rx(`${B}in${WS}+([A-Za-z][A-Za-z0-9:,\\- ]*?)${WS}*(?:\\(|${EOL})`, "i");
  /** 'A 1 mg/ml (nominal, as free base) solution in methanol' -> '1 mg/ml in methanol (as free base)'. */
  function parseConcentration(s) {
    if (!s) return null;
    const m = CONC_RE.exec(s);
    if (!m) return null;
    const unit = m[2].toLowerCase().split("μ").join("µ").split("ug").join("µg");
    let conc = `${m[1]} ${unit}/ml`;
    const solvent = search(SOLVENT_RE, s.slice(m.index + m[0].length));
    if (solvent) conc += ` in ${solvent.toLowerCase()}`;
    if (s.toLowerCase().includes("free base")) conc += " (as free base)";
    return conc;
  }

  function parseForm(supplied) {
    const low = supplied.toLowerCase();
    for (const [word, form] of [["solution", "Solution"], ["solid", "Solid"], ["powder", "Solid"],
      ["oil", "Oil"], ["liquid", "Liquid"], ["gas", "Gas"]]) {
      if (low.includes(word)) return form;
    }
    return null;
  }

  function pyFloat(s) {                                           // Python float() on [\d.]+
    if (!/^(\d+\.?\d*|\.\d+)$/.test(s)) throw new Error(`ValueError: could not convert string to float: '${s}'`);
    return Number(s);
  }

  // --------------------------------------------------------------------------- COA parser
  const RE = {
    lot: rx(`(?:Batch|Lot)${WS}*(?:No\\.?|Number)${WS}*:?${WS}*([A-Za-z0-9-]+)`, "i"),
    product: rx(`(?:Item|Product|Catalog(?:ue)?|Cat\\.)${WS}*(?:No\\.?|Number|Code)${WS}*:?${WS}*([A-Za-z0-9-]+)`, "i"),
    cas: rx(`CAS${B}[^\\n]{0,30}?(${D}{2,7}-${D}{2}-${D})${B}`, "i"),
    formula: rx(`Molecular${WS}+formula\\.?${WS}*:?${WS}*([^\\n]+?)(?:${WS}*\\(salt\\)|${WS}*\\(free base\\)|${WS}{3,}|${EOL})`, "i"),
    mwSalt: rx(`Molecular${WS}+weight${WS}*:?${WS}*([${D}.]+)${WS}*\\(salt\\)`, "i"),
    mw: rx(`(?:Formula|Molecular)${WS}+weight${WS}*:?${WS}*([${D}.]+)`, "i"),
    expiry: rx(`Expir(?:y|ation)${WS}+date${WS}*:?${WS}*([^\\n]+?)(?:${WS}{3,}|\\(|${EOL})`, "i"),
    retest: rx(`(?:Re-?test|Use[- ]by)${WS}+date${WS}*:?${WS}*([^\\n]+?)(?:${WS}{3,}|\\(|${EOL})`, "i"),
    productName: rx(`Product${WS}+name${WS}*:?${WS}*([^\\n]+)${EOL}`, "i"),
    synonyms: rx(`Synonyms?${WS}*:?${WS}*([^\\n]+)${EOL}`, "i"),
    itemNo: rx(`(Item|Product)${WS}*No`, "i"),
    heading: rx(`(CERTIFICATE|CONFIRMATION)${WS}+of${WS}+ANALYSIS`, "i"),
    nameConc: rx(`,${WS}*[${D}.]+${WS}*${NWS}*g/m?l`, "i"),
    accredited: rx(`${WS}*ACCREDITED(?![^])`),
    manufacturer: rx(`(?:Manufacturer|Producer|Manufactured by)${WS}*:?${WS}*(${NWS}[^\\n]+?)(?:${WS}{3,}|${EOL})`, "i"),
    supplied: rx(`Supplied${WS}+as${WS}*:?${WS}*([^\\n]+)${EOL}`, "i"),
  };

  function parseCoa(lines, pageWidth, filename) {
    const texts = lines.map((l) => l.text);
    const text = texts.join("\n");
    if (!pyStrip(text)) throw new Error("ValueError: no text layer found (scanned PDF?)");

    const rec = {};
    rec["Lot No."] = search(RE.lot, text);
    rec["Product No."] = search(RE.product, text);
    rec["CAS No."] = search(RE.cas, text);
    rec["Chemical Formula"] = cleanFormula(search(RE.formula, text));
    const mw = search(RE.mwSalt, text) || search(RE.mw, text);
    rec["Mol. Wt."] = mw ? pyFloat(mw) : null;
    rec["Expiry Date"] = formatDate(search(RE.expiry, text) || search(RE.retest, text));

    // Chemical name + synonyms
    const productName = search(RE.productName, text);
    let synonyms = search(RE.synonyms, text);
    let itemIdx = texts.findIndex((t) => RE.itemNo.test(t));
    if (itemIdx < 0) itemIdx = null;
    const headingIdx = texts.findIndex((t) => RE.heading.test(t));
    let chemical;
    if (productName) {                         // e.g. LGC: "d,l-MDA-D5.HCl, 1mg/ml in Methanol (as free base)"
      const seg = firstSegment(productName);
      const m = RE.nameConc.exec(seg);
      chemical = m ? seg.slice(0, m.index) : seg;
    } else if (headingIdx >= 0 && headingIdx + 1 < texts.length) {   // Cayman ISO certificate
      chemical = firstSegment(texts[headingIdx + 1]).replace(RE.accredited, "");
    } else {                                   // Cayman one-page CofA: name on line 1, formal name until "Item No."
      chemical = firstSegment(texts[0]);
      if (!synonyms && itemIdx && itemIdx > 1) {
        synonyms = "";
        for (let k = 0; k < itemIdx - 1; k++) {
          const prev = lines[k], ln = lines[k + 1];
          const wrapped = k !== 0 && prev.x1 > 0.85 * pageWidth;
          synonyms += (wrapped || !synonyms ? "" : "; ") + firstSegment(ln.text);
        }
      }
    }
    rec["Chemical"] = plain(chemical);
    rec["Synonyms"] = plain(synonyms);

    // Manufacturer
    const haystack = (text + " " + filename).toLowerCase();
    const known = KNOWN_MANUFACTURERS.find(([key]) => haystack.includes(key));
    rec["Manufacturer"] = (known && known[1]) || search(RE.manufacturer, text);

    // Form + concentration
    const supplied = search(RE.supplied, text);
    if (supplied) {
      rec["Form"] = parseForm(supplied);
      rec["Concentration"] = parseConcentration(supplied);
    } else if (productName && parseConcentration(productName)) {
      rec["Form"] = "Solution";
      rec["Concentration"] = parseConcentration(productName);
    }
    for (const col of COA_FIELDS) {
      if (rec[col] === undefined || rec[col] === null || rec[col] === "") rec[col] = NA;
    }
    return rec;
  }

  // --------------------------------------------------------------------------- Excel (.xlsx) writer
  const xmlEsc = (s) => String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colLetter = (n) => { let s = ""; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const pyStr = (v) => (typeof v === "number" ? (Number.isInteger(v) ? v.toFixed(1) : String(v)) : String(v));

  /** Same sheet as coa_to_excel.write_excel: bold blue header, column widths, frozen header, filter. */
  function buildXlsx(fflate, rows) {
    const enc = new TextEncoder();
    const lastCol = colLetter(COLUMNS.length), lastRow = rows.length + 1;
    const cell = (ref, v, header) => {
      if (v === null || v === undefined || v === "") return "";
      if (typeof v === "number") return `<c r="${ref}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${header ? ' s="1"' : ""}><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
    };
    const widths = COLUMNS.map((col) => {
      let longest = [...col].length;
      for (const r of rows) longest = Math.max(longest, [...pyStr(r[col] ?? "")].length);
      return Math.min(Math.max(12, longest + 2), 60);
    });
    let sheetRows = `<row r="1">${COLUMNS.map((c, i) => cell(colLetter(i + 1) + "1", c, true)).join("")}</row>`;
    rows.forEach((r, ri) => {
      sheetRows += `<row r="${ri + 2}">${COLUMNS.map((c, i) => cell(colLetter(i + 1) + (ri + 2), r[c] ?? "")).join("")}</row>`;
    });
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:${lastCol}${lastRow}"/>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>
<sheetData>${sheetRows}</sheetData>
<autoFilter ref="A1:${lastCol}${lastRow}"/>
<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
</worksheet>`;
    const files = {
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Inventory" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Inventory!$A$1:$${lastCol}$${lastRow}</definedName></definedNames></workbook>`,
      "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="00D9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
      "xl/worksheets/sheet1.xml": sheet,
    };
    const zipInput = {};
    for (const [k, v] of Object.entries(files)) zipInput[k] = enc.encode(v);
    return fflate.zipSync(zipInput, { level: 6 });
  }

  // --------------------------------------------------------------------------- uploads (same checks as the web/desktop apps)
  function safeBasename(name) {                                    // web app safe_basename
    const parts = String(name).split("\\").join("/").split("/").filter((p) => p && p !== ".");
    let base = parts.length ? parts[parts.length - 1] : "";
    base = base.replace(/[^A-Za-z0-9._ ()-]/gu, "_").replace(/^[ .]+|[ .]+$/g, "").replace(/^[-.]+/, "");
    return base.slice(0, 150) || "upload";
  }
  function splitExt(name) {                                        // pathlib stem / suffix
    const i = name.lastIndexOf(".");
    return i > 0 && i < name.length - 1 ? [name.slice(0, i), name.slice(i)] : [name, ""];
  }
  const baseName = (p) => p.split("/").pop();
  const safeName = (s) => String(s).replace(/[^A-Za-z0-9.-]+/g, "-").replace(/^-+|-+$/g, "") || NA;

  /** Zip entries in central-directory order: [{name, originalSize}] (no data). Throws on a bad zip. */
  function zipEntries(fflate, bytes) {
    const infos = [];
    fflate.unzipSync(bytes, { filter: (f) => { infos.push({ name: f.name, originalSize: f.originalSize }); return false; } });
    return infos;
  }

  /** coa_to_excel.main for one zip: returns {rows, pdfs: [[newName, bytes]], unreadable}. */
  async function runScript(libs, members, progress) {
    const rows = [], pdfs = [], used = new Set();
    let unreadable = 0;
    for (let i = 0; i < members.length; i++) {
      const [orig, data] = members[i];
      if (progress) progress(i + 1, members.length);
      let rec;
      try {
        const { chars, width } = await pdfChars(libs.pdfjs, data);
        rec = parseCoa(pageLines(chars), width, orig);
      } catch (e) {
        unreadable++;
        rec = Object.fromEntries(COA_FIELDS.map((c) => [c, NA]));
      }
      const stem = `${safeName(rec["Lot No."])}_${safeName(rec["Expiry Date"])}`;
      let newName = `${stem}.pdf`;
      for (let n = 2; used.has(newName); n++) newName = `${stem}_${n}.pdf`;
      used.add(newName);
      pdfs.push([newName, data]);
      rows.push(rec);
    }
    return { rows, pdfs, unreadable };
  }

  /** One uploaded file (.pdf or .zip) -> rows + renamed PDFs, or a failure reason. */
  async function convertUpload(libs, name, bytes, progress) {
    const res = { name, ok: false, reason: "", coas: 0, xlsx_name: "", note: "" };
    const base = safeBasename(name);
    const [stem, suffix] = splitExt(base);
    const ext = suffix.toLowerCase();
    if (ext !== ".pdf" && ext !== ".zip") {
      res.reason = `unsupported file type '${ext || "none"}' (use .pdf or .zip)`;
      return { res };
    }
    let members;
    if (ext === ".pdf") {
      if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) {
        res.reason = "not a valid PDF file";
        return { res };
      }
      members = [[base, bytes]];
    } else {
      let infos;
      try { infos = zipEntries(libs.fflate, bytes); } catch (e) { res.reason = "not a valid zip file"; return { res }; }
      const nPdf = infos.filter((i) => i.name.toLowerCase().endsWith(".pdf")).length;
      if (!nPdf) { res.reason = "zip contains no PDF files"; return { res }; }
      if (nPdf > MAX_PDFS_PER_ZIP) { res.reason = `zip contains ${nPdf} PDFs (limit ${MAX_PDFS_PER_ZIP})`; return { res }; }
      if (infos.reduce((s, i) => s + i.originalSize, 0) > MAX_ZIP_UNCOMPRESSED_MB * 1024 * 1024) {
        res.reason = `zip expands to more than ${MAX_ZIP_UNCOMPRESSED_MB} MB`;
        return { res };
      }
      const wanted = (n) => !n.endsWith("/") && n.toLowerCase().endsWith(".pdf") && !n.includes("__MACOSX") && !baseName(n).startsWith("._");
      let files;
      try {
        files = libs.fflate.unzipSync(bytes, { filter: (f) => wanted(f.name) });
      } catch (e) {
        res.reason = "conversion failed: " + String(e.message || e).slice(0, 300);
        return { res };
      }
      members = infos.filter((i) => wanted(i.name) && files[i.name]).map((i) => [baseName(i.name), files[i.name]]);
    }
    const { rows, pdfs, unreadable } = await runScript(libs, members, progress);
    const coas = rows.length;
    if (coas === 0) { res.reason = "no PDFs found"; return { res }; }
    if (unreadable >= coas) { res.reason = "none of the PDFs could be read (corrupt or scanned without text?)"; return { res }; }
    Object.assign(res, {
      ok: true, coas, xlsx_name: `${stem}_COA_inventory.xlsx`,
      note: unreadable ? `${unreadable} of ${coas} PDFs could not be read (rows marked NA)` : "",
    });
    return { res, rows, pdfs };
  }

  /** All uploaded files: [{name, data}] -> {results, outputs: [{name, data}], all: zip bytes | null}. */
  async function convertAll(libs, files, progress) {
    const results = [], outputs = [], zipFiles = {}, used = new Set();
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      const { res, rows, pdfs } = await convertUpload(libs, f.name, f.data,
        progress && ((k, n) => progress(f.name, i + 1, files.length, k, n)));
      if (res.ok) {
        const stem = res.xlsx_name.slice(0, -5);
        for (let n = 2; used.has(res.xlsx_name); n++) res.xlsx_name = `${stem}_${n}.xlsx`;
        used.add(res.xlsx_name);
        const xlsx = buildXlsx(libs.fflate, rows);
        outputs.push({ name: res.xlsx_name, data: xlsx });
        res.pdf_dir = res.xlsx_name.slice(0, -"_COA_inventory.xlsx".length) + "_renamed_COAs";
        zipFiles[res.xlsx_name] = [xlsx, { level: 0 }];
        const sorted = [...pdfs].sort((p, q) => (p[0] < q[0] ? -1 : p[0] > q[0] ? 1 : 0));
        for (const [pname, data] of sorted) zipFiles[`${res.pdf_dir}/${pname}`] = [data, { level: 0 }];
        res.pdf_count = pdfs.length;
      }
      results.push(res);
    }
    const all = outputs.length ? libs.fflate.zipSync(zipFiles) : null;
    return { results, outputs, all };
  }

  root.COA = { COLUMNS, COA_FIELDS, pdfChars, dedupeChars, pageLines, parseCoa, formatDate, cleanFormula,
    parseConcentration, buildXlsx, safeBasename, convertUpload, convertAll };
})(typeof window !== "undefined" ? window : globalThis);
