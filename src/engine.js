/*
 * TplEngine — mesin deteksi & format untuk add-in "Template Paste".
 * Murni JavaScript (tanpa Office.js) supaya bisa diuji di browser biasa.
 *
 * Alur:  HTML/teks tempelan  ->  blocks  ->  classify()  ->  items
 *        items = [{type, level, runs:[{text,b,i,sup,sub,link}], rows}]
 *        type: 'body' | 'heading' | 'caption' | 'note' | 'ref' | 'table' | 'skip'
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TplEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ------------------------------------------------------------------ */
  /*  Profil default — diturunkan dari naskah copyedit AIEDS             */
  /* ------------------------------------------------------------------ */
  const DEFAULT_PROFILE = {
    name: "AIEDS (Artificial Intelligence in Educational Decision Sciences)",
    base: "aieds",
    font: "Cambria",
    color: "#000000",
    limits: {
      bodyWords: 0,           // batas kata isi artikel; 0 = belum diatur (isi sesuai jurnal)
      minBodyWords: 0,        // batas minimal (opsional)
      abstractWords: 0,
      excludeKeys: [],        // bagian yang tidak ikut dihitung, mis. ["contrib","ai"]
      countReferences: false, // referensi ikut dihitung?
      countTables: false,     // teks di dalam tabel ikut dihitung?
    },
    journal: { name: "Artificial Intelligence in Educational Decision Sciences", abbrev: "AIEDS" },
    check: { maxMonths: 3, minDays: 0, requireRevised: true },
    layout: {
      insertMode: "section",  // 'section' = di akhir bagian; 'cursor' = di posisi kursor
      blankBeforeSection: true, // sisakan 1 baris kosong sebelum judul bagian berikutnya
    },
    sectionHeading: { style: "Normal", size: 12, bold: true, italic: false, caps: true, align: "Centered", before: 0, after: 6, line: 1 },
    h2:      { style: "Normal", size: 11, bold: true,  italic: false, align: "Justified", before: 3, after: 3, line: 1 },
    h3:      { style: "Normal", size: 11, bold: true,  italic: false, align: "Justified", before: 3, after: 3, line: 1 },
    body:    { style: "Normal", size: 11, align: "Justified", firstIndentCm: 0, before: 3, after: 3, line: 1 },
    caption: { style: "Normal", size: 11, figureSize: 11, align: "Centered", labelBold: true, stripTitleItalic: true, before: 3, after: 3, line: 1 },
    note:    { style: "Normal", size: 11, align: "Justified", labelItalic: false, before: 3, after: 4, line: 1 },
    abstract: {
      size: 10, align: "Left", before: 0, after: 0, line: 1, italic: false,
      labelBold: true,          // label abstrak terstruktur (Purpose – …) dibuat bold
      sepBold: true,            // pemisah " – " ikut bold
      labelSep: " – ",          // pemisah label; kosong = pakai pemisah dari tempelan
      joinLabeled: "linebreak", // 'linebreak' = satu paragraf dipisah line break; 'paragraphs' = paragraf terpisah
      labels: ["Purpose", "Aim", "Aims", "Objective", "Objectives", "Background", "Design/methodology/approach", "Methodology", "Methods", "Method", "Design", "Approach",
        "Findings", "Results", "Research limitations/implications", "Research implications/limitations", "Research implications", "Research limitations", "Theoretical implications", "Managerial implications", "Policy implications", "Contribution", "Contributions", "Limitations", "Practical implications", "Social implications", "Implications",
        "Originality/value", "Originality", "Value", "Conclusion", "Conclusions", "Novelty", "Tujuan", "Metode", "Hasil", "Temuan", "Kesimpulan", "Keterbatasan", "Keterbatasan penelitian", "Implikasi", "Implikasi penelitian", "Orisinalitas", "Kebaruan"],
    },
    // Kutipan langsung / wawancara — APA 7 (Publication Manual 8.27 & 8.36):
    // >= 40 kata = blok menjorok 0,5 in (1,27 cm), tanpa tanda kutip, tidak italic, atribusi (…) setelah tanda baca akhir.
    quote:   { style: "Normal", size: 11, align: "Justified", leftCm: 1.27, rightCm: 0, italic: false, before: 3, after: 3, line: 1,
               apa: 1, minBlockWords: 40, removeMarks: true, attribution: "paren", mergeAttribution: true },
    // Rumus: disisipkan sebagai persamaan Word asli (OMML). numbering: 'tab' (nomor rata kanan) | 'after' | 'native' (Word 365 "#") | 'none'
    equation: { align: "Centered", numbering: "tab", textWidthCm: 16.5, size: 11, before: 6, after: 6, autoConvert: true },
    // autoFit: 'content' (AutoFit to contents) | 'window' | 'none'; repeatHeader: judul kolom diulang tiap halaman
    table:   { size: 10, headerBold: true, apaBorders: true, firstColLeft: true, cellAlign: "Centered", autoFit: "content", repeatHeader: true, align: "Centered" },
    references: {
      style: "Normal", size: 11, align: "Justified", hangingCm: 1.27, before: 0, after: 6, line: 1,
      sort: false, autoItalic: true, linkDoi: true, stripNumbering: true,
    },
    detect: {
      maxHeadingWords: 15,
      stripHeadingNumbers: true,
      titleCaseHeadings: false,
      keepItalic: true,
      keepBold: true,
      keepSupSub: true,
      mergeBrokenLines: true,
      requireCapsSectionTitle: true,
      detectQuotes: true,
    },
    sections: [
      { key: "abstract", title: "ABSTRACT", aliases: ["Abstract", "Abstrak"], kind: "abstract" },
      { key: "intro",   title: "INTRODUCTION", aliases: ["Introduction", "Pendahuluan", "Background"] },
      { key: "method",  title: "METHOD", aliases: ["Methods", "Methodology", "Research Method", "Research Methods", "Materials and Methods", "Metode", "Metode Penelitian", "Metodologi"] },
      { key: "rd",      title: "RESULTS AND DISCUSSION", aliases: ["Results and Discussions", "Result and Discussion", "Results & Discussion", "Hasil dan Pembahasan"],
        majorSubheadings: ["Results", "Result", "Findings", "Discussion", "Discussions", "Hasil", "Hasil Penelitian", "Pembahasan"] },
      { key: "concl",   title: "CONCLUSION", aliases: ["Conclusions", "Kesimpulan", "Simpulan", "Conclusion and Recommendations"] },
      { key: "contrib", title: "AUTHOR CONTRIBUTION STATEMENT", aliases: ["Author Contributions", "Author Contribution", "Kontribusi Penulis"] },
      { key: "ai",      title: "AI DISCLOSURE STATEMENT", aliases: ["AI Disclosure", "Declaration of AI Use"] },
      { key: "refs",    title: "REFERENCES", aliases: ["Reference", "Bibliography", "Daftar Pustaka", "Referensi", "Works Cited"], kind: "references" },
    ],
  };

  const clone = (o) => JSON.parse(JSON.stringify(o));
  // TplMath (math.js) — opsional; dimuat sebelum engine.js di task pane
  let _tm;
  const TM = () => {
    if (_tm) return _tm;
    _tm = (typeof TplMath !== "undefined" && TplMath) || (typeof self !== "undefined" && self.TplMath) || null;
    if (!_tm && typeof require === "function") { try { _tm = require("./math.js"); } catch (e) { _tm = null; } }
    return _tm;
  };

  /* Profil bawaan JEPSI — diturunkan dari naskah copyedit Banik & Roy (2026) */
  const JEPSD_PROFILE = (function () {
    const p = clone(DEFAULT_PROFILE);
    p.journal = { name: "Journal of Education Psychology and Social Development", abbrev: "JEPSI" };
    p.name = "JEPSI (Journal of Education Psychology and Social Development)";
    p.base = "jepsi";
    p.equation = Object.assign({}, p.equation, { textWidthCm: 15.92 });
    const sp = { before: 0, after: 6, line: 1 };
    Object.assign(p.sectionHeading, { size: 11 }, sp);
    Object.assign(p.h2, sp); Object.assign(p.h3, sp);
    Object.assign(p.body, sp);
    Object.assign(p.caption, sp);
    Object.assign(p.note, sp, { labelItalic: true });
    Object.assign(p.quote, sp, { leftCm: 1.27, italic: false }); // naskah copyedit JEPSI memakai italic; APA 7: tidak italic
    Object.assign(p.table, { cellAlign: "Left" });
    Object.assign(p.references, sp, { hangingCm: 1.27 });
    Object.assign(p.abstract, { size: 9, align: "Justified", before: 0, after: 8, line: 1.08, labelSep: "", joinLabeled: "paragraphs" });
    p.sections.find((s) => s.key === "intro").aliases.push("Introduction and Literature Review");
    return p;
  })();

  /* Profil bawaan VOICE — diturunkan dari naskah copyedit Kaharu, Umar & Waluyo (2026) */
  const VOICE_PROFILE = (function () {
    const p = clone(DEFAULT_PROFILE);
    p.name = "VOICE (Journal of Vocational, Informatics and Computer Education)";
    p.base = "voice";
    p.equation = Object.assign({}, p.equation, { numbering: "after", textWidthCm: 16.5 }); // contoh template: rumus di tengah, "(1)" langsung setelahnya
    p.journal = { name: "Journal of Vocational, Informatics and Computer Education", abbrev: "VOICE" };
    const sp = { before: 6, after: 6, line: 1 };
    Object.assign(p.sectionHeading, { size: 12 }, sp);
    Object.assign(p.h2, sp); Object.assign(p.h3, sp);
    Object.assign(p.body, sp);
    Object.assign(p.caption, sp, { size: 10, figureSize: 11 });   // caption tabel 10 pt, gambar 11 pt
    Object.assign(p.note, sp, { size: 10 });
    Object.assign(p.quote, sp);
    Object.assign(p.table, { size: 10, cellAlign: "Left" });
    Object.assign(p.references, { before: 0, after: 4, line: 1, hangingCm: 0.83, autoItalic: false }); // daftar pustaka VOICE tanpa italic
    Object.assign(p.abstract, { size: 10, align: "Justified", before: 0, after: 0, line: 1, labelSep: " – ", sepBold: false, joinLabeled: "paragraphs" });
    const concl = p.sections.findIndex((x) => x.key === "concl");
    p.sections.splice(concl + 1, 0, { key: "ack", title: "ACKNOWLEDGMENT", aliases: ["Acknowledgment", "Acknowledgement", "Acknowledgements", "ACKNOWLEDGEMENT", "Ucapan Terima Kasih"] });
    return p;
  })();

  /* Profil bawaan IJOBER — diturunkan dari naskah copyedit Kurniawan & Leniwati (2026) */
  const IJOBER_PROFILE = (function () {
    const p = clone(DEFAULT_PROFILE);
    p.name = "IJOBER";
    p.base = "ijober";
    p.equation = Object.assign({}, p.equation, { textWidthCm: 16.0 });
    p.journal = { name: "", abbrev: "IJOBER" };
    p.check = Object.assign({}, p.check, { requireCite: false });
    const sp = { before: 3, after: 3, line: 1 };
    // judul bagian bernomor ("1. INTRODUCTION"), gaya Heading 1, rata kiri
    Object.assign(p.sectionHeading, { style: "Heading 1", size: 11, bold: true, caps: true, align: "Left", before: 24, after: 6, line: 1 });
    Object.assign(p.h2, sp); Object.assign(p.h3, sp);
    Object.assign(p.body, sp);
    Object.assign(p.caption, { size: 10, figureSize: 10, before: 0, after: 0, line: 1, style: "Normal" }); // gaya Caption di template ini italic -> pakai Normal
    Object.assign(p.note, { size: 10, before: 0, after: 3 });
    Object.assign(p.quote, sp);
    Object.assign(p.table, { size: 10, cellAlign: "Left" });
    Object.assign(p.references, { before: 0, after: 3, line: 1, hangingCm: 0.85, autoItalic: false });
    Object.assign(p.abstract, { size: 10, align: "Justified", before: 0, after: 0, line: 1, labelSep: "", joinLabeled: "paragraphs" });
    p.detect.stripHeadingNumbers = false; // sub-heading bernomor (2.1, 3.5) dipertahankan
    p.sections = [
      { key: "abstract", title: "ABSTRACT", aliases: ["Abstract", "Abstrak"], kind: "abstract" },
      { key: "intro", title: "1. INTRODUCTION", aliases: ["Introduction", "Pendahuluan"] },
      { key: "method", title: "2. METHOD", aliases: ["Methods", "Methodology", "Research Method", "Research Methods", "Metode", "Metode Penelitian"] },
      { key: "rd", title: "3. RESULTS AND DISCUSSION", aliases: ["Results and Discussions", "Result and Discussion", "Hasil dan Pembahasan"], majorSubheadings: [] },
      { key: "concl", title: "4. CONCLUSIONS AND SUGGESTION", aliases: ["Conclusion", "Conclusions", "Conclusion and Suggestion", "Conclusions and Suggestions", "Kesimpulan", "Kesimpulan dan Saran"] },
      { key: "refs", title: "REFERENCES", aliases: ["Reference", "Bibliography", "Daftar Pustaka", "Referensi"], kind: "references" },
    ];
    return p;
  })();

  /* Profil bawaan COMIT — diturunkan dari naskah copyedit Wu et al. (2026), Vol. 1 No. 1 */
  const COMIT_PROFILE = (function () {
    const p = clone(DEFAULT_PROFILE);
    p.name = "COMIT (Computational Media and Immersive Technologies)";
    p.base = "comit";
    p.journal = { name: "Computational Media and Immersive Technologies", abbrev: "COMIT" };
    p.equation = Object.assign({}, p.equation, { textWidthCm: 15.92 }); // A4, margin 2,54 cm
    const sp = { before: 0, after: 6, line: 1 };
    Object.assign(p.sectionHeading, { size: 12, bold: true, caps: true, align: "Centered", before: 0, after: 8, line: 1.08 });
    Object.assign(p.h2, sp); Object.assign(p.h3, sp);
    Object.assign(p.body, sp);
    Object.assign(p.caption, sp, { size: 11, figureSize: 11, align: "Centered" });
    Object.assign(p.note, { size: 10, align: "Justified", before: 0, after: 8, line: 1.08, italic: true, labelItalic: true, labelBold: true }); // "Note." bold-italic, isi italic
    Object.assign(p.quote, sp);
    Object.assign(p.table, { size: 10, headerBold: true, apaBorders: true, cellAlign: "Left" });
    // daftar pustaka gaya IEEE bernomor [1]: nomor dipertahankan, nama jurnal/judul buku italic, "doi: …" apa adanya
    Object.assign(p.references, { size: 11, align: "Justified", hangingCm: 0.88, before: 0, after: 6, line: 1,
      citeStyle: "ieee", stripNumbering: false, sort: false, autoItalic: true, linkDoi: false, doiFormat: "keep" });
    Object.assign(p.abstract, { size: 9, align: "Justified", before: 0, after: 8, line: 1.08, labelSep: "", joinLabeled: "paragraphs" });
    const back = { size: 11, before: 0, after: 8, line: 1.08 }; // judul bagian penutup 11 pt
    p.sections = [
      { key: "abstract", title: "ABSTRACT", aliases: ["Abstract", "Abstrak"], kind: "abstract" },
      { key: "intro", title: "INTRODUCTION", aliases: ["Introduction", "Pendahuluan"] },
      { key: "method", title: "METHOD", aliases: ["Methods", "Methodology", "Research Method", "Research Methods", "Materials and Methods", "Metode", "Metode Penelitian"] },
      { key: "rd", title: "RESULTS AND DISCUSSION", aliases: ["Results and Discussions", "Result and Discussion", "Results & Discussion", "Hasil dan Pembahasan"],
        majorSubheadings: ["Results", "Result", "Findings", "Discussion", "Discussions", "Hasil", "Pembahasan"] },
      { key: "concl", title: "CONCLUSIONS", aliases: ["Conclusion", "Conclusions and Recommendations", "Kesimpulan", "Simpulan"] },
      { key: "ack", title: "ACKNOWLEDGMENT", aliases: ["AKNOWLEDGMENT", "Acknowledgement", "Acknowledgements", "Acknowledgments", "Ucapan Terima Kasih"], heading: { size: 11, before: 12, after: 6, line: 1 } },
      { key: "contrib", title: "AUTHOR CONTRIBUTION STATEMENT", aliases: ["Author Contributions", "Author Contribution", "Kontribusi Penulis"], heading: back },
      { key: "ai", title: "AI DISCLOSURE STATEMENT", aliases: ["AI Disclosure", "Declaration of AI Use"], heading: back },
      { key: "refs", title: "REFERENCES", aliases: ["Reference", "Bibliography", "Daftar Pustaka", "Referensi"], kind: "references", heading: back },
    ];
    return p;
  })();

  /** Template bawaan (urutan = urutan di menu). */
  const PRESETS = [DEFAULT_PROFILE, JEPSD_PROFILE, VOICE_PROFILE, IJOBER_PROFILE, COMIT_PROFILE];
  const presetFor = (base) => PRESETS.find((x) => x.base === base) || DEFAULT_PROFILE;

  /** Gabungkan profil tersimpan dengan default (agar kunci baru tetap ada). */
  function mergeProfile(p) {
    const out = clone(presetFor(p && p.base));
    if (!p) return out;
    for (const k of Object.keys(p)) {
      if (p[k] && typeof p[k] === "object" && !Array.isArray(p[k]) && out[k] && typeof out[k] === "object" && !Array.isArray(out[k])) {
        out[k] = Object.assign(out[k], p[k]);
      } else out[k] = p[k];
    }
    return out;
  }

  /* ------------------------------------------------------------------ */
  /*  Utilitas teks                                                      */
  /* ------------------------------------------------------------------ */
  const WS = /[\s  -​  　]+/g;
  const countWords = (t) => (t || "").replace(WS, " ").trim().split(" ").filter((w) => /[\p{L}\p{N}]/u.test(w)).length;
  const plain = (runs) => runs.map((r) => r.text).join("");
  const sameFmt = (a, b) => !!a.b === !!b.b && !!a.i === !!b.i && !!a.sup === !!b.sup && !!a.sub === !!b.sub && (a.link || "") === (b.link || "");

  /** Rapikan runs: spasi ganda -> satu, trim ujung, gabung run berformat sama. */
  function normalizeRuns(runs) {
    const out = [];
    for (const r0 of runs) {
      if (!r0 || r0.text == null) continue;
      const r = Object.assign({}, r0, { text: String(r0.text).replace(/[\r\n\t\v\f]+/g, " ").replace(WS, " ") });
      if (!r.text) continue;
      const prev = out[out.length - 1];
      if (prev && sameFmt(prev, r)) prev.text += r.text;
      else out.push(r);
    }
    // collapse spasi lintas run
    for (let k = 1; k < out.length; k++) {
      if (out[k - 1].text.endsWith(" ") && out[k].text.startsWith(" ")) out[k].text = out[k].text.slice(1);
    }
    if (out.length) out[0].text = out[0].text.replace(/^ +/, "");
    if (out.length) out[out.length - 1].text = out[out.length - 1].text.replace(/ +$/, "");
    return out.filter((r) => r.text.length);
  }

  /** Buang n karakter pertama dari runs. */
  function dropPrefix(runs, n) {
    const out = [];
    for (const r of runs) {
      if (n <= 0) { out.push(Object.assign({}, r)); continue; }
      if (r.text.length <= n) { n -= r.text.length; continue; }
      out.push(Object.assign({}, r, { text: r.text.slice(n) }));
      n = 0;
    }
    return normalizeRuns(out);
  }

  /** Terapkan fungsi fmt(r) pada rentang karakter [s,e) dari runs (split bila perlu). */
  function patchRange(runs, s, e, patch) {
    const out = [];
    let pos = 0;
    for (const r of runs) {
      const a = pos, b = pos + r.text.length;
      pos = b;
      if (b <= s || a >= e) { out.push(r); continue; }
      const cut1 = Math.max(s, a) - a, cut2 = Math.min(e, b) - a;
      if (cut1 > 0) out.push(Object.assign({}, r, { text: r.text.slice(0, cut1) }));
      out.push(Object.assign({}, r, patch, { text: r.text.slice(cut1, cut2) }));
      if (cut2 < r.text.length) out.push(Object.assign({}, r, { text: r.text.slice(cut2) }));
    }
    return out;
  }

  function mapRunsText(runs, fn) {
    return runs.map((r) => Object.assign({}, r, { text: fn(r.text) }));
  }

  const MINOR = new Set("a an and as at but by for from in into nor of on or per the to vs via with yang dan di ke dari atau untuk pada dalam dengan".split(" "));
  function titleCase(s) {
    let first = true;
    return s.replace(/[\p{L}][\p{L}'’\-]*/gu, (w, off) => {
      const lower = w.toLowerCase();
      const prevChar = s[off - 2] || "";
      const forceCap = first || /[:—–-]/.test(prevChar);
      first = false;
      if (/[A-Z].*[A-Z]/.test(w)) return w; // akronim (FCM, AI, SEM-PLS)
      if (!forceCap && MINOR.has(lower)) return lower;
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    });
  }

  /* ------------------------------------------------------------------ */
  /*  Parser HTML tempelan (Word, Google Docs, ChatGPT, web)             */
  /* ------------------------------------------------------------------ */
  const BLOCK = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "PRE", "SECTION", "ARTICLE", "HEADER", "FOOTER", "DD", "DT", "FIGCAPTION", "CAPTION", "ADDRESS", "ASIDE", "MAIN", "NAV", "UL", "OL", "DL", "HR", "CENTER"]);
  const SKIP = new Set(["STYLE", "SCRIPT", "HEAD", "TITLE", "META", "LINK", "NOSCRIPT", "TEMPLATE", "SVG", "MATH", "BUTTON", "OBJECT"]);

  function styleProp(el, prop) {
    const st = el.getAttribute && el.getAttribute("style");
    if (!st) return null;
    const m = new RegExp("(?:^|;)\\s*" + prop + "\\s*:\\s*([^;]+)", "i").exec(st);
    return m ? m[1].trim().toLowerCase() : null;
  }

  function fmtFrom(el, parent) {
    const f = Object.assign({}, parent);
    const t = el.tagName;
    if (t === "B" || t === "STRONG") f.b = true;
    if (t === "I" || t === "EM" || t === "CITE" || t === "VAR" || t === "DFN") f.i = true;
    if (t === "SUP") f.sup = true;
    if (t === "SUB") f.sub = true;
    if (t === "A" && el.getAttribute("href") && /^https?:/i.test(el.getAttribute("href"))) f.href = el.getAttribute("href");
    const fw = styleProp(el, "font-weight");
    if (fw) f.b = /bold|bolder|[6-9]00/.test(fw) ? true : /normal|lighter|[1-4]00/.test(fw) ? false : f.b;
    const fs = styleProp(el, "font-style");
    if (fs) f.i = /italic|oblique/.test(fs) ? true : fs === "normal" ? false : f.i;
    const tt = styleProp(el, "text-transform");
    if (tt) f.upper = /uppercase/.test(tt) ? true : tt === "none" ? false : f.upper;
    if (/mso-style-name:\s*"?Heading/i.test(el.getAttribute("style") || "")) f.headingStyle = true;
    const va = styleProp(el, "vertical-align");
    if (va) { if (/super/.test(va)) f.sup = true; else if (/sub/.test(va)) f.sub = true; else if (va === "baseline") { f.sup = false; f.sub = false; } }
    return f;
  }

  /**
   * @param {string} html
   * @param {DOMParser} [ParserImpl]
   * @returns {{blocks:Array, images:number, tables:number}}
   */
  /* ---------------- rumus di HTML (Word, KaTeX/ChatGPT, MathJax, Wikipedia, MathML) ---------------- */
  const MATH_SEL = "[data-tpl-omml], math, .katex, mjx-container, .mwe-math-element";
  function mathFromEl(el) {
    const T = TM();
    if (!T || !el.getAttribute) return null;
    const tag = el.tagName.toUpperCase();
    const cls = el.getAttribute("class") || "";
    const om = el.getAttribute("data-tpl-omml");
    if (om != null) {
      const r = T.ommlFromWordHtml(decodeURIComponent(om));
      if (!r.ok) return { skip: true };
      const ast = T.ommlXmlToAst(r.xml);
      if (!ast) return { skip: true }; // OMML rusak
      return { omml: r.xml, ast, display: r.display, text: T.toUnicode(ast).trim(), source: "word" };
    }
    let mathEl = null, display = false, tex = null;
    if (/\bkatex-display\b/.test(cls)) { display = true; mathEl = el.querySelector("math"); }
    else if (/(^|\s)katex(\s|$)/.test(cls)) { mathEl = el.querySelector("math"); display = !!(el.closest && el.closest(".katex-display")); }
    else if (tag === "MJX-CONTAINER") { mathEl = el.querySelector("math"); display = el.getAttribute("display") === "true"; }
    else if (/\bmwe-math-element\b/.test(cls)) {
      mathEl = el.querySelector("math");
      display = !!el.querySelector('[class*="display"]');
      if (!mathEl) { const img = el.querySelector("img[alt]"); if (img) tex = img.getAttribute("alt"); }
    } else if (tag === "MATH") { mathEl = el; display = el.getAttribute("display") === "block"; }
    else if (/\bkatex-html\b|\bMathJax_Preview\b/.test(cls)) return { skip: true };
    else return null;
    let ast = null, number = null;
    if (mathEl) {
      const alt = mathEl.getAttribute("alttext");
      const fm = T.fromMathML(mathEl);
      if (fm.tex) tex = fm.tex; else if (alt) tex = alt; else ast = fm.ast;
    }
    if (tex != null) { const r = T.parse(tex); if (r.error) return { skip: true }; ast = r.ast; number = r.number; }
    if (!ast) return { skip: true };
    return { ast, tex: tex || T.toLatex(ast), number, display, text: T.toUnicode(ast).trim(), source: "web" };
  }
  const RE_EQNUM = /^\s*\(\s*(\d+[a-z]?)\s*\)\s*$/;
  /** "(1)" di paragraf tersendiri tepat setelah rumus -> nomor rumus tsb. */
  function attachEquationNumbers(blocks) {
    const out = [];
    for (const b of blocks) {
      const prev = out[out.length - 1];
      const t = b.kind === "p" ? plain(b.runs) : "";
      if (prev && prev.math && !prev.math.number && RE_EQNUM.test(t)) { prev.math = Object.assign({}, prev.math, { number: RE_EQNUM.exec(t)[1] }); continue; }
      out.push(b);
    }
    return out;
  }

  function parseHtml(html, ParserImpl, opts) {
    opts = opts || {};
    const P = ParserImpl || (typeof DOMParser !== "undefined" ? DOMParser : null);
    // Buang komentar kondisional Word & fragment marker
    html = String(html)
      // persamaan Word: OMML ada di komentar kondisional; gambar pengganti dibuang
      .replace(/<!--\[if gte msEquation \d+\]>([\s\S]*?)<!\[endif\]-->/gi, (m, x) => `<span data-tpl-omml="${encodeURIComponent(x)}"></span>`)
      .replace(/<!\[if !msEquation\]>[\s\S]*?<!\[endif\]>/gi, "")
      .replace(/<!--\[if !msEquation\]-->[\s\S]*?<!--\[endif\]-->/gi, "")
      .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "").replace(/<!--[\s\S]*?-->/g, "");
    const doc = new P().parseFromString(html, "text/html");
    const blocks = [];
    let cur = null;
    let images = 0, tables = 0;
    const olStack = [];

    const flush = () => {
      if (cur && cur.runs.length) {
        const runs = normalizeRuns(cur.runs);
        // paragraf yang isinya hanya satu rumus (+ nomor "(1)") -> rumus baris tersendiri
        const mp = cur.maths || [];
        const rest = cur.runs.filter((r) => !r.math).map((r) => r.text).join("");
        const numM = /^\s*\(\s*(\d+[a-z]?)\s*\)\s*$/.exec(rest);
        if (mp.length === 1 && (!rest.trim() || numM)) {
          const m = Object.assign({}, mp[0], { display: true });
          if (numM && !m.number) m.number = numM[1];
          blocks.push({ kind: "p", tag: "p", cls: "", runs: [{ text: m.text || "[rumus]", b: false, i: false }], html: true, math: m });
        } else if (runs.length) blocks.push({ kind: "p", tag: cur.tag, cls: cur.cls, runs, html: true });
      }
      cur = null;
    };
    const pushMathBlock = (m) => {
      flush();
      blocks.push({ kind: "p", tag: "p", cls: "", runs: [{ text: m.text || "[rumus]", b: false, i: false }], html: true, math: m });
    };
    const ensure = (tag, cls) => { if (!cur) cur = { tag: tag || "p", cls: cls || "", runs: [] }; };

    function walk(node, fmt, blockTag, blockCls) {
      if (node.nodeType === 3) {
        const txt = node.nodeValue;
        if (!txt) return;
        if (!cur && !/\S/.test(txt)) return;
        ensure(blockTag, blockCls);
        cur.runs.push({ text: fmt.upper ? txt.toUpperCase() : txt, b: !!fmt.b, i: !!fmt.i, sup: !!fmt.sup, sub: !!fmt.sub, link: fmt.href || "" });
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node;
      const tag = el.tagName.toUpperCase();
      const mth = mathFromEl(el);
      if (mth) {
        if (mth.skip) return;
        if (mth.display) pushMathBlock(mth);
        else {
          ensure(blockTag, blockCls);
          cur.runs.push({ text: mth.text || "", b: false, i: false, sup: false, sub: false, link: "", math: true });
          (cur.maths = cur.maths || []).push(mth);
        }
        return;
      }
      if (SKIP.has(tag)) return;
      if (/mso-list\s*:\s*ignore/i.test(el.getAttribute("style") || "")) return; // nomor/bullet palsu dari Word
      // daftar catatan kaki/akhir & komentar dari clipboard Word
      if (/mso-element\s*:\s*(footnote|endnote|comment)(-list)?/i.test(el.getAttribute("style") || "")) return;
      if (/^_(ftn|edn|msocom)/i.test(el.getAttribute("id") || "") && tag === "DIV") return;
      if (tag === "IMG" || tag === "PICTURE" || tag === "CANVAS" || tag === "V:SHAPE" || tag === "V:IMAGEDATA") { images++; return; }
      if (tag === "BR") { flush(); return; }
      if (tag === "TABLE") {
        flush();
        // Tabel 1 sel (mis. teks yang disalin dari dalam kotak/sel) atau mode "ratakan" (abstrak):
        // isi sel diperlakukan sebagai paragraf biasa, bukan tabel data.
        const allCells = Array.from(el.querySelectorAll("td,th")).filter((td) => td.closest("table") === el);
        // tabel penomoran rumus (| rumus | (1) |) -> satu rumus bernomor
        const trs = Array.from(el.querySelectorAll("tr")).filter((tr) => tr.closest("table") === el);
        const mEl = trs.length === 1 ? el.querySelector(MATH_SEL) : null;
        if (mEl) {
          const m = mathFromEl(mEl);
          if (m && !m.skip) {
            const others = allCells.filter((td) => !td.contains(mEl)).map((td) => (td.textContent || "").replace(WS, " ").trim()).filter(Boolean);
            if (others.every((t) => /^\(\s*\d+[a-z]?\s*\)$/.test(t))) {
              const n = others.length ? /\d+[a-z]?/.exec(others[0])[0] : null;
              pushMathBlock(Object.assign({}, m, { display: true, number: m.number || n }));
              return;
            }
          }
        }
        // tabel halaman depan (ARTICLE INFO | ABSTRACT …) saat tempel keseluruhan: isinya dijadikan paragraf
        const isAbsHead = (x) => /^\s*(abstract|abstrak)\s*$/i.test((x.textContent || "").replace(WS, " "));
        const front = opts.flattenFront && allCells.some((td) => isAbsHead(td) || Array.from(td.querySelectorAll("p,div,h1,h2,h3,h4,b,strong")).some(isAbsHead));
        if (opts.flattenTables || allCells.length <= 1 || front) {
          for (const td of allCells) { for (const c of td.childNodes) walk(c, fmt, "p", ""); flush(); }
          return;
        }
        const rows = [];
        el.querySelectorAll("tr").forEach((tr) => {
          if (tr.closest("table") !== el) return; // abaikan tabel bersarang
          const cells = [];
          tr.querySelectorAll("td,th").forEach((td) => {
            if (td.closest("tr") !== tr) return;
            const t = (td.textContent || "").replace(WS, " ").trim();
            cells.push(t);
            const cs = parseInt(td.getAttribute("colspan") || "1", 10);
            for (let k = 1; k < cs && k < 30; k++) cells.push("");
          });
          if (cells.length) rows.push(cells);
        });
        if (rows.length) { tables++; blocks.push({ kind: "table", rows }); }
        return;
      }
      const f = fmtFrom(el, fmt);
      const isBlock = BLOCK.has(tag) || /^(block|list-item|table|flex|grid)/.test(styleProp(el, "display") || "");
      if (tag === "OL") olStack.push(0);
      if (isBlock) {
        flush();
        let bt = /^H[1-6]$/.test(tag) ? tag.toLowerCase() : tag === "LI" ? "li" : tag === "BLOCKQUOTE" ? "blockquote" : (blockTag === "blockquote" ? "blockquote" : "p");
        const cls = el.getAttribute("class") || "";
        const m = /MsoHeading(\d)|Heading(\d)/i.exec(cls);
        if (m) bt = "h" + (m[1] || m[2]);
        if (/MsoTitle/i.test(cls)) bt = "h1";
        const role = el.getAttribute("role"), aria = el.getAttribute("aria-level");
        if (role === "heading" && aria) bt = "h" + aria;
        let prefix = "";
        if (tag === "LI" && olStack.length && el.parentElement && el.parentElement.tagName === "OL") {
          olStack[olStack.length - 1]++;
          prefix = olStack[olStack.length - 1] + ". ";
        }
        if (prefix) { ensure(bt, cls); cur.runs.push({ text: prefix, b: false, i: false }); }
        for (const c of el.childNodes) walk(c, f, bt, cls);
        flush();
      } else {
        for (const c of el.childNodes) walk(c, f, blockTag, blockCls);
      }
      if (tag === "OL") olStack.pop();
    }

    const bodyEl = doc.body || doc.documentElement;
    for (const c of bodyEl.childNodes) walk(c, {}, "p", "");
    flush();
    return { blocks: attachEquationNumbers(blocks), images, tables, source: "html" };
  }

  /* ------------------------------------------------------------------ */
  /*  Parser teks biasa (Notepad, PDF, Markdown/ChatGPT)                 */
  /* ------------------------------------------------------------------ */
  function mdInline(text) {
    // **bold**, *italic*, _italic_ — hanya bila diapit karakter non-spasi
    const runs = [];
    // tanpa lookbehind (tidak didukung WebView lama di Word for Mac)
    const re = /(\*\*|__)(\S(?:[\s\S]*?\S)?)\1|(^|[^\w*])(\*|_)([^\s*_](?:[^*_\n]*?[^\s*_])?)\4(?![\w*])/g;
    let last = 0, m;
    while ((m = re.exec(text))) {
      const start = m.index + (m[1] ? 0 : m[3].length);
      if (start > last) runs.push({ text: text.slice(last, start) });
      if (m[1]) runs.push({ text: m[2], b: true });
      else runs.push({ text: m[5], i: true });
      last = re.lastIndex;
    }
    if (last < text.length) runs.push({ text: text.slice(last) });
    return runs;
  }

  const endsSentence = (s) => /[.!?:;)"”’\]]\s*$/.test(s);
  // baris lanjutan (huruf kecil di awal) — kecuali penanda daftar "a. …", "ii) …"
  const startsLower = (s) => /^[\p{Ll}(\[,;–-]/u.test(s.trim()) && !/^(?:[a-z]|[ivx]{1,4})[.)]\s/.test(s.trim());

  function parseText(text, opts) {
    opts = opts || {};
    const merge = opts.mergeBrokenLines !== false;
    text = String(text).replace(/\r\n?/g, "\n").replace(/\u2028|\u2029/g, "\n");
    // rumus LaTeX (ChatGPT/Markdown): $$…$$, \[…\], \begin{equation}…  -> rumus baris; \(…\) dan $…$ -> teks Unicode
    const maths = [];
    const T = TM();
    if (T && /\$|\\[[(]|\\begin\{/.test(text)) {
      const hold = (tex) => { maths.push(tex); return "\n\u0001M" + (maths.length - 1) + "\u0001\n"; };
      text = text.replace(/\$\$([\s\S]+?)\$\$|\\\[([\s\S]+?)\\\]/g, (m, a, b) => hold(a != null ? a : b));
      text = text.replace(/\\begin\{(equation|align|gather|multline)(\*?)\}[\s\S]*?\\end\{\1\2\}/g, (m) => hold(m));
      const uni = (m, tex) => { const r = T.parse(tex); return r.error ? m : T.toUnicode(r.ast).trim(); };
      text = text.replace(/\\\(([\s\S]+?)\\\)/g, uni);
      text = text.replace(/\$([^\s$](?:[^$\n]*?[^\s$])?)\$/g, (m, tex) => (/[\\^_{}]/.test(tex) || /^[A-Za-z]$/.test(tex) ? uni(m, tex) : m));
    }
    // tabel yang disalin sebagai teks: sel dipisah Tab (Word/Excel). Baris yang sel terakhirnya terpotong ganti baris disambung.
    const tabTables = [];
    {
      const L = text.split("\n"), outL = [];
      const isMarker = (c) => /^([•▪◦·\-–*]|\d+[.)]|[a-zA-Z][.)]|[ivx]{1,4}[.)])$/.test(c.trim());
      const isTabRow = (l) => /\t/.test(l) && l.split("\t").filter((c) => c.trim()).length >= 2 && !isMarker(l.split("\t")[0]);
      for (let i = 0; i < L.length; i++) {
        if (!isTabRow(L[i]) || !(isTabRow(L[i + 1] || "") || isTabRow(L[i + 2] || ""))) { outL.push(L[i]); continue; }
        let j = i; const raw = [];
        while (j < L.length && (isTabRow(L[j]) || (L[j].trim() && !/\t/.test(L[j]) && raw.length && isTabRow(L[j + 1] || "") && L[j].length < 80))) { raw.push(L[j]); j++; }
        const cells = raw.map((l) => l.replace(/\t+\s*$/, "").split("\t").map((c) => c.trim()));
        const counts = cells.map((c) => c.length).sort((a, b) => a - b);
        const want = counts[Math.floor(counts.length * 0.75)] || counts[counts.length - 1];
        const rows = [];
        for (let r = 0; r < cells.length; r++) {
          let row = cells[r].slice();
          while (row.length < want && r + 1 < cells.length && row.length + cells[r + 1].length - 1 <= want) {
            const nx = cells[++r];
            row[row.length - 1] = (row[row.length - 1] + " " + nx[0]).trim();
            row = row.concat(nx.slice(1));
          }
          while (row.length < want) row.push("");
          rows.push(row.slice(0, Math.max(want, row.length)));
        }
        if (rows.length >= 2) { tabTables.push(rows); outL.push("\u0003T" + (tabTables.length - 1) + "\u0003"); i = j - 1; }
        else outL.push(...raw), i = j - 1;
      }
      text = outL.join("\n");
    }
    const rawLines = text.split("\n");
    const hasBlankSep = /\n[ \t ]*\n/.test(text.trim());
    let paras = [];
    if (hasBlankSep) {
      // paragraf dipisah baris kosong; baris di dalamnya disambung (khas PDF)
      let acc = [];
      for (const ln of rawLines) {
        if (/^\s*$/.test(ln)) { if (acc.length) paras.push(acc); acc = []; }
        else acc.push(ln.trim());
      }
      if (acc.length) paras.push(acc);
      paras = paras.map((ls) => {
        // jangan sambung bila baris pertama heading markdown
        if (ls.length > 1 && /^#{1,6}\s/.test(ls[0])) return [ls[0], ls.slice(1).join(" ")];
        // baris panjang (salinan Word/web, satu baris = satu paragraf) tidak disambung; baris pendek seragam (PDF) disambung
        if (merge && ls.length > 1) return mergeLines(ls); // hanya baris lanjutan yang disambung (PDF), judul/kalimat utuh tetap terpisah
        return [joinLines(ls)];
      }).flat();
    } else {
      const ls = rawLines.map((l) => l.trim()).filter((l) => l.length);
      if (merge && ls.length > 2) paras = mergeLines(ls);
      else paras = ls;
    }
    if (maths.length || tabTables.length) paras = paras.flatMap((p) => p.split(/\s*(\u0001M\d+\u0001|\u0003T\d+\u0003)\s*/).filter((x) => x.trim()));
    const blocks = [];
    for (let p of paras) {
      const tt = /^\u0003T(\d+)\u0003$/.exec(p);
      if (tt) { blocks.push({ kind: "table", rows: tabTables[+tt[1]] }); continue; }
      const mm = /^\u0001M(\d+)\u0001$/.exec(p);
      if (mm) {
        const r = T.parse(maths[+mm[1]]);
        if (!r.error) {
          const text0 = T.toUnicode(r.ast).trim();
          blocks.push({ kind: "p", tag: "p", cls: "", runs: [{ text: text0, b: false, i: false, sup: false, sub: false, link: "" }], math: { ast: r.ast, tex: r.tex, number: r.number, display: true, text: text0, source: "latex" } });
        }
        continue;
      }
      let tag = "p";
      const h = /^(#{1,6})\s+(.*)$/.exec(p);
      if (h) { tag = "h" + h[1].length; p = h[2]; }
      p = p.replace(/^\s*(?:[-*•▪◦·]\s+)/, ""); // bullet
      // paragraf berisi LaTeX: jangan tafsirkan "_" sebagai italic Markdown (x_i)
      const md = /\\[a-zA-Z]{2,}|[_^]\{/.test(p) ? [{ text: p }] : mdInline(p);
      const runs = normalizeRuns(md.map((r) => ({ text: r.text, b: !!r.b, i: !!r.i, sup: false, sub: false, link: "" })));
      if (runs.length) blocks.push({ kind: "p", tag, cls: "", runs, md: !!h });
    }
    return { blocks: attachEquationNumbers(blocks), images: 0, tables: tabTables.length, source: "text" };
  }

  function joinLines(ls) {
    return ls.reduce((acc, l) => {
      if (!acc) return l;
      if (/-$/.test(acc) && /^\p{Ll}/u.test(l)) return acc.slice(0, -1) + l; // pemenggalan kata
      return acc + " " + l;
    }, "");
  }

  function mergeLines(ls) {
    const longs = ls.filter((l) => l.length > 40).map((l) => l.length).sort((a, b) => a - b);
    const median = longs.length ? longs[Math.floor(longs.length / 2)] : 0;
    // Satu baris = satu paragraf (salinan dari Word/web) bila banyak baris sangat panjang.
    // Baris hasil salin PDF umumnya pendek & seragam (< 130 karakter).
    const veryLong = ls.filter((l) => l.length > 160).length;
    if (veryLong >= Math.max(1, ls.length * 0.25)) {
      // tetap sambung baris yang jelas lanjutan (diawali huruf kecil)
      const out0 = [];
      for (const l of ls) {
        const prev = out0[out0.length - 1];
        if (prev != null && !endsSentence(prev) && /^\p{Ll}/u.test(l) && startsLower(l)) out0[out0.length - 1] = joinLines([prev, l]);
        else out0.push(l);
      }
      return out0;
    }
    const out = [];
    for (const l of ls) {
      const prev = out[out.length - 1];
      if (prev != null && !/^#{1,6}\s/.test(prev) && !endsSentence(prev) &&
          (startsLower(l) || (median > 50 && prev.length >= median * 0.8 && prev.length > 50))) {
        out[out.length - 1] = joinLines([prev, l]);
      } else out.push(l);
    }
    return out;
  }

  /**
   * Titik masuk tunggal: data clipboard -> blocks.
   * @param {{html?:string, text?:string}} clip
   */
  function parseInput(clip, sec, profile, ParserImpl) {
    profile = profile || DEFAULT_PROFILE;
    const isRef = sec && sec.kind === "references";
    const html = clip.html || "";
    // HTML dipakai bila memuat struktur paragraf; jika tidak, pakai teks biasa
    if (html && (/<(p|div|h[1-6]|li|br|table|tr|dd|dt|blockquote|math|mjx-container)\b/i.test(html) || /katex|mwe-math|msEquation/i.test(html))) {
      const r = parseHtml(html, ParserImpl, { flattenTables: !!(sec && sec.kind === "abstract") });
      if (r.blocks.length) return isRef ? r : recoverFlatTables(recoverLatex(r));
    }
    const r = parseText(clip.text || "", { mergeBrokenLines: profile.detect.mergeBrokenLines && !isRef });
    return isRef ? r : recoverFlatTables(recoverLatex(r));
  }

  /* ---------------- Tempel keseluruhan: pecah artikel per Heading 1 ---------------- */
  // nama lain judul bagian (KAPITAL) yang umum di naskah, dipetakan ke kunci bagian template
  const WHOLE_SYNONYMS = [
    ["abstract", ["ABSTRACT", "ABSTRAK"]],
    ["intro", ["INTRODUCTION", "PENDAHULUAN", "BACKGROUND", "INTRODUCTION AND BACKGROUND"]],
    ["method", ["METHOD", "METHODS", "METHODOLOGY", "RESEARCH METHOD", "RESEARCH METHODS", "RESEARCH METHODOLOGY", "MATERIALS AND METHODS", "MATERIAL AND METHODS", "METODE", "METODE PENELITIAN", "METODOLOGI", "METODOLOGI PENELITIAN"]],
    ["rd", ["RESULTS AND DISCUSSION", "RESULTS AND DISCUSSIONS", "RESULT AND DISCUSSION", "FINDINGS AND DISCUSSION", "RESULTS & DISCUSSION", "HASIL DAN PEMBAHASAN", "HASIL PENELITIAN DAN PEMBAHASAN"]],
    ["rd:Results", ["RESULTS", "RESULT", "FINDINGS", "HASIL", "HASIL PENELITIAN"]],
    ["rd:Discussion", ["DISCUSSION", "DISCUSSIONS", "PEMBAHASAN"]],
    ["concl", ["CONCLUSION", "CONCLUSIONS", "CONCLUSION AND SUGGESTION", "CONCLUSIONS AND SUGGESTION", "CONCLUSIONS AND SUGGESTIONS", "CONCLUSION AND RECOMMENDATION", "CONCLUSION AND RECOMMENDATIONS", "CONCLUSIONS AND RECOMMENDATIONS", "KESIMPULAN", "SIMPULAN", "KESIMPULAN DAN SARAN", "PENUTUP"]],
    ["ack", ["ACKNOWLEDGMENT", "ACKNOWLEDGMENTS", "ACKNOWLEDGEMENT", "ACKNOWLEDGEMENTS", "AKNOWLEDGMENT", "UCAPAN TERIMA KASIH"]],
    ["contrib", ["AUTHOR CONTRIBUTION STATEMENT", "AUTHOR CONTRIBUTIONS", "AUTHOR CONTRIBUTION", "AUTHORS CONTRIBUTIONS", "AUTHORS' CONTRIBUTIONS", "CREDIT AUTHORSHIP CONTRIBUTION STATEMENT", "KONTRIBUSI PENULIS"]],
    ["ai", ["AI DISCLOSURE STATEMENT", "AI DISCLOSURE", "DECLARATION OF AI USE", "DECLARATION OF GENERATIVE AI USE", "DECLARATION OF GENERATIVE AI AND AI-ASSISTED TECHNOLOGIES IN THE WRITING PROCESS"]],
    ["refs", ["REFERENCES", "REFERENCE", "BIBLIOGRAPHY", "DAFTAR PUSTAKA", "REFERENSI", "LITERATURE CITED", "WORKS CITED"]],
  ];
  // bagian penutup yang tidak ada di template -> dilewati (bukan ditempel ke bagian sebelumnya)
  const RE_BACKMATTER = /^(CONFLICTS? OF INTERESTS?|COMPETING INTERESTS?|DECLARATION OF COMPETING INTERESTS?|DECLARATIONS?|FUNDING( STATEMENT| INFORMATION)?|DATA AVAILABILITY( STATEMENT)?|ETHICS?( STATEMENT| APPROVAL)?|ETHICAL (APPROVAL|STATEMENT|CONSIDERATIONS?)|APPENDIX( [A-Z0-9]+)?|APPENDICES|ABBREVIATIONS|SUPPLEMENTARY MATERIALS?|BIOGRAPH(Y|IES)|ABOUT THE AUTHORS?|AUTHOR BIOGRAPHIES|ACKNOWLEDG?E?MENTS?|AKNOWLEDGMENT|UCAPAN TERIMA KASIH|LAMPIRAN|KONFLIK KEPENTINGAN|PENDANAAN|INFORMED CONSENT|NOTES?)$/;
  /** Peta judul (normHead) -> {key, sub?} untuk profil ini. */
  function wholeHeadingMap(profile) {
    const map = new Map();
    const has = (k) => profile.sections.some((s) => s.key === k);
    for (const [k0, names] of WHOLE_SYNONYMS) {
      const [k, sub] = k0.split(":");
      if (!has(k)) continue;
      for (const n of names) if (!map.has(normHead(n))) map.set(normHead(n), { key: k, sub: sub || null });
    }
    // judul & nama lain dari template (termasuk bagian buatan pengguna) — prioritas tertinggi
    for (const s of profile.sections) for (const n of sectionNames(s)) map.set(n, { key: s.key, sub: null });
    return map;
  }
  /**
   * Teks satu paragraf -> heading 1? {key, sub} untuk bagian yang dikenal, {unknown:true} untuk judul KAPITAL lain, null bila bukan.
   * Heading 1 = KAPITAL semua, pendek, boleh bernomor (1. / IV. / A.). "Abstract" boleh huruf biasa.
   */
  function wholeHeading(text, map, opts) {
    let t = String(text || "").replace(/^#{1,6}\s+/, "").replace(/^\*\*(.*)\*\*$/, "$1").replace(/\s+/g, " ").trim();
    if (!t || t.length > 90 || countWords(t) > 10) return null;
    const letters = t.replace(/[^\p{L}]/gu, "");
    const caps = letters.length >= 3 && letters === letters.toUpperCase() && letters !== letters.toLowerCase();
    const n = normHead(t);
    const hit = map.get(n);
    if (hit && (caps || hit.key === "abstract")) return hit;
    if (hit && !hit.sub && opts && opts.headingLike) return Object.assign({ weak: true }, hit); // "Method" bergaya Heading/bold (huruf kapital dari format Word)
    if (!caps) return null;
    // judul KAPITAL lain (mis. LITERATURE REVIEW): >= 2 kata, hanya huruf/spasi/&/-
    const core = t.replace(/^(?:(?:\d+\.)*\d+|[IVX]{1,5}|[A-H])[.)]?\s+/, "");
    if (RE_BACKMATTER.test(core)) return { unknown: true, back: true, title: core };
    if (/^[\p{Lu}][\p{Lu}'’&\- ]+$/u.test(core) && core.split(/\s+/).filter((w) => /\p{L}{2,}/u.test(w)).length >= 2) return { unknown: true, title: core };
    return null;
  }
  /**
   * Pecah satu tempelan artikel utuh menjadi bagian-bagian template.
   * return { parts: {key: {blocks, images, tables}}, order: [key], unknown: [{title, key|null}], front: n, notFound: [key] }
   */
  function splitWhole(clip, profile, ParserImpl) {
    profile = profile || DEFAULT_PROFILE;
    const map = wholeHeadingMap(profile);
    const keys = profile.sections.map((s) => s.key);
    const conclIdx = keys.indexOf("concl");
    const isMain = (k) => k && profile.sections.find((s) => s.key === k).kind !== "abstract" && profile.sections.find((s) => s.key === k).kind !== "references" && (conclIdx < 0 || keys.indexOf(k) < conclIdx);
    const res = { parts: {}, order: [], unknown: [], front: 0, images: 0 };
    const html = clip.html || "";
    const useHtml = html && (/<(p|div|h[1-6]|li|br|table|tr)\b/i.test(html) || /msEquation|katex/i.test(html));
    // 1) satuan: blok HTML atau baris teks
    let units;
    if (useHtml) {
      const r = parseHtml(html, ParserImpl, { flattenFront: true });
      res.images = r.images || 0;
      units = r.blocks.map((b) => ({ b, text: b.kind === "p" && !b.math ? plain(b.runs) : "", headingLike: b.kind === "p" && (/^h[1-6]$/.test(b.tag || "") || b.runs.every((x) => x.b || !/\S/.test(x.text))) }));
    } else {
      const ls = String(clip.text || "").replace(/\r\n?/g, "\n").split("\n");
      units = ls.map((l) => ({ line: l, text: l, headingLike: true }));
    }
    // 2) potong per heading 1
    let cur = null;
    const segs = {};
    const add = (k, u) => { (segs[k] = segs[k] || []).push(u); };
    for (const u of units) {
      let h = u.text.trim() ? wholeHeading(u.text, map, { headingLike: u.headingLike }) : null;
      // judul huruf biasa hanya diterima bila urutannya sesuai template (sesudah bagian yang sedang aktif)
      if (h && h.weak && !(keys.indexOf(h.key) > keys.indexOf(cur) && !res.order.includes(h.key))) h = null;
      if (h && h.key) {
        cur = h.key;
        if (!res.order.includes(cur)) res.order.push(cur);
        if (h.sub) { // RESULTS / DISCUSSION terpisah -> sub-heading utama di bagian Results and Discussion
          const t = titleCase(h.sub.toLowerCase());
          add(cur, useHtml ? { b: { kind: "p", tag: "h2", cls: "", runs: [{ text: t, b: true, i: false, sup: false, sub: false, link: "" }], html: true } } : { line: t });
        }
        continue;
      }
      if (h && h.back && cur) { res.unknown.push({ title: h.title, key: null }); cur = "__skip"; continue; }
      if (h && h.unknown && isMain(cur)) {
        res.unknown.push({ title: h.title, key: cur });
        const t = h.title;
        add(cur, useHtml ? { b: { kind: "p", tag: "h2", cls: "", runs: [{ text: t, b: true, i: false, sup: false, sub: false, link: "" }], html: true } } : { line: t });
        continue;
      }
      if (h && h.unknown && cur && !isMain(cur) && cur !== "abstract") { res.unknown.push({ title: h.title, key: null }); cur = "__skip"; continue; }
      if (!cur) { if (u.text.trim() || u.b) res.front++; continue; }
      if (cur === "__skip") continue;
      add(cur, u);
    }
    // 3) tiap bagian diproses seperti tempel per bagian
    for (const k of Object.keys(segs)) {
      const sec = profile.sections.find((s) => s.key === k);
      let r;
      if (useHtml) {
        const blocks = segs[k].map((u) => u.b);
        r = { blocks, images: 0, tables: blocks.filter((b) => b.kind === "table").length, source: "html" };
        if (sec.kind !== "references") r = recoverFlatTables(recoverLatex(r));
      } else {
        r = parseInput({ text: segs[k].map((u) => u.line).join("\n") }, sec, profile);
      }
      r.blocks = r.blocks.filter((b) => b.kind === "table" || b.math || plain(b.runs).trim());
      if (sec.kind === "abstract") { // nama penulis/afiliasi yang terletak setelah kotak abstrak
        let last = -1;
        r.blocks.forEach((b, i) => { if (b.kind === "p" && countWords(plain(b.runs)) >= 25) last = i; });
        if (last >= 0 && r.blocks.slice(last + 1).every((b) => b.kind === "p" && countWords(plain(b.runs)) <= 12)) r.blocks = r.blocks.slice(0, last + 1);
      }
      if (r.blocks.length) res.parts[k] = r;
    }
    res.order = res.order.filter((k) => res.parts[k]);
    res.notFound = keys.filter((k) => !res.parts[k]);
    return res;
  }

  /* ---------------- tabel yang tersalin sebagai baris-baris teks ---------------- */
  /**
   * "Table 4. …" lalu sel-sel tabel satu per baris (salinan teks dari Word/ChatGPT/PDF) -> tabel lagi.
   * Jumlah kolom ditebak: kolom pertama berisi label (bukan angka), kolom lain campuran nilai.
   */
  function recoverFlatTables(res) {
    const blocks = res.blocks, out = [];
    const isCell = (b) => b && b.kind === "p" && !b.math && (() => { const t = plain(b.runs).trim(); return t && countWords(t) <= 14 && !RE_CAPTION.test(t) && !RE_NOTE.test(t) && !/[.?!:]$/.test(t.replace(/\b\d+\.\d+$|\bet al\.$|\be\.g\.$/, "")); })();
    const numeric = (c) => /^[-−–+±<>≤≥~≈]?\s*\(?\d/.test(c) || /^[-−–]$/.test(c);
    for (let i = 0; i < blocks.length; i++) {
      out.push(blocks[i]);
      const cap = blocks[i].kind === "p" ? plain(blocks[i].runs).trim() : "";
      if (!/^(Table|Tabel)\s+\d+[A-Za-z]?\s*[.:]/i.test(cap)) continue;
      let j = i + 1; const cells = [];
      while (j < blocks.length && isCell(blocks[j])) { cells.push(plain(blocks[j].runs).trim()); j++; }
      if (cells.length < 4) continue;
      let best = null;
      for (let k = 2; k <= 8; k++) {
        if (cells.length % k || cells.length / k < 2) continue;
        const rows = []; for (let r = 0; r < cells.length; r += k) rows.push(cells.slice(r, r + k));
        const data = rows.slice(1);
        const s0 = data.filter((r) => !numeric(r[0])).length / data.length;
        let pen = 0;
        for (let c = 1; c < k; c++) if (data.length >= 2 && data.every((r) => !numeric(r[c]) && countWords(r[c]) <= 4) && data.some((r, x) => r[c] !== data[0][c])) pen++;
        const anyNum = data.some((r) => r.slice(1).some(numeric));
        const score = s0 - pen / k + (anyNum ? 0.05 : 0) - k * 0.001;
        if (!best || score > best.score) best = { k, rows, score };
      }
      if (!best || best.score < 0.6) continue;
      out.push({ kind: "table", rows: best.rows, recovered: true });
      i = j - 1;
    }
    return Object.assign({}, res, { blocks: out });
  }

  /* ---------------- LaTeX "polos" (mis. salinan ChatGPT yang backslash-nya hilang) ---------------- */
  const latexish = (t) => /\\[a-zA-Z]{2,}|[_^]\{|\\[,;!{}|]/.test(t);
  const OPEN_M = /^(\\\[|\[|\$\$)$/, CLOSE_M = /^(\\\]|\]|\$\$)$/;
  function mathBlock(tex, T) {
    const r = T.parse(tex);
    if (r.error || !r.ast || !r.ast.items || !r.ast.items.length) return null;
    const text = T.toUnicode(r.ast).trim();
    return { kind: "p", tag: "p", cls: "", runs: [{ text, b: false, i: false, sup: false, sub: false, link: "" }], math: { ast: r.ast, tex: r.tex, number: r.number, display: true, text, source: "latex" } };
  }
  /** Cari rumus dalam kalimat: \(…\), $…$, dan (…) berisi LaTeX (backslash \( hilang). */
  function inlineMathSpans(text, lost) {
    const spans = [];
    let m;
    const reP = /\\\(([\s\S]+?)\\\)/g;
    while ((m = reP.exec(text))) spans.push([m.index, m.index + m[0].length, m[1]]);
    const reD = /\$([^\s$](?:[^$\n]*?[^\s$])?)\$/g;
    while ((m = reD.exec(text))) if (/[\\^_{}]/.test(m[1]) || /^[A-Za-z]$/.test(m[1])) spans.push([m.index, m.index + m[0].length, m[1]]);
    // ( … ) seimbang
    for (let i = 0; i < text.length; i++) {
      if (text[i] !== "(" || text[i - 1] === "\\") continue;
      let d = 0, j = i;
      for (; j < text.length; j++) { if (text[j] === "(") d++; else if (text[j] === ")") { d--; if (!d) break; } }
      if (j >= text.length) continue;
      const c = text.slice(i + 1, j).trim();
      const strong = /\\[a-zA-Z]{2,}/.test(c) && !/\s\p{Ll}{3,}\s+\p{Ll}{3,}\s/u.test(c.replace(/\\[a-zA-Z]+\{[^}]*\}/g, ""));
      const weak = lost && (/[_^]/.test(c) || /^\{[^{}]*\}$/.test(c) || /^[A-Za-z]$/.test(c) || /^[A-Za-z]\s*[=<>]\s*[\d.,]+$/.test(c) || /^\d*\.\d+$/.test(c)) && !/\s\p{L}{3,}\s/u.test(c);
      if (strong || weak) { spans.push([i, j + 1, c]); i = j; }
    }
    spans.sort((a, b) => a[0] - b[0]);
    return spans.filter((sp, k) => !spans.slice(0, k).some((o) => sp[0] < o[1]));
  }
  /**
   * Cari rumus baris LaTeX dalam daftar teks paragraf: "[" … "]" terpisah, "[ … ]" satu paragraf, atau "]" tanpa pembuka.
   * return [{start, end, tex}] (indeks inklusif)
   */
  function findDisplayLatex(texts) {
    const groups = [];
    for (let i = 0; i < texts.length; i++) {
      const t = (texts[i] || "").trim();
      if (OPEN_M.test(t)) {
        let j = i + 1; const lines = [];
        while (j < texts.length && j - i < 60 && !CLOSE_M.test((texts[j] || "").trim())) { lines.push(texts[j] || ""); j++; }
        if (j < texts.length && lines.length && (latexish(lines.join(" ")) || /=/.test(lines.join("")))) { groups.push({ start: i, end: j, tex: lines.join("\n") }); i = j; continue; }
      }
      const one = /^(?:\\\[|\[(?=\s))\s*([\s\S]+?)\s*(?:\\\]|\])$/.exec(t) || /^\$\$([\s\S]+)\$\$$/.exec(t);
      if (one && (latexish(one[1]) || /=/.test(one[1]))) { groups.push({ start: i, end: i, tex: one[1] }); continue; }
      if (CLOSE_M.test(t) || (/\s\]$/.test(t) && latexish(t) && !/^\[/.test(t))) {
        const lines = CLOSE_M.test(t) ? [] : [t.replace(/\s*\]$/, "")];
        const floor = groups.length ? groups[groups.length - 1].end + 1 : 0;
        let k = i - 1;
        while (k >= floor && i - k <= 30) {
          const tt = (texts[k] || "").trim();
          if (!tt || !(latexish(tt) || /^[=+\-−]/.test(tt) || /^[+\-=]$/.test(tt)) || isProse(tt)) break;
          lines.unshift(tt); k--;
        }
        if (lines.length && latexish(lines.join(" "))) groups.push({ start: k + 1, end: i, tex: lines.join("\n") });
      }
    }
    return groups;
  }
  /** Kalimat biasa (bukan baris rumus): ada >= 3 kata berurutan di luar perintah LaTeX. */
  const isProse = (t) => /\p{L}{3,}\s+\p{L}{2,}\s+\p{L}{3,}/u.test(String(t).replace(/\\[a-zA-Z]+(\{[^{}]*\})*/g, " ").replace(/[_^]\{[^{}]*\}/g, " "));
  const isLostLatex = (all, foundDisplay) => !!foundDisplay || (String(all).match(/\(\\[a-zA-Z]/g) || []).length >= 2;
  function recoverLatex(res) {
    const T = TM();
    if (!T) return res;
    const blocks = res.blocks;
    const texts = blocks.map((b) => (b.kind === "p" && !b.math ? plain(b.runs) : ""));
    const all = texts.join("\n");
    if (!/\\[a-zA-Z]{2,}|[_^]\{|\$/.test(all)) return res;
    const groups = findDisplayLatex(texts);
    const out = [];
    let g = 0, made = 0;
    for (let i = 0; i < blocks.length; i++) {
      const gr = groups[g];
      if (gr && gr.start === i) {
        const mb = mathBlock(gr.tex, T);
        g++;
        if (mb) { out.push(mb); i = gr.end; made++; continue; }
      }
      out.push(blocks[i]);
    }
    let lost = isLostLatex(all, made);
    // rumus di dalam kalimat -> runs berformat (miring, subskrip, superskrip)
    const final = out.map((b) => {
      if (b.kind !== "p" || b.math) return b;
      const text = plain(b.runs);
      const spans = inlineMathSpans(text, lost);
      if (!spans.length) return b;
      let runs = [], pos = 0, changed = false;
      for (const [a, e, tex] of spans) {
        const r = T.parse(tex);
        if (r.error || !r.ast) continue;
        runs = runs.concat(sliceRuns(b.runs, pos, a));
        runs = runs.concat(T.toRuns(r.ast).map((x) => Object.assign({ link: "" }, x, { mathInline: true })));
        pos = e; changed = true;
      }
      if (!changed) return b;
      runs = runs.concat(sliceRuns(b.runs, pos, text.length));
      return Object.assign({}, b, { runs: normalizeRuns(runs) });
    });
    return Object.assign({}, res, { blocks: attachEquationNumbers(final) });
  }

  /* ------------------------------------------------------------------ */
  /*  Klasifikasi                                                        */
  /* ------------------------------------------------------------------ */
  const RE_CAPTION = /^(Table|Tables|Figure|Fig\.|Tabel|Gambar|Grafik|Chart|Diagram|Algorithm|Algoritma)\s+(\d+[A-Za-z]?|[IVX]+)\s*[.:]/i;
  const RE_LABEL_ONLY = /^(Table|Tabel|Figure|Fig\.|Gambar|Grafik|Chart|Diagram|Algorithm|Algoritma)\s+(\d+[A-Za-z]?|[IVX]+)\s*[.:]?\s*$/i;
  /** Caption gaya APA 7 dua baris ("Figure 1" lalu judul italic) -> satu caption "Figure 1. Judul". */
  function mergeCaptionLabels(blocks) {
    const out = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i], n = blocks[i + 1];
      const t = b.kind === "p" ? plain(b.runs).trim() : "";
      if (t && RE_LABEL_ONLY.test(t) && n && n.kind === "p" && countWords(plain(n.runs)) <= 40 && !RE_CAPTION.test(plain(n.runs).trim())) {
        const label = t.replace(/[.:]?\s*$/, ".");
        out.push(Object.assign({}, n, { tag: "p", runs: normalizeRuns([{ text: label + " ", b: true }].concat(n.runs)), mergedCaption: true }));
        i++;
        continue;
      }
      out.push(b);
    }
    return out;
  }
  const RE_NOTE = /^(Notes?|Catatan|Sumber|Source|Keterangan)\s*[.:]/i;
  const RE_CITATION = /\([^()]*\b(1[89]|20)\d{2}[a-z]?\b[^()]*\)|\bet al\./i;
  const RE_NUMBERED = /^((?:\d+\.)*\d+)[.)]?\s+(?=[\p{L}“"'(])|^([A-H]|[IVX]{1,4})[.)]\s+(?=\p{Lu})/u;

  const normHead = (s) => String(s || "").replace(WS, " ").trim()
    .replace(/^(?:(?:\d+\.)*\d+|[IVX]{1,5}|[A-H])[.)]?\s+/, "")
    .replace(/[.:;]+$/, "").trim().toUpperCase();

  function sectionNames(sec) {
    return [sec.title].concat(sec.aliases || []).map(normHead);
  }

  /** Baris persamaan (mis. "Y = β0 + β1X1 + ε" atau "… (1)"). */
  const PLAIN_MATH = (s) => String(s).replace(/[\u{1D400}-\u{1D7FF}\u210E]/gu, (c) => c.normalize("NFKC"));
  const EQ_WORDS = new Set(["log", "ln", "exp", "sin", "cos", "tan", "max", "min", "lim", "sup", "inf", "det", "var", "cov", "corr", "prob", "sgn", "mod", "where", "whereas", "with", "dengan", "dimana", "which"]);
  function looksLikeEquation(text) {
    const t = PLAIN_MATH(text).replace(/^[\s•▪◦·\-–]+/, "").trim();
    // kalimat/definisi ("Yit = District-level economic growth …", "Volume Loading (VOLLOADING = …)") bukan rumus
    const prose = (t.match(/[A-Za-z][a-z]{3,}/g) || []).filter((w) => !EQ_WORDS.has(w.toLowerCase()));
    if (prose.length >= 2 || (prose.length === 1 && t.indexOf(prose[0]) === 0 && /^[A-Za-z][a-z]{3,}\s/.test(t))) return /^\(\d+[a-z]?\)$/.test(t);
    if (!t || countWords(t) > 30) return false;
    if (/^\(\d+[a-z]?\)$/.test(t)) return true;
    if (!/[=≤≥≈]/.test(t)) return false;
    if (/\b(the|and|of|is|are|was|were|that|with|yang|dan|adalah)\b/i.test(t)) return false;
    return /[+\-−×÷∑∏√∫^βαγδεθλμσπρτφχψω₀-₉]|\b[A-Za-z]\d|\b[A-Za-z]\s*[=+]/u.test(t);
  }

  /** Baris atribusi kutipan (nama/peran partisipan) yang berdiri sendiri setelah kutipan. */
  const RE_ROLE = /\b(informant|participant|respondent|interviewee|teacher|student|manager|head|officer|staff|parent|principal|lecturer|director|practitioner|auditor|employee|translated|informan|narasumber|responden|partisipan|guru|siswa|kepala|dosen|pegawai|diterjemahkan|P\d+|R\d+|S\d+)\b/i;
  function looksLikeAttribution(text, runs) {
    const t = text.trim();
    if (!t || countWords(t) > 25 || /[.!?]\s+\p{Lu}/u.test(t)) return false;
    const allItalic = runs.length && runs.every((r) => r.i || !/\S/.test(r.text));
    const wrapped = /^[(\[].*[)\]]\.?$/.test(t);
    const dashed = /^[—–-]\s*\S/.test(t);
    return wrapped || dashed || allItalic || RE_ROLE.test(t);
  }

  function looksLikeHeading(text, runs, block, det) {
    if (looksLikeEquation(text)) return false;
    const t = text.trim();
    const words = countWords(t);
    // sinyal kuat (bold penuh / bernomor / "Theme 1:") boleh lebih panjang
    const strongSig = (runs.length && runs.every((r) => r.b || !/\S/.test(r.text))) || RE_NUMBERED.test(t) || /^(Theme|Tema|Sub-?theme|Subtema)\s+\d+\s*[:.]/i.test(t) || (block.tag && /^h[1-6]$/.test(block.tag));
    const maxW = strongSig ? Math.max(det.maxHeadingWords || 15, 25) : (det.maxHeadingWords || 15);
    if (!words || words > maxW || t.length > (strongSig ? 220 : 160)) return false;
    if (/^\d+(\.\d+)*\.?$/.test(t)) return false;
    if (RE_CAPTION.test(t) || RE_NOTE.test(t)) return false;
    if (!(block.tag && /^h[1-6]$/.test(block.tag)) && (/[=;<>≤≥]/.test(t) || /^\(.*\)$/.test(t) || /^[↓↑→←⇓⇒]+$/.test(t))) return false; // statistik / keterangan dalam kurung / panah diagram
    const letters = (t.match(/\p{L}/gu) || []).length;
    const digits = (t.match(/\d/g) || []).length;
    if (letters < 3 || digits > letters * 0.5) return false;
    // sinyal format kuat dari HTML
    if (block.tag && /^h[1-6]$/.test(block.tag)) return true;
    const allBold = runs.length && runs.every((r) => r.b || !/\S/.test(r.text));
    const numbered = RE_NUMBERED.test(t);
    if (/[.,;]$/.test(t) && !(allBold && /\.$/.test(t) && words <= 8)) return false;
    if (/\?$/.test(t) && !allBold) return false;
    if (/:$/.test(t) && !allBold) return false;
    if (RE_CITATION.test(t)) return false;
    if (!/^[\p{Lu}\d"“'‘(\[]/u.test(t) && !/^[a-h][.)]\s+\p{Lu}/u.test(t)) return false; // "a. Maritime Transportation …" boleh
    if (allBold || numbered || /^(Theme|Tema|Sub-?theme|Subtema)\s+\d+\s*[:.]/i.test(t)) return true;
    // teks biasa: pendek, tanpa tanda titik akhir, huruf besar di awal
    return words <= (det.maxHeadingWords || 15);
  }

  /** Kutipan langsung (mis. wawancara): diawali tanda kutip, setelah tanda kutip penutup hanya atribusi pendek. */
  function looksLikeQuote(text, block) {
    if (block.tag === "blockquote") return true;
    const t = text.trim();
    if (!/^[“"‘«]/.test(t)) return false;
    const w = countWords(t);
    if (w < 3 || w > 250) return false;
    const close = Math.max(t.lastIndexOf("”"), t.lastIndexOf('"'), t.lastIndexOf("»"));
    if (close <= 0) return false;
    const tail = t.slice(close + 1);
    return countWords(tail) <= 20; // "Girl student, S3" / "(Teacher, S1)" / "— Parent"
  }

  function numberingDepth(t) {
    const m = /^((?:\d+\.)*\d+)\.?\s+/.exec(t);
    if (m) return m[1].split(".").length;
    if (/^[A-H][.)]\s+/.test(t)) return 2;
    if (/^[a-h][.)]\s+/.test(t)) return 4; // a. b. c. = sub-sub-heading
    if (/^[IVX]{1,4}[.)]\s+/.test(t)) return 1;
    if (/^\d+\)\s+/.test(t)) return 3; // gaya IEEE: IV. → A. → 1)
    return 0;
  }

  /**
   * @param {Array} blocks hasil parseHtml/parseText
   * @param {object} sec   konfigurasi bagian (dari profile.sections)
   * @param {object} profile
   */
  function classify(blocks, sec, profile) {
    profile = profile || DEFAULT_PROFILE;
    const det = profile.detect;
    sec = sec || {};
    if (sec.kind === "references") return classifyReferences(blocks, sec, profile);

    const myNames = new Set(sectionNames(sec));
    const majors = new Set((sec.majorSubheadings || []).map(normHead));
    const items = [];
    blocks = mergeCaptionLabels(blocks);

    for (let idx = 0; idx < blocks.length; idx++) {
      const bl = blocks[idx];
      if (bl.kind === "table") { items.push({ type: "table", rows: bl.rows, runs: [] }); continue; }
      if (bl.math) { items.push({ type: "equation", runs: bl.runs, math: bl.math }); continue; }
      let runs = bl.runs;
      const text = plain(runs);
      if (!text.trim()) continue;
      if (myNames.has(normHead(text)) && countWords(text) <= 8) { items.push({ type: "skip", runs, why: "judul bagian" }); continue; }
      if (RE_CAPTION.test(text) || RE_LABEL_ONLY.test(text.trim())) { items.push({ type: "caption", runs }); continue; }
      if (RE_NOTE.test(text)) { items.push({ type: "note", runs }); continue; }
      if (det.detectQuotes !== false && looksLikeQuote(text, bl)) { items.push({ type: "quote", runs }); continue; }
      // atribusi kutipan pada paragraf tersendiri (mis. "Risk Management Manager (translated …)")
      const prevIt = items[items.length - 1];
      if (prevIt && prevIt.type === "quote" && looksLikeAttribution(text, runs) && !runs.every((r) => r.b)) { items.push({ type: "quoteAttr", runs }); continue; }
      if (looksLikeEquation(text)) { items.push({ type: "equation", runs }); continue; }
      // heading tidak boleh paragraf terakhir yang berdiri sendiri? -> tetap boleh, tapi butuh sinyal kuat
      const isLast = idx === blocks.length - 1;
      const nextBl = blocks[idx + 1];
      const leadIn = (nextBl && nextBl.math) || /\b(is|are|be|as|by|gives|follows|where|that|then|of|to|the|and)\s*:?$/i.test(text.trim());
      if (looksLikeHeading(text, runs, bl, det)) {
        const strong = (bl.tag && /^h[1-6]$/.test(bl.tag)) || runs.every((r) => r.b || !/\S/.test(r.text)) || RE_NUMBERED.test(text.trim());
        if (leadIn && !strong) { items.push({ type: "body", runs }); continue; }
        if (!isLast || strong) {
          const next = blocks[idx + 1];
          const nextText = next && next.kind === "p" ? plain(next.runs) : "";
          // dua baris pendek berturut-turut tanpa sinyal kuat -> kemungkinan daftar, bukan heading
          // dua baris pendek berturut-turut = daftar, KECUALI keduanya diikuti paragraf panjang
          // (heading + sub-heading, mis. "Theoretical Background" lalu "Social media usage motivation")
          const next2 = blocks[idx + 2];
          const next2Text = next2 && next2.kind === "p" ? plain(next2.runs) : "";
          const headThenSub = nextText && countWords(next2Text) > (det.maxHeadingWords || 15) && !/[.:;,]$/.test(nextText.trim()) && !/[.:;,]$/.test(text.trim());
          const nextAlsoShort = nextText && countWords(nextText) <= 4 && !/^h[1-6]$/.test(next.tag || "") && !headThenSub;
          if (strong || !nextAlsoShort || countWords(nextText) === 0 || (next && next.kind === "table")) {
            items.push({ type: "heading", level: 0, runs, tag: bl.tag, depth: numberingDepth(text.trim()), major: majors.has(normHead(text)) });
            continue;
          }
        }
      }
      items.push({ type: "body", runs });
    }

    // Tentukan level heading
    const heads = items.filter((it) => it.type === "heading");
    const tagLevels = [...new Set(heads.map((h) => (h.tag && /^h[1-6]$/.test(h.tag)) ? +h.tag[1] : 99))].filter((x) => x < 99).sort((a, b) => a - b);
    const depths = [...new Set(heads.map((h) => h.depth).filter((d) => d > 0))].sort((a, b) => a - b);
    let seenMajor = false;
    for (const h of heads) {
      let lvl = 2;
      if (h.major) { lvl = 2; seenMajor = true; }
      else if (seenMajor) lvl = 3;
      else if (h.depth > 0 && depths.length > 1) lvl = h.depth === depths[0] ? 2 : 3;
      else if (h.tag && /^h[1-6]$/.test(h.tag) && tagLevels.length > 1) lvl = (+h.tag[1] === tagLevels[0]) ? 2 : 3;
      h.level = lvl;
      // bersihkan teks heading
      let r = h.runs;
      const t = plain(r);
      if (det.stripHeadingNumbers) {
        const m = RE_NUMBERED.exec(t);
        if (m) r = dropPrefix(r, m[0].length);
      }
      r = mapRunsText(r, (s) => s.replace(/^\*+|\*+$/g, ""));
      r = normalizeRuns(r).map((x) => Object.assign({}, x, { text: x.text }));
      if (det.titleCaseHeadings) r = mapRunsText(r, titleCase);
      // buang titik/titik dua di akhir heading
      if (r.length) r[r.length - 1].text = r[r.length - 1].text.replace(/[.:]\s*$/, "");
      h.runs = normalizeRuns(r);
      delete h.tag; delete h.depth; delete h.major;
    }
    return items;
  }

  /* ------------------------------------------------------------------ */
  /*  Referensi                                                          */
  /* ------------------------------------------------------------------ */
  const RE_REF_START = /^(?:\[\d+\]\s*|\d+\.\s+)?(?:(?:van|von|de|der|den|da|di|du|del|della|le|la|al|el|bin|binti|ten|ter)\s+)*(?:[\p{Lu}][\p{L}'’\-]+(?:[ -][\p{L}'’\-]+){0,3},\s*(?:[\p{Lu}]\.|[\p{Lu}][\p{Ll}]+)|[\p{Lu}][\p{L}&.,'’\- ]{1,120}?\s\((?:\d{4}[a-z]?|n\.d\.|in press)[^)]*\)|[\p{Lu}][\p{L}'’\-]+(?: [\p{Lu}][\p{L}'’\-]+)*\.?\s\((?:\d{4}|n\.d\.))/u;

  const RE_PARTICLE = /^(?:van|von|de|der|den|da|di|du|del|della|le|la|al|el|bin|binti|ten|ter)\s+\p{Lu}[\p{L}'’\-]+,/u;
  function splitRefBlocks(blocks, det) {
    // pecah blok menjadi baris; untuk teks biasa sambung baris-baris lanjutan
    let lines = [];
    for (const b of blocks) {
      if (b.kind === "table") { b.rows.forEach((row) => lines.push({ runs: normalizeRuns([{ text: row.join(" ") }]) })); continue; }
      lines.push({ runs: b.runs, fromHtml: !!b.html });
    }
    lines = lines.filter((l) => plain(l.runs).trim());
    if (!det.mergeBrokenLines || lines.length < 2) return lines;
    // HTML (Word/Google Docs): setiap paragraf sudah satu referensi
    if (lines.length && lines.every((l) => l.fromHtml)) return lines;
    const starts = lines.map((l) => RE_REF_START.test(plain(l.runs).trim()));
    const nonStart = starts.filter((s) => !s).length;
    if (nonStart === 0) return lines;
    // mode sambung: baris yang bukan awal referensi digabung ke referensi sebelumnya
    const out = [];
    lines.forEach((l, k) => {
      const t = plain(l.runs).trim();
      const prev = out[out.length - 1];
      const prevT = prev ? plain(prev.runs).trim() : "";
      const prevLine = k > 0 ? plain(lines[k - 1].runs).trim() : "";
      const prevEndsRef = /https?:\/\/\S+$|doi\.org\/\S+$/i.test(prevLine) ||
        (/[.)]$/.test(prevLine) && !/(?:^|[\s,(])\p{Lu}\.$/u.test(prevLine)); // "…, D." = inisial penulis, belum selesai
      const continuation = prev && !/^\[\d+\]\s/.test(t) && (!starts[k] || (startsLower(t) && !RE_PARTICLE.test(t)) || !prevEndsRef || /[,&]$|\band$/.test(prevLine));
      if (continuation) {
        const glue = /-$/.test(prevT) && /^\p{Ll}/u.test(t) ? "" : " ";
        if (!glue) prev.runs[prev.runs.length - 1].text = prev.runs[prev.runs.length - 1].text.replace(/-$/, "");
        prev.runs = normalizeRuns(prev.runs.concat([{ text: glue }], l.runs));
      } else out.push({ runs: l.runs.slice() });
    });
    return out;
  }

  /** Miringkan "Nama Jurnal, Volume" (APA 7) bila tempelan tidak membawa italic. */
  function apaAutoItalic(runs) {
    if (runs.some((r) => r.i)) return runs;
    const t = plain(runs);
    // ... (2020). Judul artikel. Nama Jurnal, 12(3), 45–67.
    const yr = /\((?:\d{4}[a-z]?|n\.d\.|in press)[^)]*\)\.\s+/i.exec(t);
    if (!yr) return runs;
    const after = yr.index + yr[0].length;
    const rest = t.slice(after);
    const j = /^(.+?[.?!])\s+([^.?!]*?\S),\s*(\d+)(?=\s*(?:\(|,|\.|$))/.exec(rest);
    if (j && j[2].length > 2 && !/^(In|Dalam)\s/.test(j[2]) && !/https?:/i.test(j[2])) {
      const jStart = rest.indexOf(j[2], j[1].length);
      const volEnd = rest.indexOf(j[3], jStart + j[2].length) + j[3].length;
      if (jStart > 0 && volEnd > jStart) return patchRange(runs, after + jStart, after + volEnd, { i: true });
    }
    // Laporan / halaman web tanpa penerbit: (2020). Judul laporan. https://…  -> judul italic
    const web = /^(.+?)\.\s+(https?:\/\/\S+)\s*$/.exec(rest);
    if (web && !/\d+\(\d+\)/.test(web[1]) && !/^In\s/i.test(web[1]) && web[1].indexOf(". ") < 0) {
      let end = web[1].length;
      const tail = /\s+(\([^()]*\)|\[[^\[\]]*\])\s*$/.exec(web[1]);
      if (tail) end -= tail[0].length;
      return patchRange(runs, after, after + end, { i: true });
    }
    // Buku: (2020). Judul buku (edisi). Penerbit.  -> miringkan judul bila tidak ada pola jurnal
    const bk = /^([^.]+(?:\.[^.\s][^.]*)*?)(?:\s\([^)]*ed\.[^)]*\))?\.\s+[\p{Lu}][^.]+\.?\s*(?:https?:\/\/\S+)?\s*$/u.exec(rest);
    if (bk && !/\d+\(\d+\)/.test(rest) && !/,\s*\d+\s*[,(]/.test(rest) && !/^In\s/i.test(rest)) {
      let end = bk[1].length;
      const tail = /\s+(\([^()]*\)|\[[^\[\]]*\])\s*$/.exec(bk[1]);
      if (tail) end -= tail[0].length;
      return patchRange(runs, after, after + end, { i: true });
    }
    return runs;
  }

  /**
   * IEEE: miringkan nama jurnal / judul buku-prosiding bila tempelan tidak membawa italic.
   *   … "Judul artikel," Nama Jurnal, vol. 7, …      -> Nama Jurnal italic
   *   … "Judul bab," in Judul Buku, Ed., Penerbit, …  -> Judul Buku italic
   */
  function ieeeAutoItalic(runs) {
    if (runs.some((r) => r.i)) return runs;
    const t = plain(runs);
    const q = /[,.]["”]\s+/.exec(t);
    if (!q) return runs;
    let a = q.index + q[0].length;
    const rest = t.slice(a);
    let end;
    if (/^in\s/i.test(rest)) {
      a += 3;
      const c = /,\s/.exec(t.slice(a));
      end = c ? a + c.index : -1;
    } else {
      const c = /,\s*(?:vol\.|no\.|pp\.|Art\.\s*no\.|art\.|\d{4}\b|\d+\s*,|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s)/.exec(rest);
      end = c ? a + c.index : -1;
    }
    if (end <= a || end - a > 200) return runs;
    return patchRange(runs, a, end, { i: true });
  }

  function linkUrls(runs) {
    const t = plain(runs);
    const re = /\bhttps?:\/\/[^\s<>"]+[^\s<>".,;)\]]|\bdoi:\s*(10\.\d{4,9}\/\S+[^\s.,;)])/gi;
    let m, out = runs;
    const spans = [];
    while ((m = re.exec(t))) spans.push([m.index, m.index + m[0].length, m[1] ? "https://doi.org/" + m[1] : m[0]]);
    for (const [s, e, url] of spans) out = patchRange(out, s, e, { link: url });
    return out;
  }

  /* ---------------- referensi ganda ---------------- */
  const refNum = (t) => { const m = /^\s*\[(\d+)\]/.exec(t) || /^\s*(\d+)\.\s/.exec(t); return m ? +m[1] : null; };
  const refDoi = (t) => { const m = /\b(10\.\d{4,9}\/[^\s"<>]+?)[.,;)\]]*(?:\s|$)/i.exec(t + " "); return m ? m[1].toLowerCase() : null; };
  const refBody = (t) => String(t).replace(/^\s*(?:\[\d+\]|\d+\.)\s*/, "");
  const refNorm = (t) => refBody(t).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/https?:\/\/(dx\.)?doi\.org\//g, "doi ").replace(/\bdoi:\s*/g, "doi ").replace(/[^a-z0-9]+/g, "");
  const refYear = (t) => { const m = /\((\d{4}[a-z]?)\)|\b(?:19|20)\d{2}[a-z]?\b/.exec(refBody(t)); return m ? m[1] || m[0] : null; }; // 2020a ≠ 2020b
  const refAuthor = (t) => { const b = refBody(t).normalize("NFD").replace(/[̀-ͯ]/g, ""); const m = /^(?:[A-Z]\.\s*(?:[A-Z]\.\s*)*)?([A-Za-z'’\-]+)/.exec(b); return m ? m[1].toLowerCase() : ""; };
  const STOP_REF = new Set(["the", "and", "for", "with", "from", "doi", "https", "org", "vol", "pp"]);
  const refTokens = (t) => new Set((refBody(t).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").match(/[a-z]{3,}/g) || []).filter((w) => !STOP_REF.has(w)));
  const jaccard = (a, b) => { let n = 0; a.forEach((x) => { if (b.has(x)) n++; }); return n / Math.max(1, a.size + b.size - n); };
  /**
   * Cari referensi ganda. texts = teks tiap paragraf referensi (urut).
   * Ganda bila: teks sama (tanpa nomor/format), DOI sama, atau penulis pertama + tahun sama dan isi >= 85% mirip.
   * Yang dipertahankan: yang lebih lengkap (ber-DOI / lebih panjang) — kecuali gaya bernomor: selalu yang pertama.
   * return { remove: [i], keep: {i: j}, numbered }
   */
  function findDuplicateRefs(texts) {
    const n = texts.length;
    const filled = texts.filter((t) => String(t).trim());
    const numbered = filled.length > 0 && filled.every((t) => refNum(t) != null);
    const info = texts.map((t) => ({ t: String(t), norm: refNorm(t), doi: refDoi(t), year: refYear(t), author: refAuthor(t), tok: refTokens(t) }));
    const remove = new Set(), keep = {};
    for (let i = 0; i < n; i++) {
      if (remove.has(i) || info[i].norm.length < 20) continue;
      for (let j = i + 1; j < n; j++) {
        if (remove.has(j) || info[j].norm.length < 20) continue;
        const a = info[i], b = info[j];
        const same = a.norm === b.norm || (a.doi && a.doi === b.doi) ||
          (a.year && a.year === b.year && a.author && a.author === b.author && jaccard(a.tok, b.tok) >= 0.85);
        if (!same) continue;
        let k = i, r = j;
        if (!numbered) { const score = (x) => (x.doi ? 1000 : 0) + x.t.length; if (score(b) > score(a)) { k = j; r = i; } }
        remove.add(r); keep[r] = k;
        if (r === i) break;
      }
    }
    Object.keys(keep).forEach((r) => { let k = keep[r]; while (keep[k] != null) k = keep[k]; keep[r] = k; });
    return { remove: [...remove].sort((a, b) => a - b), keep, numbered };
  }
  /** Peta nomor lama -> nomor baru (IEEE) setelah referensi ganda dihapus. removed: {nomorLama: nomorLamaYangDipertahankan} */
  function renumberMap(oldNums, removed) {
    const map = new Map();
    oldNums.forEach((o, i) => map.set(o, i + 1));
    Object.keys(removed).forEach((o) => map.set(+o, map.get(removed[o])));
    return map;
  }
  // satu kelompok sitasi bernomor: [12] · [3, 12] · [3–5] · [3]–[5] · [1], [3], [7]
  const RE_CITE_GROUP = /\[\s*\d+(?:\s*[,–\-]\s*\d+)*\s*\](?:\s*(?:[–\-]|,)\s*\[\s*\d+(?:\s*[,–\-]\s*\d+)*\s*\])*/g;
  /** Tulis ulang satu kelompok sitasi memakai peta nomor baru (nomor ganda digabung, urut, rentang dipertahankan). */
  function renumberCitation(str, map) {
    const nums = [];
    const pushRange = (a, b) => { for (let x = a; x <= b && x - a < 300; x++) nums.push(x); };
    const brackets = str.match(/\[[^\]]*\]/g) || [];
    const seps = str.split(/\[[^\]]*\]/).slice(1, -1).map((x) => x.trim());
    brackets.forEach((br, k) => {
      if (k > 0 && /^[–\-]$/.test(seps[k - 1]) && nums.length) { const prev = nums[nums.length - 1]; const m = /\d+/.exec(br); nums.pop(); pushRange(prev, +m[0]); return; }
      br.replace(/[\[\]\s]/g, "").split(",").forEach((part) => { const m = /^(\d+)[–\-](\d+)$/.exec(part); if (m) pushRange(+m[1], +m[2]); else if (/^\d+$/.test(part)) nums.push(+part); });
    });
    const mapped = [...new Set(nums.map((x) => (map.has(x) ? map.get(x) : x)))].sort((a, b) => a - b);
    const ranges = [];
    mapped.forEach((x) => { const l = ranges[ranges.length - 1]; if (l && x === l[1] + 1) l[1] = x; else ranges.push([x, x]); });
    const dash = /-/.test(str) && !/–/.test(str) ? "-" : "–";
    const one = ([a, b]) => (a === b ? [String(a)] : b === a + 1 ? [String(a), String(b)] : [`${a}${dash}${b}`]);
    if (brackets.length > 1) // gaya "[1], [3]" / "[3]–[5]": tiap nomor dalam kurung sendiri
      return ranges.map(([a, b]) => (a === b ? `[${a}]` : b === a + 1 ? `[${a}], [${b}]` : `[${a}]${dash}[${b}]`)).join(", ");
    return "[" + ranges.flatMap(one).join(", ") + "]";
  }

  function classifyReferences(blocks, sec, profile) {
    const det = profile.detect, rc = profile.references;
    const myNames = new Set(sectionNames(sec));
    let refs = splitRefBlocks(blocks, det);
    const items = [];
    for (const r of refs) {
      let runs = normalizeRuns(r.runs);
      const t = plain(runs).trim();
      if (!t) continue;
      if (myNames.has(normHead(t)) && countWords(t) <= 4) { items.push({ type: "skip", runs, why: "judul bagian" }); continue; }
      if (rc.stripNumbering) {
        const m = /^(?:\[\d+\]|\d+\.(?!\d))\s*/.exec(t);
        if (m) runs = dropPrefix(runs, m[0].length);
      }
      // normalisasi doi: -> https://doi.org/
      if (rc.doiFormat !== "keep") runs = mapRunsText(runs, (s) => s.replace(/\bdoi:\s*(10\.)/gi, "https://doi.org/$1"));
      if (rc.autoItalic) runs = rc.citeStyle === "ieee" ? ieeeAutoItalic(runs) : apaAutoItalic(runs);
      if (rc.linkDoi) runs = linkUrls(runs);
      items.push({ type: "ref", runs: normalizeRuns(runs) });
    }
    if (rc.sort) {
      const key = (it) => plain(it.runs).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
      const skips = items.filter((i) => i.type !== "ref");
      const rs = items.filter((i) => i.type === "ref").sort((a, b) => key(a).localeCompare(key(b)));
      return skips.concat(rs);
    }
    return items;
  }

  /* ------------------------------------------------------------------ */
  /*  Format akhir per item (dipakai saat menyisipkan ke Word)           */
  /* ------------------------------------------------------------------ */
  const CM = 28.3464567;

  /** Format paragraf (satuan pt) untuk satu item. */
  function paraFormat(it, profile) {
    const P = profile;
    let c, first = 0, left = 0;
    switch (it.type) {
      case "section": c = it.sec && it.sec.heading ? Object.assign({}, P.sectionHeading, it.sec.heading) : P.sectionHeading; break;
      case "heading": c = it.level === 3 ? P.h3 : P.h2; break;
      case "caption": {
        c = P.caption;
        const isFig = /^(Figure|Fig\.|Gambar|Grafik|Chart|Diagram)\b/i.test(plain(it.runs || []).trim());
        if (isFig && c.figureSize) c = Object.assign({}, c, { size: c.figureSize });
        break;
      }
      case "note": c = P.note; break;
      case "abstract": c = Object.assign({}, P.abstract || DEFAULT_PROFILE.abstract, { bold: false }); break;
      case "quote": case "quoteAttr": c = P.quote || P.body; left = (c.leftCm || 0) * CM; break;
      case "equation": c = Object.assign({}, P.body, { align: (P.equation && P.equation.align) || "Centered" }); break;
      case "ref": c = P.references; left = (c.hangingCm || 0) * CM; first = -left; break;
      case "blank": c = Object.assign({}, P.body, { before: 0, after: 0 }); break;
      default: c = P.body; first = (c.firstIndentCm || 0) * CM;
    }
    return {
      style: c.style || "Normal",
      size: c.size || P.body.size,
      align: c.align || "Justified",
      before: +c.before || 0,
      after: +c.after || 0,
      line: +c.line || 1,
      firstIndent: Math.round(first * 100) / 100,
      leftIndent: Math.round(left * 100) / 100,
      rightIndent: it.type === "quote" || it.type === "quoteAttr" ? Math.round(((c.rightCm || 0) * CM) * 100) / 100 : 0,
      bold: !!c.bold,
      italic: it.type === "quoteAttr" ? false : !!c.italic,
    };
  }

  /** Runs final (format karakter) untuk satu item. */
  function finalRuns(it, profile) {
    const det = profile.detect;
    let runs = (it.runs || []).map((r) => Object.assign({}, r));
    if (it.type === "section") {
      const t = plain(runs);
      return [{ text: profile.sectionHeading.caps ? t.toUpperCase() : t, b: !!profile.sectionHeading.bold, i: !!profile.sectionHeading.italic }];
    }
    if (it.type === "heading") {
      const c = it.level === 3 ? profile.h3 : profile.h2;
      return runs.map((r) => ({ text: r.text, b: !!c.bold, i: !!c.italic || (det.keepItalic && !!r.i), sup: det.keepSupSub && r.sup, sub: det.keepSupSub && r.sub }));
    }
    runs = runs.map((r) => ({
      text: r.text,
      b: det.keepBold && !!r.b,
      i: det.keepItalic && !!r.i,
      sup: det.keepSupSub && !!r.sup,
      sub: det.keepSupSub && !!r.sub,
      link: r.link || "",
    }));
    // Paragraf isi yang seluruhnya bold (mis. dari style template) -> bukan bold
    if (it.type === "body" && runs.length && runs.every((r) => r.b) && countWords(plain(runs)) > (det.maxHeadingWords || 15)) runs.forEach((r) => (r.b = false));
    if (it.type === "caption") {
      const t = plain(runs);
      const m = RE_CAPTION.exec(t);
      runs.forEach((r) => (r.b = false));
      if (m && profile.caption.labelBold) runs = patchRange(runs, 0, m[0].length, { b: true });
      // Judul caption yang (hampir) seluruhnya italic — mis. gaya APA atau style "Caption" Word — dinormalkan
      if (profile.caption.stripTitleItalic !== false) {
        const start = m ? m[0].length : 0;
        let all = 0, ital = 0, pos = 0;
        runs.forEach((r) => { const a = pos, b = pos + r.text.length; pos = b; const n = Math.max(0, b - Math.max(a, start)); all += n; if (r.i) ital += n; });
        if (m) runs = patchRange(runs, 0, m[0].length, { i: false });
        if (all && ital / all >= 0.5) runs.forEach((r) => (r.i = false));
      }
    }
    if (it.type === "note") {
      const m = RE_NOTE.exec(plain(runs));
      if (profile.note.italic) runs.forEach((r) => (r.i = true)); // seluruh catatan miring
      if (m && profile.note.labelItalic) runs = patchRange(runs, 0, m[0].length, { i: true });
      if (m && profile.note.labelBold) runs = patchRange(runs, 0, m[0].replace(/\s+$/, "").length, { b: true });
    }
    if ((it.type === "quote" || it.type === "quoteAttr") && profile.quote) {
      if (profile.quote.italic) runs.forEach((r) => (r.i = true));
      else {
        // kutipan yang seluruhnya italic dari naskah sumber -> tegak (APA); italic sebagian (istilah asing) dipertahankan
        let all = 0, ital = 0;
        runs.forEach((r) => { all += r.text.length; if (r.i) ital += r.text.length; });
        if (all && ital / all >= 0.6) runs.forEach((r) => (r.i = false));
      }
    }
    if (it.type === "ref") runs.forEach((r) => (r.b = false));
    return normalizeRuns(runs);
  }

  /**
   * Siapkan items sebelum disisipkan: kutipan diformat sesuai APA 7.
   *  - atribusi di paragraf tersendiri digabung ke kutipannya: "… ." (Peran, kode)
   *  - atribusi yang menempel di belakang tanda kutip penutup ("…” Teacher, S1") -> (Teacher, S1)
   *  - tanda kutip pembuka/penutup dibuang bila kutipan >= minBlockWords (blok)
   */
  function cleanAttribution(t) {
    let a = clean0(t).replace(/^[—–\-,;:\s]+/, "").replace(/[.;,\s]+$/, "");
    if (/^\(.*\)$/.test(a) && a.indexOf(")") === a.length - 1) a = a.slice(1, -1).trim();
    a = a.replace(/\s*\(([^()]*)\)/g, (m, x) => ", " + x.trim()).replace(/^,\s*/, "").replace(/,\s*,/g, ",");
    return a;
  }
  /**
   * Rumus -> it.eq = { omml?, ast, number } untuk disisipkan sebagai persamaan Word.
   * Rumus dari teks biasa ("Y = β0 + β1X1 + ε (1)") ikut diubah bila equation.autoConvert aktif.
   */
  function prepareEquation(it, profile) {
    const Q = profile.equation || {};
    const T = TM();
    if (!T || Q.autoConvert === false && !it.math) return it;
    if (it.math) return Object.assign({}, it, { eq: { omml: it.math.omml, ast: it.math.ast, number: it.math.number || null } });
    const text = plain(it.runs || []).trim();
    if (!text || RE_EQNUM.test(text)) return it; // nomor saja -> biarkan
    // subskrip/superskrip dari format tempelan (β<sub>1</sub>) ikut dipakai
    const runs = it.runs || [];
    const mixed = runs.some((x) => x.sub || x.sup)
      ? runs.map((x) => (x.sub ? "_{" + x.text.trim() + "}" : x.sup ? "^{" + x.text.trim() + "}" : x.text)).join("").trim()
      : text;
    // "…+εit, whereas:" -> kata penghubung di akhir tetap ditulis sebagai teks biasa
    const conn = /,?\s*\b(whereas|where|in which|with|dengan|di mana|dimana|yaitu)\s*:?\s*$/i.exec(mixed);
    const core = conn ? mixed.slice(0, conn.index) : mixed;
    const r = T.parse(core, { forceLinear: mixed !== text });
    if (r.error || !r.ast) return it;
    if (conn) r.ast.items.push({ t: "text", v: (/^,/.test(conn[0].trim()) ? ", " : " ") + conn[1] + (/:\s*$/.test(conn[0]) ? ":" : "") });
    return Object.assign({}, it, { eq: { ast: r.ast, number: r.number } });
  }
  /**
   * Runs (teks + sub/superskrip) dari OOXML satu paragraf Word — untuk "Rapikan" rumus yang masih berupa teks.
   * Tanpa DOMParser (dipakai juga di uji Node).
   */
  function runsFromOoxml(xml) {
    let x = String(xml || "");
    const b = x.indexOf("<w:body>"), e = x.lastIndexOf("</w:body>");
    if (b >= 0 && e > b) x = x.slice(b, e);
    const dec = (t) => t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
    const runs = [];
    const re = /<w:r[\s>][\s\S]*?<\/w:r>/g;
    let m;
    while ((m = re.exec(x))) {
      const r = m[0];
      const va = /<w:vertAlign w:val="(superscript|subscript)"/.exec(r);
      let text = "";
      r.replace(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g, (mm, t) => { text += t != null ? dec(t) : "\t"; return mm; });
      if (!text) continue;
      runs.push({ text, b: /<w:b\/>|<w:b w:val="(1|true|on)"/.test(r), i: /<w:i\/>|<w:i w:val="(1|true|on)"/.test(r), sup: !!va && va[1] === "superscript", sub: !!va && va[1] === "subscript" });
    }
    return runs;
  }
  /** Isi paragraf yang bukan teks biasa: rumus Word, objek (MathType/OLE), gambar. */
  function ooxmlKinds(xml) {
    const x = String(xml || "");
    return {
      omml: /<m:oMath[\s>]/.test(x),
      ole: /<w:object[\s>]|<o:OLEObject\b/.test(x),
      mathType: /ProgID="Equation\.(DSMT\d*|3)"/i.test(x),
      picture: /<w:drawing[\s>]|<w:pict[\s>]/.test(x),
    };
  }
  /** Opsi buildEquationParagraph dari profil. */
  function equationOptions(profile, extra) {
    const Q = profile.equation || {};
    return Object.assign({
      numbering: Q.numbering || "tab", textWidthCm: +Q.textWidthCm || 16, align: Q.align || "Centered",
      font: profile.font, color: profile.color, size: +Q.size || (profile.body && profile.body.size) || 11,
      before: Q.before != null ? +Q.before : 6, after: Q.after != null ? +Q.after : 6, line: 1,
    }, extra || {});
  }
  function prepareItems(items, profile) {
    const Q = Object.assign({ minBlockWords: 40, removeMarks: true, attribution: "paren", mergeAttribution: true }, profile.quote || {});
    const out = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it.type === "equation") { out.push(prepareEquation(it, profile)); continue; }
      if (it.type !== "quote") {
        if (it.type === "quoteAttr" && !Q.mergeAttribution) out.push(it);
        else if (it.type !== "quoteAttr") out.push(it);
        continue;
      }
      let runs = (it.runs || []).map((r) => Object.assign({}, r));
      let text = plain(runs);
      // atribusi menempel setelah tanda kutip penutup
      let attr = "";
      const close = Math.max(text.lastIndexOf("”"), text.lastIndexOf('"'), text.lastIndexOf("»"));
      if (Q.attribution === "paren" && close > 0 && text.slice(close + 1).trim()) {
        attr = text.slice(close + 1);
        runs = sliceRuns(runs, 0, close + 1); text = plain(runs);
      }
      const nx = items[i + 1];
      if (Q.mergeAttribution && nx && nx.type === "quoteAttr") { attr = attr ? attr + "; " + plain(nx.runs) : plain(nx.runs); i++; }
      runs = normalizeRuns(runs); text = plain(runs);
      const open = /^\s*[“"«‘]/.exec(text);
      const endQ = /[”"»’]\s*$/.exec(text);
      const body = text.replace(/^\s*[“"«‘]/, "").replace(/[”"»’]\s*$/, "");
      const isBlock = countWords(body) >= (Q.minBlockWords || 40);
      if (Q.removeMarks && isBlock && open && endQ) {
        runs = sliceRuns(runs, open.index + open[0].length, endQ.index);
        runs = normalizeRuns(runs);
      }
      if (attr && Q.attribution === "paren") {
        const a = cleanAttribution(attr);
        if (a) runs = runs.concat([{ text: " (" + a + ")", i: false, b: false }]);
      } else if (attr) runs = runs.concat([{ text: " " + clean0(attr) }]);
      out.push(Object.assign({}, it, { runs: normalizeRuns(runs), block: isBlock }));
    }
    return out;
  }

  /* ------------------------------------------------------------------ */
  /*  Analisis struktur dokumen (berdasarkan teks paragraf)              */
  /* ------------------------------------------------------------------ */
  /**
   * infos: [{text, inTable}]
   * return: [{index, key|null, title}]
   */
  function findSectionHeadings(infos, profile) {
    const det = profile.detect;
    const lookup = new Map();
    for (const s of profile.sections) if (s.kind !== "abstract") for (const n of sectionNames(s)) lookup.set(n, s.key);
    const out = [];
    infos.forEach((p, i) => {
      if (p.inTable) return;
      const t = (p.text || "").replace(WS, " ").trim();
      if (!t || countWords(t) > 8) return;
      const letters = t.replace(/[^\p{L}]/gu, "");
      const isCaps = letters.length >= 3 && letters === letters.toUpperCase();
      const key = lookup.get(normHead(t));
      if (key && (!det.requireCapsSectionTitle || isCaps)) out.push({ index: i, key, title: t });
      else if (isCaps && letters.length >= 4 && !/[.,;]$/.test(t) && !/\d/.test(t)) out.push({ index: i, key: null, title: t });
    });
    return out;
  }

  /** Rentang [start, end) paragraf milik bagian `key` (start = indeks judul). */
  function sectionRange(infos, heads, key) {
    const k = heads.findIndex((h) => h.key === key);
    if (k < 0) return null;
    const next = heads[k + 1];
    return { start: heads[k].index, end: next ? next.index : infos.length, nextIndex: next ? next.index : -1 };
  }

  /**
   * Tentukan titik sisip di akhir bagian.
   * infos boleh punya hasPicture. Return {index, loc:'After'|'Before', trailingBlanks}
   */
  function pickAnchor(infos, range) {
    const { start, end } = range;
    for (let j = end - 1; j > start; j--) {
      const p = infos[j];
      if (p.inTable) {
        // akhir tabel: paragraf setelah tabel
        const a = j + 1;
        if (a >= end) return { index: end < infos.length ? end : j, loc: end < infos.length ? "Before" : "After", trailingBlanks: 0, afterTable: true };
        return { index: a, loc: "After", trailingBlanks: Math.max(0, end - a - 1), afterTable: true };
      }
      if ((p.text || "").replace(WS, "").length || p.hasPicture) {
        let blanks = 0;
        for (let q = j + 1; q < end; q++) if (!(infos[q].text || "").replace(WS, "").length) blanks++;
        return { index: j, loc: "After", trailingBlanks: blanks };
      }
    }
    return { index: start, loc: "After", trailingBlanks: Math.max(0, end - start - 1) };
  }

  /** Di mana judul bagian yang belum ada harus dibuat. */
  function planNewSection(heads, key, profile, infosLen) {
    const order = profile.sections.map((s) => s.key);
    const myPos = order.indexOf(key);
    // judul bagian berikutnya (menurut urutan profil) yang sudah ada di dokumen
    for (const h of heads) {
      if (!h.key) continue;
      const pos = order.indexOf(h.key);
      if (pos > myPos) return { index: h.index, loc: "Before" };
    }
    return { index: infosLen - 1, loc: "End" };
  }

  /** Hitung kata per bagian. */
  function countDocument(infos, heads, profile) {
    const L = profile.limits;
    const res = { sections: [], total: 0, abstract: 0, limit: +L.bodyWords || 0, minLimit: +L.minBodyWords || 0, abstractLimit: +L.abstractWords || 0 };
    for (let k = 0; k < heads.length; k++) {
      const h = heads[k];
      const end = heads[k + 1] ? heads[k + 1].index : infos.length;
      let w = 0;
      for (let j = h.index + 1; j < end; j++) {
        if (infos[j].inTable && !L.countTables) continue;
        w += countWords(infos[j].text);
      }
      const sec = profile.sections.find((s) => s.key === h.key);
      const isRef = sec && sec.kind === "references";
      const excluded = (L.excludeKeys || []).includes(h.key);
      const counted = h.key != null && !excluded && (!isRef || L.countReferences);
      res.sections.push({ key: h.key, title: h.title, words: w, counted });
      if (counted) res.total += w;
    }
    // abstrak: paragraf panjang setelah "ABSTRACT" (melewati Keywords / Article history),
    // berhenti di judul bagian pertama atau saat keluar dari tabel abstrak
    const ai = infos.findIndex((p) => /^(abstract|abstrak)\s*[:.]?$/i.test((p.text || "").trim()));
    if (ai >= 0) {
      for (let j = ai + 1; j < infos.length; j++) {
        const t = (infos[j].text || "").trim();
        if (heads.some((h) => h.index === j)) break;
        if (!infos[j].inTable && infos[ai].inTable && t) break;
        if (/^(keywords?|kata kunci|article (history|info)|riwayat artikel|received|revised|accepted)\b/i.test(t)) continue;
        if (/^(to cite this article|how to cite|cite this article|cara mengutip)|open access|cc[ -]by|creative commons|correspond/i.test(t)) continue;
        if (countWords(t) >= 20) res.abstract += countWords(t);
      }
    }
    return res;
  }

  /** Klasifikasi ulang paragraf yang sudah ada di dokumen (untuk "Rapikan"). */
  function classifyExisting(paras, sec, profile) {
    const blocks = paras.map((p) => ({ kind: "p", tag: "p", runs: [{ text: p.text || "", b: p.bold === true, i: p.italic === true }] }));
    const det = Object.assign({}, profile.detect, { mergeBrokenLines: false });
    const prof = Object.assign({}, profile, { detect: det, references: Object.assign({}, profile.references, { sort: false }) });
    if (sec.kind === "references") return blocks.map((b) => ({ type: plain(b.runs).trim() ? "ref" : "skip" }));
    const items = [];
    blocks.forEach((b, i) => {
      if (!plain(b.runs).trim()) { items[i] = { type: "skip" }; return; }
      const one = classify([b].concat(blocks[i + 1] ? [blocks[i + 1]] : []), {}, prof)[0] || { type: "body" };
      items[i] = one;
    });
    // atribusi kutipan di paragraf tersendiri
    let prev = -1;
    items.forEach((it, i) => {
      if (!it || it.type === "skip") return;
      if (prev >= 0 && items[prev].type === "quote" && it.type !== "quote" && looksLikeAttribution(paras[i].text || "", blocks[i].runs) && paras[i].bold !== true) items[i] = { type: "quoteAttr" };
      else if (looksLikeEquation(paras[i].text || "")) items[i] = { type: "equation" };
      prev = i;
    });
    // level heading untuk R&D
    const majors = new Set((sec.majorSubheadings || []).map(normHead));
    let seen = false;
    items.forEach((it, i) => {
      if (it.type !== "heading") return;
      const t = normHead(paras[i].text);
      if (majors.has(t)) { it.level = 2; seen = true; } else it.level = seen ? 3 : 2;
    });
    return items;
  }


  /* ------------------------------------------------------------------ */
  /*  Abstrak                                                            */
  /* ------------------------------------------------------------------ */
  function sliceRuns(runs, s, e) {
    const out = []; let pos = 0;
    for (const r of runs) {
      const a = pos, b = pos + r.text.length; pos = b;
      if (b <= s || a >= e) continue;
      out.push(Object.assign({}, r, { text: r.text.slice(Math.max(s, a) - a, Math.min(e, b) - a) }));
    }
    return out;
  }
  const reEsc = (x) => x.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");

  /**
   * Pecah tempelan abstrak menjadi komponen [{type:'abstract', label, sep, runs}].
   * Label dikenali di awal paragraf ("Purpose: …", "Findings – …") atau di tengah paragraf setelah akhir kalimat.
   */
  function classifyAbstract(blocks, sec, profile) {
    const A = profile.abstract || DEFAULT_PROFILE.abstract;
    const labels = (A.labels || []).slice().sort((a, b) => b.length - a.length).map(reEsc).join("|");
    const reStart = new RegExp(`^(${labels})\\s*([–—:]|-(?=\\s)|\\.(?=\\s))\\s*`, "iu");
    const reMid = new RegExp(`([.!?)\\]”"])(\\s+)(${labels})\\s*([–—:]|\\s-\\s)\\s*`, "giu");
    const reGeneric = /^([\p{Lu}][\p{L}\/&\- ]{2,40}?)\s+([–—])\s+(?=\S)/u;
    const anyKnown = blocks.some((b) => b.kind === "p" && reStart.test(plain(b.runs).trim()));
    const items = [];
    let inKeywords = false;
    for (const bl of blocks) {
      if (bl.kind === "table") continue;
      let runs = bl.runs;
      const text = plain(runs);
      const t = text.trim();
      if (!t) continue;
      if (/^(abstract|abstrak)\s*[:.]?$/i.test(t)) { items.push({ type: "skip", runs, why: "judul" }); continue; }
      if (/^(keywords?|kata kunci)\b/i.test(t)) { items.push({ type: "skip", runs, why: "keywords" }); inKeywords = true; continue; }
      if (/^(article (info|information|history)|riwayat artikel|info artikel|received|revised|accepted|published|available online|diterima|direvisi|disetujui)\b/i.test(t) ||
          RE_CITE_LINE.test(t) || /open access|cc[ -]by|creative commons|licen[cs]ed? under/i.test(t) || /^(e-?mail|correspond)/i.test(t)) {
        items.push({ type: "skip", runs, why: "info artikel" }); continue;
      }
      // butir keywords ("Digital learning;") setelah baris Keywords
      if (inKeywords && countWords(t) <= 8 && !/[.!?]\s+\p{Lu}/u.test(t)) { items.push({ type: "skip", runs, why: "keywords" }); continue; }
      inKeywords = false;
      // titik potong: awal paragraf + label di tengah
      const cuts = [];
      const m0 = reStart.exec(text) || (anyKnown && countWords(text) > 8 ? reGeneric.exec(text) : null);
      if (m0 && countWords(m0[1]) <= 5) cuts.push({ at: 0, len: m0[0].length, label: m0[1], sep: m0[2] });
      let m;
      reMid.lastIndex = 0;
      while ((m = reMid.exec(text))) {
        const at = m.index + m[1].length + m[2].length;
        cuts.push({ at, len: m.index + m[0].length - at, label: m[3], sep: m[4].trim() });
      }
      if (!cuts.length) { items.push({ type: "abstract", label: null, runs: normalizeRuns(runs) }); continue; }
      if (cuts[0].at > 0) items.push({ type: "abstract", label: null, runs: normalizeRuns(sliceRuns(runs, 0, cuts[0].at)) });
      cuts.forEach((c, k) => {
        const end = k + 1 < cuts.length ? cuts[k + 1].at : text.length;
        const body = normalizeRuns(sliceRuns(runs, c.at + c.len, end));
        items.push({ type: "abstract", label: c.label.trim(), sep: c.sep, runs: body });
      });
    }
    return items.filter((it) => it.type !== "abstract" || plain(it.runs).trim() || it.label);
  }

  /** Paragraf keluaran abstrak: [[run…]] — run {br:true} = line break. */
  function abstractParagraphs(items, profile) {
    const A = profile.abstract || DEFAULT_PROFILE.abstract;
    const det = profile.detect;
    const comps = items.filter((i) => i.type === "abstract");
    const body = (runs) => runs.map((r) => ({ text: r.text, b: false, i: !!A.italic || (det.keepItalic && !!r.i), sup: det.keepSupSub && !!r.sup, sub: det.keepSupSub && !!r.sub }));
    const labelRun = (c) => {
      const sep = A.labelSep ? A.labelSep : (c.sep === ":" || c.sep === "." ? c.sep + " " : " " + (c.sep || "–") + " ");
      return { text: c.label + sep, b: !!A.labelBold, i: !!A.italic };
    };
    const labelRuns = (c) => {
      const r = labelRun(c);
      if (A.sepBold !== false || !r.b) return [r];
      return [{ text: c.label, b: true, i: r.i }, { text: r.text.slice(c.label.length), b: false, i: r.i }];
    };
    const paras = [];
    let cur = null;
    comps.forEach((c) => {
      const runs = (c.label ? labelRuns(c) : []).concat(body(c.runs));
      if (c.label && A.joinLabeled === "linebreak") {
        if (cur) cur.push({ br: true }); else { cur = []; paras.push(cur); }
        cur.push(...runs);
      } else { cur = null; paras.push(runs); }
    });
    return paras.map((p) => p.filter((r) => r.br || r.text));
  }

  /**
   * Paragraf abstrak yang akan diganti.
   * infos[j] = {text, inTable, row, cell}; h = indeks judul ABSTRACT.
   * return {targets:[indeks…], after:null|indeks}
   */
  function pickAbstractTarget(infos, h) {
    const H = infos[h];
    let cand = [];
    if (H.inTable) {
      // batas tabel: urutan (baris, sel) mundur = sudah masuk tabel lain
      let end = h + 1, pr = H.row, pc = H.cell;
      while (end < infos.length && infos[end].inTable) {
        const r = infos[end].row, c = infos[end].cell;
        if (r != null && (r < pr || (r === pr && c < pc))) break;
        if (r != null) { pr = r; pc = c; }
        end++;
      }
      for (let j = h + 1; j < end; j++) if (infos[j].row === H.row && infos[j].cell === H.cell) cand.push(j);
      if (!cand.length) for (let j = h + 1; j < end; j++) if (infos[j].row === H.row + 1 && infos[j].cell === H.cell) cand.push(j);
      if (!cand.length) { const j = infos.findIndex((p, k) => k > h && p.inTable && countWords(p.text) >= 20); if (j > 0) cand.push(j); }
    } else {
      for (let j = h + 1; j < infos.length; j++) {
        const t = (infos[j].text || "").trim();
        if (/^(keywords?|kata kunci)\b/i.test(t) || infos[j].inTable) break;
        const letters = t.replace(/[^\p{L}]/gu, "");
        if (t && letters.length >= 4 && letters === letters.toUpperCase() && countWords(t) <= 8) break;
        cand.push(j);
        if (!t && cand.some((k) => (infos[k].text || "").trim())) { cand.pop(); break; }
      }
    }
    return trimAbstractCandidates(infos, cand, h);
  }

  function trimAbstractCandidates(infos, cand, h) {
    const lic = cand.findIndex((j) => /open access|cc by|creative commons|licen[cs]e/i.test(infos[j].text || ""));
    if (lic >= 0) cand = cand.slice(0, lic);
    const filled = cand.filter((j) => (infos[j].text || "").trim());
    if (!filled.length) return cand.length ? { targets: [cand[0]], after: null } : { targets: [], after: h };
    return { targets: cand.slice(0, cand.indexOf(filled[filled.length - 1]) + 1), after: null };
  }

  /** Target abstrak bila rentang sel abstrak [cellStart, cellEnd) sudah diketahui (tanpa nomor baris/kolom). */
  function abstractTargetsInCell(infos, h, cellStart, cellEnd) {
    const cand = [];
    for (let j = Math.max(cellStart, h + 1); j < cellEnd; j++) cand.push(j);
    const r = trimAbstractCandidates(infos, cand, h);
    r.cellEnd = cellEnd;
    return r;
  }

  /** Paragraf yang jelas BUKAN isi abstrak (keywords, article info/history, cite, lisensi). */
  const isAbstractNoise = (t) => /^(keywords?|kata kunci|article (info|information|history)|riwayat artikel|info artikel|received|revised|accepted|published|to cite|how to cite|cite this|cara mengutip|e-?mail|correspond)/i.test(clean0(t)) || /open access|cc[ -]by|creative commons/i.test(t || "");
  const clean0 = (t) => String(t || "").replace(WS, " ").trim();

  const RE_ABSTRACT_HEAD = /^(abstract|abstrak)\s*[:.]?$/i;
  const RE_CITE_LINE = /^(to cite this article|how to cite|cite this article|cara mengutip)/i;

  return {
    classifyAbstract, abstractParagraphs, pickAbstractTarget, abstractTargetsInCell, isAbstractNoise, RE_ABSTRACT_HEAD,
    DEFAULT_PROFILE, PRESETS, presetFor, mergeProfile, clone,
    parseInput, parseHtml, splitWhole, wholeHeading, wholeHeadingMap, parseText, classify, classifyReferences, classifyExisting,
    findDuplicateRefs, renumberMap, renumberCitation, RE_CITE_GROUP, refNum, ieeeAutoItalic, paraFormat, finalRuns, prepareItems, prepareEquation, equationOptions, findDisplayLatex, inlineMathSpans, isLostLatex, runsFromOoxml, ooxmlKinds, cleanAttribution, looksLikeEquation, looksLikeAttribution, countWords, plain, normalizeRuns, titleCase, apaAutoItalic, linkUrls,
    findSectionHeadings, sectionRange, pickAnchor, planNewSection, countDocument, normHead,
    CM,
  };
});
