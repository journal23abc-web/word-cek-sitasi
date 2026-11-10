/*
 * WordOps — semua operasi ke dokumen Word (Office.js, WordApi 1.3+).
 * Bergantung pada TplEngine (engine.js).
 */
(function (root) {
  "use strict";
  const E = root.TplEngine;
  const BOOKMARK = "_TplPasteLast";
  const LINK_COLOR = "#0563C1";

  const hasDesktop = (v) => { try { return typeof Office !== "undefined" && Office.context.requirements.isSetSupported("WordApiDesktop", v); } catch (e) { return false; } };
  const has = (v) => typeof Office !== "undefined" && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("WordApi", v);

  /* ---------------- gaya ---------------- */
  const BUILTIN = {
    "normal": "Normal", "no spacing": "NoSpacing", "nospacing": "NoSpacing", "caption": "Caption",
    "bibliography": "Bibliography", "quote": "Quote", "list paragraph": "ListParagraph", "title": "Title", "subtitle": "Subtitle",
  };
  function setStyle(p, name) {
    const k = String(name || "Normal").trim().toLowerCase().replace(/\s+/g, " ");
    const m = /^heading ?([1-9])$/.exec(k);
    if (m) p.styleBuiltIn = "Heading" + m[1];
    else if (BUILTIN[k]) p.styleBuiltIn = BUILTIN[k];
    else p.style = name; // nama gaya kustom — harus persis sama dengan di dokumen
  }

  function customStyleNames(profile) {
    const names = [profile.sectionHeading, profile.h2, profile.h3, profile.body, profile.caption, profile.note, profile.references]
      .map((c) => (c && c.style) || "Normal");
    return [...new Set(names.filter((n) => { const k = n.trim().toLowerCase(); return !BUILTIN[k] && !/^heading ?[1-9]$/.test(k); }))];
  }

  /** Ganti gaya kustom yang tidak ada di dokumen dengan Normal (bila API mendukung pengecekan). */
  async function validateStyles(context, profile) {
    const custom = customStyleNames(profile);
    if (!custom.length || !has("1.5")) return { profile, missing: [] };
    const styles = context.document.getStyles();
    styles.load("items/nameLocal");
    await context.sync();
    const exist = new Set(styles.items.map((s) => s.nameLocal.toLowerCase()));
    const missing = custom.filter((n) => !exist.has(n.toLowerCase()));
    if (!missing.length) return { profile, missing };
    const p = E.clone(profile);
    for (const k of ["sectionHeading", "h2", "h3", "body", "caption", "note", "references"]) {
      if (p[k] && missing.includes(p[k].style)) p[k].style = "Normal";
    }
    return { profile: p, missing };
  }

  /* ---------------- format paragraf baru ---------------- */
  function formatNewParagraph(p, it, profile) {
    const f = E.paraFormat(it, profile);
    setStyle(p, f.style);
    p.alignment = f.align;
    p.spaceBefore = f.before;
    p.spaceAfter = f.after;
    p.lineSpacing = 12 * f.line; // Office.js: nilai poin; UI Word membaginya dengan 12
    p.leftIndent = f.leftIndent;
    p.firstLineIndent = f.firstIndent;
    p.rightIndent = f.rightIndent || 0;
    // format tanda paragraf (mempengaruhi tinggi baris terakhir & teks yang diketik setelahnya)
    p.font.set({ name: profile.font, size: f.size, bold: f.bold, italic: f.italic, color: profile.color, underline: "None", superscript: false, subscript: false });
    p.font.highlightColor = null;
    if (it.type === "blank") return;
    for (const r of E.finalRuns(it, profile)) {
      if (!r.text) continue;
      const rg = p.insertText(r.text, "End");
      rg.font.set({
        name: profile.font, size: f.size, bold: !!r.b, italic: !!r.i,
        superscript: !!r.sup, subscript: !!r.sub,
        color: r.link ? LINK_COLOR : profile.color, underline: r.link ? "Single" : "None",
      });
      if (r.link) rg.hyperlink = r.link;
    }
  }

  function rectValues(rows) {
    const w = Math.max(1, ...rows.map((r) => r.length));
    return rows.map((r) => { const a = r.slice(0, w); while (a.length < w) a.push(""); return a.map((c) => String(c)); });
  }

  /** Format tabel APA (dipanggil setelah sync agar paragraf sel bisa dimuat). */
  /**
   * Format tabel sesuai template (dipanggil setelah t.paras & t.rows dimuat).
   * opts.existing = tabel yang sudah ada di dokumen (italic di sel dipertahankan).
   */
  function formatTable(t, profile, opts) {
    opts = opts || {};
    const T = profile.table;
    const line = (b, on) => { if (on) { b.type = "Single"; b.width = 0.5; b.color = "#000000"; } else b.type = "None"; };
    t.paras.items.forEach((p) => {
      p.styleBuiltIn = "Normal";
      p.spaceBefore = 0; p.spaceAfter = 0; p.lineSpacing = 12; p.firstLineIndent = 0; p.leftIndent = 0; p.rightIndent = 0;
    });
    const tbl = t.table;
    const f = { name: profile.font, size: T.size, bold: false, color: profile.color, highlightColor: null };
    if (!opts.existing) f.italic = false;
    tbl.font.set(f);
    const rows = t.rows.items, n = rows.length;
    // judul kolom: 1 baris, atau 2 baris bila baris pertama berisi sel gabungan (sel lebih sedikit dari baris data)
    const cc = rows.map((r) => (typeof r.cellCount === "number" ? r.cellCount : 0));
    const maxC = Math.max(0, ...cc);
    const hRows = n >= 3 && cc[0] && cc[0] < maxC && cc[1] === maxC ? 2 : 1;
    // judul kolom diulang di atas tiap halaman (Repeat as header row)
    if (T.repeatHeader !== false) tbl.headerRowCount = Math.min(hRows, Math.max(1, n - 1));
    if (T.apaBorders) {
      ["Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((loc) => line(tbl.getBorder(loc), false));
      ["Top", "Bottom"].forEach((loc) => line(tbl.getBorder(loc), true));
      // garis per baris — menimpa garis yang dipasang per sel (umum di naskah hasil konversi)
      rows.forEach((row, r) => {
        ["Left", "Right", "InsideVertical"].forEach((loc) => line(row.getBorder(loc), false));
        if (r === 0) line(row.getBorder("Top"), true);
        if (r < hRows) { line(row.getBorder("Bottom"), r === hRows - 1); if (r > 0) line(row.getBorder("Top"), false); }
        else {
          if (r >= hRows + 1) line(row.getBorder("Top"), false);
          line(row.getBorder("Bottom"), r === n - 1);
        }
      });
    }
    rows.slice(0, hRows).forEach((r) => { r.font.bold = !!T.headerBold; });
    tbl.horizontalAlignment = T.cellAlign || "Centered";
    if (T.firstColLeft) rows.forEach((row) => { row.cells.getFirst().horizontalAlignment = "Left"; });
    // AutoFit to contents (Word desktop baru: WordApiDesktop 1.4), lalu tabel di tengah halaman
    let autofit = null;
    if (T.autoFit && T.autoFit !== "none") {
      if (hasDesktop("1.4")) { tbl.autoFitBehavior(T.autoFit === "window" ? "Window" : "Content"); autofit = true; }
      else if (T.autoFit === "window") { tbl.autoFitWindow(); autofit = true; }
      else autofit = false;
    }
    if (T.align) tbl.alignment = T.align;
    return { autofit, headerRows: hRows };
  }

  /**
   * Sisipkan items berurutan mulai dari `start` (Paragraph) di lokasi `loc` ('After'|'Before').
   * return {first, last, count, tables}
   */
  async function insertItems(context, start, loc, items0, profile, opts) {
    opts = opts || {};
    const items = E.prepareItems(items0, profile); // kutipan: format APA 7
    let cursor = start, cursorIsTable = false, detachNext = !!opts.anchorIsList;
    let first = null, last = null, count = 0, sinceSync = 0, equations = 0, eqFallback = 0, autofitUnsupported = false;
    const tables = [];
    for (const it of items) {
      if (it.type === "skip") continue;
      if (it.type === "table") {
        const vals = rectValues(it.rows || []);
        if (!vals.length) continue;
        let host = cursor, hostLoc = loc;
        if (cursorIsTable) { // dua tabel berurutan: beri paragraf pemisah
          host = cursor.insertParagraph("", "After");
          formatNewParagraph(host, { type: "blank" }, profile);
          hostLoc = "After";
          first = first || host; last = host;
        }
        const tbl = host.insertTable(vals.length, vals[0].length, hostLoc, vals);
        tables.push({ table: tbl });
        first = first || tbl; last = tbl; cursor = tbl; cursorIsTable = true; loc = "After"; count++;
      } else if (it.type === "equation" && it.eq && M()) {
        // rumus -> persamaan Word asli (OMML)
        const p = cursor.insertParagraph("", loc);
        if (detachNext) { p.detachFromList(); detachNext = false; }
        if (opts.textWidthCm == null) opts.textWidthCm = await readTextWidthCm(context);
        // coba: OMML asli (dari Word) -> OMML buatan add-in -> teks berformat; satu rumus gagal tidak menghentikan semuanya
        const tries = [];
        if (it.eq.omml) tries.push({ omml: it.eq.omml });
        if (it.eq.ast) tries.push({ ast: it.eq.ast });
        let eqp = null;
        for (const eq of tries) {
          try {
            const pkg = M().buildEquationPackage(eq, E.equationOptions(profile, { number: it.eq.number, textWidthCm: opts.textWidthCm || undefined }));
            eqp = await placeEquation(context, p, pkg);
            break;
          } catch (e) { console.warn("rumus ditolak Word, coba cara lain", e, eq); }
        }
        if (!eqp) {
          const runs = it.eq.ast ? M().toRuns(it.eq.ast) : (it.runs || []);
          const num = it.eq.number ? [{ text: "  (" + it.eq.number + ")" }] : [];
          formatNewParagraph(p, Object.assign({}, it, { type: "equation", runs: runs.concat(num) }), profile);
          eqp = p; eqFallback++;
        }
        first = first || eqp; last = eqp; cursor = eqp; cursorIsTable = false; loc = "After"; count++; equations++;
        sinceSync = 0;
        continue;
      } else {
        const p = cursor.insertParagraph("", loc);
        if (detachNext) { p.detachFromList(); detachNext = false; } // jangan ikut jadi daftar bernomor
        formatNewParagraph(p, it, profile);
        first = first || p; last = p; cursor = p; cursorIsTable = false; loc = "After"; count++;
      }
      if (++sinceSync >= 40) { await context.sync(); sinceSync = 0; }
    }
    await context.sync();
    if (tables.length) {
      for (const t of tables) { t.paras = t.table.getRange("Whole").paragraphs; t.paras.load("items"); t.rows = t.table.rows; t.rows.load("items/cellCount"); }
      await context.sync();
      for (const t of tables) { const fr = formatTable(t, profile); if (fr && fr.autofit === false) autofitUnsupported = true; }
      await context.sync();
    }
    return { first, last, count, equations, eqFallback, autofitUnsupported, endsWithTable: cursorIsTable };
  }

  /* ---------------- rumus (OMML) ---------------- */
  const M = () => root.TplMath || null;
  let _widthCache = null;
  /** Lebar area teks (cm) dari pgSz/pgMar di OOXML paragraf terakhir; null bila tidak terbaca. */
  async function readTextWidthCm(context) {
    if (_widthCache) return _widthCache;
    try {
      const last = context.document.body.paragraphs.getLast();
      const ox = last.getRange("Whole").getOoxml();
      await context.sync();
      const sz = /<w:pgSz\b[^>]*\bw:w="(\d+)"/.exec(ox.value), mar = /<w:pgMar\b[^>]*>/.exec(ox.value);
      if (!sz || !mar) return null;
      const l = /w:left="(-?\d+)"/.exec(mar[0]), r = /w:right="(-?\d+)"/.exec(mar[0]), g = /w:gutter="(\d+)"/.exec(mar[0]);
      const tw = +sz[1] - (l ? +l[1] : 0) - (r ? +r[1] : 0) - (g ? +g[1] : 0);
      if (!(tw > 2000)) return null;
      _widthCache = Math.round((tw / 567) * 100) / 100;
      return _widthCache;
    } catch (e) { return null; }
  }
  const hasMathXml = (x) => /<m:oMath[\s>]/.test(x || "");
  /**
   * Ganti paragraf kosong `p` dengan paragraf persamaan dari paket OOXML.
   * Word kadang menambah satu paragraf kosong setelah sisipan OOXML — dibersihkan di sini.
   * return paragraf persamaan.
   */
  async function placeEquation(context, p, pkg) {
    const next = p.getNextOrNullObject();
    next.load("isNullObject");
    await context.sync();
    p.insertOoxml(pkg, "Replace");
    await context.sync();
    const lastP = next.isNullObject ? context.document.body.paragraphs.getLast() : next.getPreviousOrNullObject();
    lastP.load("text,isNullObject");
    await context.sync();
    if (lastP.isNullObject) return p;
    if ((lastP.text || "").trim()) return lastP;
    const prevP = lastP.getPreviousOrNullObject();
    prevP.load("text,isNullObject");
    const oxLast = lastP.getRange("Whole").getOoxml();
    await context.sync();
    if (hasMathXml(oxLast.value)) return lastP;
    if (prevP.isNullObject) return lastP;
    const oxPrev = prevP.getRange("Whole").getOoxml();
    await context.sync();
    if (hasMathXml(oxPrev.value)) { lastP.delete(); await context.sync(); return prevP; }
    return lastP;
  }

  /* ---------------- membaca struktur dokumen ---------------- */
  async function readStructure(context, profile, withBold) {
    const paras = context.document.body.paragraphs;
    paras.load(withBold ? "items/text,items/tableNestingLevel,items/font/bold,items/font/italic" : "items/text,items/tableNestingLevel");
    await context.sync();
    const infos = paras.items.map((p) => ({ text: p.text, inTable: p.tableNestingLevel > 0, level: p.tableNestingLevel, bold: withBold ? p.font.bold : undefined, italic: withBold ? p.font.italic : undefined }));
    const heads = E.findSectionHeadings(infos, profile);
    return { paras: paras.items, infos, heads };
  }

  async function markPictures(context, paras, infos, range) {
    const probe = [];
    for (let j = range.end - 1; j > range.start && probe.length < 12; j--) {
      if (infos[j].inTable) break;
      if ((infos[j].text || "").trim()) break;
      const c = paras[j].inlinePictures; c.load("items"); probe.push([j, c]);
    }
    if (!probe.length) return;
    await context.sync();
    probe.forEach(([j, c]) => { infos[j].hasPicture = c.items.length > 0; });
  }

  function bookmarkAndSelect(context, first, last) {
    if (!first) return;
    const rng = first.getRange("Start").expandTo(last.getRange("End"));
    if (has("1.4")) {
      try { context.document.deleteBookmark(BOOKMARK); } catch (e) { /* abaikan */ }
      rng.insertBookmark(BOOKMARK);
    }
    first.getRange("Start").select("Start");
  }

  /* ---------------- API publik ---------------- */

  /** Status bagian di dokumen: {key: true/false} + headings. */
  async function scan(profile) {
    return Word.run(async (context) => {
      const { infos, heads } = await readStructure(context, profile);
      const found = {};
      heads.forEach((h) => { if (h.key) found[h.key] = true; });
      if (infos.some((p) => E.RE_ABSTRACT_HEAD.test((p.text || "").trim()))) found.abstract = true;
      return { found, count: E.countDocument(infos, heads, profile) };
    });
  }

  /**
   * Sisipkan items ke bagian `key`.
   * mode 'section': di akhir bagian (buat judul bagian bila belum ada)
   * mode 'cursor' : di posisi kursor
   */
  async function insertSection(key, items, profile0, mode) {
    return Word.run(async (context) => {
      const { profile, missing } = await validateStyles(context, profile0);
      const sec = profile.sections.find((s) => s.key === key);
      let start, loc, created = false, needTrailingBlank = false;
      let preFirst = null;

      if (mode === "cursor") {
        const sel = context.document.getSelection();
        const p = sel.paragraphs.getLast();
        p.load("text,tableNestingLevel");
        await context.sync();
        if (p.tableNestingLevel > 0) throw new Error("Kursor berada di dalam tabel. Letakkan kursor di luar tabel.");
        start = p; loc = p.text.trim() ? "After" : "Before";
      } else {
        const { paras, infos, heads } = await readStructure(context, profile);
        const range = E.sectionRange(infos, heads, key);
        if (range) {
          await markPictures(context, paras, infos, range);
          const a = E.pickAnchor(infos, range);
          start = paras[a.index]; loc = a.loc;
          needTrailingBlank = profile.layout.blankBeforeSection && range.nextIndex >= 0 && a.trailingBlanks === 0 && a.loc === "After";
          if (a.afterTable && a.loc === "Before") needTrailingBlank = profile.layout.blankBeforeSection;
        } else {
          // judul bagian belum ada -> buat
          const plan = E.planNewSection(heads, key, profile, paras.length);
          let heading;
          if (plan.loc === "Before") {
            const nextHead = paras[plan.index];
            heading = nextHead.insertParagraph("", "Before");
            const prevInfo = infos[plan.index - 1];
            if (prevInfo && (prevInfo.text || "").trim() && !prevInfo.inTable) {
              const blank = heading.insertParagraph("", "Before");
              formatNewParagraph(blank, { type: "blank" }, profile);
              preFirst = blank;
            }
            needTrailingBlank = profile.layout.blankBeforeSection;
          } else {
            const lastP = paras[paras.length - 1];
            const lastEmpty = !(infos[infos.length - 1].text || "").trim();
            if (lastEmpty) {
              heading = lastP.insertParagraph("", "Before");
              const prev = infos[infos.length - 2];
              if (prev && (prev.text || "").trim()) {
                const blank = heading.insertParagraph("", "Before");
                formatNewParagraph(blank, { type: "blank" }, profile);
                preFirst = blank;
              }
            } else {
              const blank = lastP.insertParagraph("", "After");
              formatNewParagraph(blank, { type: "blank" }, profile);
              preFirst = blank;
              heading = blank.insertParagraph("", "After");
            }
          }
          formatNewParagraph(heading, { type: "section", runs: [{ text: sec.title }], sec }, profile);
          preFirst = preFirst || heading;
          start = heading; loc = "After"; created = true;
        }
      }

      let anchorIsList = false;
      try { start.load("isListItem"); await context.sync(); anchorIsList = !!start.isListItem; } catch (e) { anchorIsList = false; }
      const res = await insertItems(context, start, loc, items, profile, { anchorIsList });
      let last = res.last;
      if (res.count && needTrailingBlank) {
        const blank = last.insertParagraph("", "After");
        formatNewParagraph(blank, { type: "blank" }, profile);
        last = blank;
      }
      const first = preFirst || res.first;
      bookmarkAndSelect(context, first, last || first);
      await context.sync();

      // hitung ulang kata
      const st = await readStructure(context, profile);
      return { inserted: res.count, created, missingStyles: missing, eqFallback: res.eqFallback || 0, autofitUnsupported: !!res.autofitUnsupported, count: E.countDocument(st.infos, st.heads, profile) };
    });
  }

  /** Terapkan format template ke paragraf yang SUDAH ada di bagian `key`. */
  async function reformatSection(key, profile) {
    return Word.run(async (context) => {
      const sec = profile.sections.find((s) => s.key === key);
      const { paras, infos, heads } = await readStructure(context, profile, true);
      const range = E.sectionRange(infos, heads, key);
      if (!range) throw new Error(`Judul "${sec.title}" tidak ditemukan di dokumen.`);
      const idx = [];
      for (let j = range.start + 1; j < range.end; j++) if (!infos[j].inTable) idx.push(j);
      const items = E.classifyExisting(idx.map((j) => ({ text: infos[j].text, bold: infos[j].bold, italic: infos[j].italic })), sec, profile);

      // judul bagian
      const hp = paras[range.start];
      applyExisting(hp, { type: "section", sec }, profile, infos[range.start].bold);
      if (profile.sectionHeading.caps) {
        const t = infos[range.start].text.trim();
        if (t !== t.toUpperCase()) hp.insertText(t.toUpperCase(), "Replace");
      }
      let n = 0, quotesFixed = 0;
      const Q = Object.assign({ minBlockWords: 40, removeMarks: true, attribution: "paren", mergeAttribution: true }, profile.quote || {});
      let lastQuote = -1;
      idx.forEach((j, k) => {
        const it = items[k];
        if (!it || it.type === "skip") return;
        if (it.type === "equation") return; // persamaan tidak disentuh
        if (it.type === "quoteAttr" && Q.mergeAttribution && lastQuote >= 0 && Q.attribution === "paren") {
          // gabungkan atribusi ke akhir kutipan: "… ." (Peran, kode) — lalu hapus paragraf atribusinya
          const a = E.cleanAttribution(infos[j].text || "");
          if (a) {
            const r = paras[lastQuote].insertText(" (" + a + ")", "End");
            r.font.set({ italic: false, bold: false, name: profile.font, size: E.paraFormat({ type: "quote" }, profile).size });
            paras[j].delete();
            n++;
          }
          lastQuote = -1;
          return;
        }
        const t = (infos[j].text || "").trim();
        // baris persamaan "… (1)" / nomor saja / paragraf gambar: jangan diubah perataannya
        if (/\t/.test(infos[j].text || "") && /\(\d+[a-z]?\)\s*$/.test(t)) return;
        if (t.replace(/[^\p{L}]/gu, "").length < 2) return;
        applyExisting(paras[j], it, profile, infos[j].bold, infos[j].italic);
        if (it.type === "quote") { fixExistingQuote(paras[j], infos[j].text || "", Q); lastQuote = j; quotesFixed++; } else lastQuote = -1;
        if (it.type === "note" && profile.note && (profile.note.labelBold || profile.note.labelItalic)) {
          const m = /^(Notes?|Catatan|Sumber|Source|Keterangan)\s*[.:]/i.exec(t);
          if (m) {
            const hit = paras[j].search(m[0], { matchCase: true }).getFirstOrNullObject();
            if (profile.note.labelBold) hit.font.bold = true;
            if (profile.note.labelItalic) hit.font.italic = true;
          }
        }
        if (it.type === "caption") {
          const m = /^(Table|Tables|Figure|Fig\.|Tabel|Gambar|Grafik|Chart|Diagram)\s+(\d+[A-Za-z]?|[IVX]+)\s*[.:]?/i.exec(t);
          if (m && profile.caption.labelBold) {
            const hit = paras[j].search(m[0], { matchCase: true }).getFirstOrNullObject();
            hit.font.bold = true;
          }
        }
        n++;
      });
      await context.sync();

      // referensi ganda dihapus (IEEE: nomor & sitasi di badan artikel ikut disesuaikan)
      const dup = sec.kind === "references" ? await dedupeReferences(context, paras, infos, idx, range) : { removed: 0 };
      if (sec.kind === "references") return { formatted: n, headings: 0, quotes: 0, tables: 0, tableErrors: 0, equations: 0, dupRemoved: dup.removed, dupList: dup.list || [], renumbered: dup.renumbered || 0, citations: dup.citations || 0 };

      // rumus: baris rumus berupa teks -> persamaan Word; nomor "(1)" di baris terpisah digabung
      const eqRes = await fixEquations(context, paras, infos, idx, items, profile);
      n += eqRes.converted;

      // tabel di bagian ini — jalan dari paragraf tabel pertama, lompati seluruh paragraf tabel itu
      let tablesDone = 0, tableErrors = 0, autofitUnsupported = false;
      for (let j = range.start + 1; j < range.end; ) {
        if (!infos[j].inTable) { j++; continue; }
        let next = j + 1;
        try {
          const tbl = paras[j].parentTableOrNullObject;
          const tps = tbl.getRange("Whole").paragraphs;
          tps.load("items");
          await context.sync();
          next = j + Math.max(1, tps.items.length);
          const t = { table: tbl, paras: tps, rows: tbl.rows };
          t.rows.load("items/cellCount");
          await context.sync();
          const fr = formatTable(t, profile, { existing: true });
          if (fr && fr.autofit === false) autofitUnsupported = true;
          await context.sync();
          tablesDone++;
        } catch (e) {
          tableErrors++; console.warn("tabel gagal dirapikan", e);
          while (next < range.end && infos[next].inTable) next++;
        }
        j = next;
      }
      return { formatted: n, headings: items.filter((i) => i && i.type === "heading").length, quotes: quotesFixed, tables: tablesDone, tableErrors,
        equations: eqRes.converted, inlineMath: eqRes.inline, autofitUnsupported, equationsOk: eqRes.alreadyOk, mathType: eqRes.mathType, equationErrors: eqRes.errors };
    });
  }

  /**
   * Rapikan rumus di bagian: rumus yang masih teks biasa (Y = β0 + β1X1 + ε) diubah jadi persamaan Word
   * dengan penomoran template; nomor "(1)" yang terpisah di baris berikutnya ikut digabung.
   * Rumus Word yang sudah benar tidak disentuh; objek MathType/gambar hanya dihitung (tidak bisa diubah).
   */
  async function fixEquations(context, paras, infos, idx, items, profile) {
    const res = { converted: 0, alreadyOk: 0, mathType: 0, errors: 0, inline: 0 };
    if (!M() || (profile.equation && profile.equation.autoConvert === false)) return res;
    let width = null;
    // (a) LaTeX polos yang terpecah per baris: "[" … "]" (salinan ChatGPT tanpa backslash) -> satu persamaan
    const texts = idx.map((j) => infos[j].text || "");
    const groups = E.findDisplayLatex(texts).filter((g) => idx[g.end] - idx[g.start] === g.end - g.start); // harus berurutan di dokumen
    const used = new Set();
    for (let g = groups.length - 1; g >= 0; g--) {
      const gr = groups[g];
      const r = M().parse(gr.tex);
      if (r.error || !r.ast) continue;
      try {
        if (width == null) width = (await readTextWidthCm(context)) || 0;
        const pkg = M().buildEquationPackage({ ast: r.ast }, E.equationOptions(profile, { number: r.number, textWidthCm: width || undefined }));
        for (let k = gr.end; k > gr.start; k--) paras[idx[k]].delete();
        await context.sync();
        await placeEquation(context, paras[idx[gr.start]], pkg);
        for (let k = gr.start; k <= gr.end; k++) used.add(idx[k]);
        res.converted++;
      } catch (e) { res.errors++; console.warn("rumus gagal diubah", e); }
    }
    // (b) rumus di dalam kalimat: (x_i), (\mathbf W\in\mathbb R^{d}) -> huruf miring + subskrip/superskrip
    const lost = E.isLostLatex(texts.join("\n"), groups.length);
    for (let k = 0; k < idx.length; k++) {
      const j = idx[k];
      if (used.has(j) || !items[k] || !["body", "note", "caption", "quote"].includes(items[k].type)) continue;
      const t = infos[j].text || "";
      const spans = E.inlineMathSpans(t, lost);
      for (let q = spans.length - 1; q >= 0; q--) {
        const [a, e, tex] = spans[q];
        const r = M().parse(tex);
        const src = t.slice(a, e);
        if (r.error || !r.ast || src.length > 250) continue;
        const runs = M().toRuns(r.ast);
        if (!runs.length) continue;
        const hit = paras[j].search(src.replace(/\^/g, "^^"), { matchCase: true }).getFirstOrNullObject();
        hit.load("isNullObject");
        await context.sync();
        if (hit.isNullObject) continue;
        let prev = null;
        runs.forEach((x, n) => {
          const rg = n === 0 ? hit.insertText(x.text, "Replace") : prev.insertText(x.text, "After");
          rg.font.set({ italic: !!x.i, bold: !!x.b, subscript: !!x.sub, superscript: !!x.sup });
          prev = rg;
        });
        res.inline++;
      }
      if (spans.length) await context.sync();
    }
    const RE_NUM = /^\s*\(\s*(\d+[a-z]?)\s*\)\s*$/;
    // kandidat: paragraf rumus + paragraf kosong / "(n)" (bisa berisi objek MathType)
    const cand = [];
    idx.forEach((j, k) => {
      if (used.has(j)) return;
      const t = (infos[j].text || "").trim();
      if ((items[k] && items[k].type === "equation") || !t || RE_NUM.test(t)) cand.push({ j, k, t });
    });
    if (!cand.length) return res;
    const ox = cand.map((c) => paras[c.j].getRange("Whole").getOoxml());
    await context.sync();
    cand.forEach((c, i) => { c.xml = ox[i].value; c.kinds = E.ooxmlKinds(c.xml); });
    const byJ = new Map(cand.map((c) => [c.j, c]));
    res.mathType = cand.filter((c) => !used.has(c.j)).filter((c) => c.kinds.mathType || (c.kinds.ole && !c.t.replace(RE_NUM, ""))).length;
    // dari bawah ke atas
    for (let i = cand.length - 1; i >= 0; i--) {
      const c = cand[i];
      if (!c.t || RE_NUM.test(c.t) || used.has(c.j)) continue; // kosong / nomor saja / sudah diubah
      if (c.kinds.omml) { res.alreadyOk++; continue; }
      if (c.kinds.ole || c.kinds.picture) continue;
      const k = c.k;
      if (!items[k] || items[k].type !== "equation") continue;
      // nomor di paragraf berikutnya ("(1)" sendirian, tanpa objek/gambar)
      const nx = byJ.get(c.j + 1);
      const sepNum = nx && RE_NUM.test(nx.t) && !nx.kinds.omml && !nx.kinds.ole && !nx.kinds.picture ? RE_NUM.exec(nx.t)[1] : null;
      const runs = E.runsFromOoxml(c.xml);
      const it = E.prepareEquation({ type: "equation", runs: runs.length ? runs : [{ text: infos[c.j].text }] }, profile);
      if (!it.eq) continue;
      const number = it.eq.number || sepNum || null;
      try {
        if (width == null) width = (await readTextWidthCm(context)) || 0;
        const pkg = M().buildEquationPackage({ ast: it.eq.ast }, E.equationOptions(profile, { number, textWidthCm: width || undefined }));
        await placeEquation(context, paras[c.j], pkg);
        if (sepNum && !it.eq.number) { paras[c.j + 1].delete(); await context.sync(); }
        res.converted++;
      } catch (e) { res.errors++; console.warn("rumus gagal diubah", e); }
    }
    return res;
  }

  /**
   * Hapus referensi ganda di daftar referensi. Gaya bernomor [n]: nomor referensi diurutkan ulang dan
   * sitasi [n] di badan artikel (termasuk tabel) diganti ke nomor baru.
   */
  async function dedupeReferences(context, paras, infos, idx, range) {
    const res = { removed: 0, list: [], renumbered: 0, citations: 0 };
    const refIdx = idx.filter((j) => (infos[j].text || "").trim());
    const texts = refIdx.map((j) => infos[j].text);
    const d = E.findDuplicateRefs(texts);
    if (!d.remove.length) return res;
    const lastDoc = infos.length - 1;
    for (let k = d.remove.length - 1; k >= 0; k--) {
      const j = refIdx[d.remove[k]];
      if (j === lastDoc) paras[j].clear(); else paras[j].delete(); // paragraf terakhir dokumen tidak bisa dihapus
    }
    await context.sync();
    res.removed = d.remove.length;
    res.list = d.remove.map((r) => texts[r].trim().slice(0, 90));
    if (!d.numbered) return res;

    // --- penomoran ulang [n] ---
    const removed = {};
    d.remove.forEach((r) => { removed[E.refNum(texts[r])] = E.refNum(texts[d.keep[r]]); });
    const rest = refIdx.map((j, i) => i).filter((i) => !d.remove.includes(i));
    const map = E.renumberMap(rest.map((i) => E.refNum(texts[i])), removed);
    for (const i of rest) {
      const t = texts[i], old = E.refNum(t), nw = map.get(old);
      if (old === nw) continue;
      const lab = /^\s*(\[\d+\]|\d+\.)/.exec(t)[1];
      const hit = paras[refIdx[i]].search(lab, { matchCase: true }).getFirstOrNullObject();
      hit.insertText(/^\[/.test(lab) ? `[${nw}]` : `${nw}.`, "Replace");
      res.renumbered++;
    }
    await context.sync();

    // --- sitasi di badan artikel: ganti per kelompok, yang terpanjang dulu, lewat token sementara ---
    const jobs = [];
    infos.forEach((info, j) => {
      if (j >= range.start && j < range.end) return; // daftar referensi sendiri
      const t = info.text || "";
      if (!/\[\s*\d/.test(t)) return;
      const seen = new Set();
      (t.match(E.RE_CITE_GROUP) || []).forEach((g) => {
        if (seen.has(g)) return; seen.add(g);
        const ng = E.renumberCitation(g, map);
        if (ng !== g && g.length <= 250) jobs.push({ j, old: g, nw: ng });
      });
    });
    if (!jobs.length) return res;
    jobs.forEach((jb, k) => { jb.tok = "⟦" + k + "⟧"; });
    const lens = [...new Set(jobs.map((jb) => jb.old.length))].sort((a, b) => b - a);
    for (const L of lens) { // fase 1: sitasi lama -> token (yang panjang dulu agar "[1]" tidak terambil dari "[1], [3]")
      const batch = jobs.filter((jb) => jb.old.length === L);
      batch.forEach((jb) => { jb.hits = paras[jb.j].search(jb.old, { matchCase: true }); jb.hits.load("items"); });
      await context.sync();
      batch.forEach((jb) => { const its = jb.hits.items; its.forEach((h) => h.insertText(jb.tok, "Replace")); res.citations += its.length; });
      await context.sync();
    }
    // fase 2: token -> sitasi baru
    jobs.forEach((jb) => { jb.hits2 = paras[jb.j].search(jb.tok, { matchCase: true }); jb.hits2.load("items"); });
    await context.sync();
    jobs.forEach((jb) => jb.hits2.items.forEach((h) => h.insertText(jb.nw, "Replace")));
    await context.sync();
    return res;
  }

  /** Kutipan yang sudah ada: atribusi di belakang tanda kutip -> (…); tanda kutip blok >= 40 kata dihapus. */
  function fixExistingQuote(p, text, Q) {
    const t = text.replace(/\s+$/, "");
    const close = Math.max(t.lastIndexOf("”"), t.lastIndexOf('"'), t.lastIndexOf("»"));
    const openM = /^\s*([“"«])/.exec(t);
    const tail = close > 0 ? t.slice(close + 1).trim() : "";
    if (tail && Q.attribution === "paren" && tail.length < 250 && !/^\(.*\)$/.test(tail)) {
      const hit = p.search(tail, { matchCase: true }).getLastOrNullObject();
      hit.insertText("(" + E.cleanAttribution(tail) + ")", "Replace").font.set({ italic: false, bold: false });
    }
    const body = close > 0 ? t.slice(openM ? openM.index + 1 : 0, close) : t;
    if (Q.removeMarks && openM && close > 0 && E.countWords(body) >= (Q.minBlockWords || 40)) {
      const closeCh = t[close];
      p.search(closeCh, { matchCase: true }).getLastOrNullObject().delete();
      p.search(openM[1], { matchCase: true }).getFirstOrNullObject().delete();
    }
  }

  function applyExisting(p, it, profile, curBold, curItalic) {
    const f = E.paraFormat(it, profile);
    p.alignment = f.align;
    p.spaceBefore = f.before;
    p.spaceAfter = f.after;
    p.lineSpacing = 12 * f.line;
    p.leftIndent = f.leftIndent;
    p.firstLineIndent = f.firstIndent;
    p.rightIndent = f.rightIndent || 0;
    p.font.name = profile.font;
    p.font.size = f.size;
    if (it.type === "heading" || it.type === "section") {
      p.font.bold = f.bold;
      if (f.italic) p.font.italic = true;
    } else if (it.type === "quote") {
      if (f.italic) p.font.italic = true;
      else if (curItalic === true) p.font.italic = false; // APA: kutipan tidak italic
      if (curBold === true) p.font.bold = false;
    } else if (it.type === "note") {
      if (curBold === true) p.font.bold = false;
      if (profile.note && profile.note.italic) p.font.italic = true;
    } else if (it.type === "caption") {
      // caption: judul tidak bold/italic (label "Table 1." dibold terpisah)
      if (curBold === true) p.font.bold = false;
      if (profile.caption.stripTitleItalic !== false && curItalic === true) p.font.italic = false;
    } else if (curBold === true) {
      p.font.bold = false; // paragraf seluruhnya bold karena gaya template -> normalkan
    }
  }

  /**
   * Hapus SEMUA isi bagian `key` (teks, sub-heading, tabel, gambar) — judul bagian tetap.
   * Isi yang dihapus dikembalikan sebagai OOXML supaya bisa dipulihkan.
   */
  async function clearSection(key, profile) {
    return Word.run(async (context) => {
      const sec = profile.sections.find((s) => s.key === key);
      const { paras, infos, heads } = await readStructure(context, profile);
      const range = E.sectionRange(infos, heads, key);
      if (!range) throw new Error(`Judul "${sec.title}" tidak ditemukan di dokumen.`);
      const n = range.end - range.start - 1;
      let words = 0;
      for (let j = range.start + 1; j < range.end; j++) words += E.countWords(infos[j].text);
      const onlyBlank = infos.slice(range.start + 1, range.end).every((p) => !(p.text || "").trim() && !p.inTable);
      if (n <= 0 || (onlyBlank && n <= 1)) return { deleted: 0, words: 0, ooxml: null };

      const first = paras[range.start + 1];
      const hasNext = range.nextIndex >= 0;
      const lastIdx = paras.length - 1;
      // cadangan isi (untuk "Kembalikan")
      const ox = (hasNext ? first.getRange("Start").expandTo(paras[range.end].getRange("Start")) : first.getRange("Start").expandTo(paras[lastIdx].getRange("End"))).getOoxml();
      await context.sync();
      if (hasNext) first.getRange("Start").expandTo(paras[range.end].getRange("Start")).delete();
      else {
        // bagian terakhir dokumen: paragraf terakhir tidak boleh ikut dihapus (Word membuatnya ulang dan judul bagian
        // bisa tergabung) — hapus sampai sebelum paragraf terakhir, lalu kosongkan paragraf terakhir
        if (range.start + 1 < lastIdx) first.getRange("Start").expandTo(paras[lastIdx].getRange("Start")).delete();
        paras[lastIdx].clear();
      }
      await context.sync();
      if (hasNext && profile.layout.blankBeforeSection) {
        // sisakan satu baris kosong sebelum judul bagian berikutnya
        const st = await readStructure(context, profile);
        const r2 = E.sectionRange(st.infos, st.heads, key);
        if (r2 && r2.end - r2.start - 1 === 0) {
          const blank = st.paras[r2.start].insertParagraph("", "After");
          formatNewParagraph(blank, { type: "blank" }, profile);
        }
      }
      // letakkan kursor di akhir judul bagian — ambil ulang paragrafnya (objek lama bisa sudah tidak berlaku)
      try {
        const st2 = await readStructure(context, profile);
        const r3 = E.sectionRange(st2.infos, st2.heads, key);
        if (r3) { st2.paras[r3.start].getRange("End").select("End"); await context.sync(); }
      } catch (e) { /* memindahkan kursor tidak penting */ }
      return { deleted: n, words, ooxml: ox.value };
    });
  }

  /** Pulihkan isi yang dihapus clearSection (disisipkan tepat di bawah judul bagian). */
  async function restoreSection(key, ooxml, profile) {
    return Word.run(async (context) => {
      const sec = profile.sections.find((s) => s.key === key);
      const { paras, infos, heads } = await readStructure(context, profile);
      const range = E.sectionRange(infos, heads, key);
      if (!range) throw new Error(`Judul "${sec.title}" tidak ditemukan di dokumen.`);
      const holder = paras[range.start].insertParagraph("", "After");
      holder.insertOoxml(ooxml, "Replace");
      await context.sync();
      return true;
    });
  }

  /* ---------------- sel & tabel tanpa rowIndex/cellIndex ----------------
   * Word melempar GeneralException untuk TableCell.rowIndex/cellIndex pada tabel
   * dengan sel gabungan. Jadi posisi sel dihitung dari isi TableCell.body.
   */
  async function cellRange(context, paras, infos, j) {
    const cell = paras[j].parentTableCellOrNullObject;
    const ps = cell.body.paragraphs;
    ps.load("items/text");
    await context.sync();
    const texts = ps.items.map((p) => p.text);
    // posisi paragraf j di dalam selnya: cocokkan urutan teks di sekitar j
    let best = -1;
    for (let k = 0; k < texts.length; k++) {
      const start = j - k;
      if (start < 0) break;
      let okAll = true;
      for (let q = 0; q < texts.length; q++) { const g = infos[start + q]; if (!g || g.text !== texts[q]) { okAll = false; break; } }
      if (okAll) { best = start; break; }
    }
    if (best < 0) throw new Error("Posisi sel tidak dapat ditentukan.");
    return { start: best, end: best + texts.length };
  }

  /* ---------------- abstrak ---------------- */
  async function locateAbstract(context, profile) {
    const { paras, infos, heads } = await readStructure(context, profile);
    const h = infos.findIndex((p) => E.RE_ABSTRACT_HEAD.test((p.text || "").trim()));
    if (h < 0) return null;
    if (!infos[h].inTable) return { paras, infos, heads, h, pick: E.pickAbstractTarget(infos, h) };
    let tblEnd = h + 1;
    while (tblEnd < infos.length && infos[tblEnd].inTable) tblEnd++;
    try {
      const hc = await cellRange(context, paras, infos, h);
      let pick = null;
      if (hc.end - h > 1) pick = E.abstractTargetsInCell(infos, h, hc.start, hc.end);          // abstrak di sel yang sama (AIEDS, VOICE)
      if (!pick || !pick.targets.some((j) => (infos[j].text || "").trim())) {
        // sel lain: sel pertama setelah judul yang berisi paragraf panjang (bukan keywords/history/cite)
        const j = infos.findIndex((p, q) => q >= hc.end && q < tblEnd && E.countWords(p.text) >= 20 && !E.isAbstractNoise(p.text));
        if (j > 0) { const c = await cellRange(context, paras, infos, j); pick = E.abstractTargetsInCell(infos, h, c.start, c.end); }
        else if (!pick) {
          // template kosong: sel pertama setelah judul yang bukan keywords/history/cite
          let q = hc.end;
          for (let guard = 0; q < tblEnd && guard < 12; guard++) {
            const c = await cellRange(context, paras, infos, q);
            const first = infos.slice(c.start, c.end).map((p) => (p.text || "").trim()).find(Boolean) || "";
            if (!first || /open access|cc[ -]by|licen/i.test(first) || !E.isAbstractNoise(first)) { pick = E.abstractTargetsInCell(infos, h, c.start, c.end); break; }
            q = c.end;
          }
        }
      }
      if (pick) return { paras, infos, heads, h, pick };
    } catch (e) { console.warn("lokasi sel abstrak via body gagal, pakai cara teks", e); }
    // cadangan berbasis teks: paragraf panjang pertama setelah judul di tabel yang sama
    const j = infos.findIndex((p, q) => q > h && q < tblEnd && E.countWords(p.text) >= 20 && !E.isAbstractNoise(p.text));
    return { paras, infos, heads, h, pick: j > 0 ? { targets: [j], after: null, cellEnd: j + 1 } : { targets: [], after: h } };
  }

  /** Hapus paragraf sisa; paragraf terakhir dalam sel tabel tidak bisa dihapus Word -> dikosongkan saja. */
  function dropExtras(paras, infos, list, cellEnd) {
    for (let k = list.length - 1; k >= 0; k--) {
      const j = list[k];
      const last = infos[j].inTable && (cellEnd != null ? j === cellEnd - 1 : !(infos[j + 1] && infos[j + 1].inTable));
      if (last) paras[j].clear(); else paras[j].delete();
    }
  }

  function writeAbstractParagraph(p, runs, profile) {
    const f = E.paraFormat({ type: "abstract" }, profile);
    // gaya paragraf (mis. Heading 3 + latar merah muda di AIEDS) sengaja TIDAK diubah
    p.alignment = f.align;
    p.spaceBefore = f.before;
    p.spaceAfter = f.after;
    p.lineSpacing = 12 * f.line;
    p.leftIndent = 0; p.firstLineIndent = 0;
    p.font.set({ name: profile.font, size: f.size, bold: false, italic: f.italic, color: profile.color, underline: "None" });
    let last = null;
    for (const r of runs) {
      if (r.br) { if (last) last.insertBreak("Line", "After"); continue; }
      if (!r.text) continue;
      last = p.insertText(r.text, "End");
      last.font.set({ name: profile.font, size: f.size, bold: !!r.b, italic: !!r.i, superscript: !!r.sup, subscript: !!r.sub, color: profile.color, underline: "None" });
    }
  }

  /** Ganti isi abstrak di dokumen (atau sisipkan di kursor). */
  async function insertAbstract(items, profile, mode) {
    const out = E.abstractParagraphs(items, profile);
    if (!out.length) throw new Error("Tidak ada teks abstrak untuk disisipkan.");
    return Word.run(async (context) => {
      let p0, backup = null;
      if (mode === "cursor") {
        const sel = context.document.getSelection().paragraphs.getLast();
        sel.load("text,tableNestingLevel");
        await context.sync();
        p0 = sel.text.trim() ? sel.insertParagraph("", "After") : sel;
        if (!sel.text.trim()) p0.clear();
      } else {
        const loc = await locateAbstract(context, profile);
        if (!loc) throw new Error("Judul ABSTRACT tidak ditemukan di dokumen. Pakai mode 'Posisi kursor'.");
        const { paras, infos, pick } = loc;
        if (pick.targets.length) {
          const t = pick.targets;
          const ox = paras[t[0]].getRange("Start").expandTo(paras[t[t.length - 1]].getRange("End")).getOoxml();
          await context.sync();
          backup = ox.value;
          // pakai ulang paragraf yang ada (gaya & latar tiap paragraf terjaga), sisanya dibuang
          const reuse = t.slice(0, out.length);
          reuse.forEach((j) => paras[j].clear());
          dropExtras(paras, infos, t.slice(out.length), pick.cellEnd);
          reuse.forEach((j, k) => writeAbstractParagraph(paras[j], out[k], profile));
          p0 = paras[t[0]];
          let prev = paras[reuse[reuse.length - 1]];
          for (let k = reuse.length; k < out.length; k++) {
            const q = prev.insertParagraph("", "After");
            writeAbstractParagraph(q, out[k], profile);
            prev = q;
          }
          p0.getRange("Start").select("Start");
          await context.sync();
          const words0 = out.reduce((a, p) => a + E.countWords(p.filter((r) => !r.br).map((r) => r.text).join(" ")), 0);
          return { inserted: out.length, words: words0, backup };
        }
        p0 = paras[pick.after].insertParagraph("", "After");
      }
      writeAbstractParagraph(p0, out[0], profile);
      let prev = p0;
      for (let k = 1; k < out.length; k++) {
        const q = prev.insertParagraph("", "After");
        writeAbstractParagraph(q, out[k], profile);
        prev = q;
      }
      p0.getRange("Start").select("Start");
      await context.sync();
      const words = out.reduce((a, p) => a + E.countWords(p.filter((r) => !r.br).map((r) => r.text).join(" ")), 0);
      return { inserted: out.length, words, backup };
    });
  }

  /** Kosongkan abstrak (sisakan satu paragraf kosong). */
  async function clearAbstract(profile) {
    return Word.run(async (context) => {
      const loc = await locateAbstract(context, profile);
      if (!loc) throw new Error("Judul ABSTRACT tidak ditemukan di dokumen.");
      const { paras, infos, pick } = loc;
      const t = pick.targets;
      if (!t.length || t.every((j) => !(infos[j].text || "").trim())) return { deleted: 0, words: 0, ooxml: null };
      const words = t.reduce((a, j) => a + E.countWords(infos[j].text), 0);
      const ox = paras[t[0]].getRange("Start").expandTo(paras[t[t.length - 1]].getRange("End")).getOoxml();
      await context.sync();
      dropExtras(paras, infos, t.slice(1), pick.cellEnd);
      paras[t[0]].clear();
      await context.sync();
      return { deleted: t.length, words, ooxml: ox.value };
    });
  }

  async function restoreAbstract(ooxml, profile) {
    return Word.run(async (context) => {
      const loc = await locateAbstract(context, profile);
      if (!loc) throw new Error("Judul ABSTRACT tidak ditemukan di dokumen.");
      const { paras, infos, pick } = loc;
      const t = pick.targets;
      if (!t.length) { paras[pick.after].insertParagraph("", "After").insertOoxml(ooxml, "Replace"); }
      else {
        dropExtras(paras, infos, t.slice(1), pick.cellEnd);
        paras[t[0]].insertOoxml(ooxml, "Replace");
      }
      await context.sync();
      return true;
    });
  }

  /** Kumpulkan data untuk Cek kesesuaian: badan dokumen, semua header/footer, halaman awal. */
  async function collectMeta(profile) {
    return Word.run(async (context) => {
      const { paras, infos } = await readStructure(context, profile);
      const secs = context.document.sections;
      secs.load("items");
      await context.sync();
      const parts = [];
      secs.items.forEach((sct, si) => {
        ["FirstPage", "Primary", "EvenPages"].forEach((type) => {
          for (const kind of ["header", "footer"]) {
            const body = kind === "header" ? sct.getHeader(type) : sct.getFooter(type);
            const ps = body.paragraphs;
            ps.load("items/text");
            parts.push({ kind, type, section: si, ps });
          }
        });
      });
      await context.sync();
      const hf = parts.map((x) => ({ kind: x.kind, type: x.type, section: x.section, text: x.ps.items.map((p) => p.text).join("\n") }))
        .filter((x) => x.text.trim());
      // nomor halaman awal (w:pgNumType w:start) — dari OOXML paragraf terakhir; abaikan bila tidak tersedia
      let startPage = null;
      try {
        const ox = paras[paras.length - 1].getRange("Whole").getOoxml();
        await context.sync();
        const m = /<w:pgNumType[^>]*w:start="(\d+)"/.exec(ox.value);
        if (m) startPage = +m[1];
      } catch (e) { /* tidak didukung */ }
      return { infos, hf, startPage };
    });
  }

  async function countWords(profile) {
    return Word.run(async (context) => {
      const { infos, heads } = await readStructure(context, profile);
      return E.countDocument(infos, heads, profile);
    });
  }

  async function undoLast() {
    if (!has("1.4")) throw new Error("Fitur ini butuh Word versi lebih baru (WordApi 1.4). Gunakan Ctrl+Z.");
    return Word.run(async (context) => {
      const rng = context.document.getBookmarkRangeOrNullObject(BOOKMARK);
      rng.load("isNullObject,text");
      await context.sync();
      if (rng.isNullObject) return false;
      rng.delete();
      context.document.deleteBookmark(BOOKMARK);
      await context.sync();
      return true;
    });
  }

  /* ---------------- tab Rumus ---------------- */
  /** Nomor rumus berikutnya: (n) terbesar di akhir paragraf rumus + 1. */
  async function nextEquationNumber() {
    return Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();
      let max = 0;
      for (const p of paras.items) {
        const t = (p.text || "").trim();
        if (t.length > 400) continue;
        const m = /(?:^|[\s\t#)\]}])\((\d{1,3})[a-z]?\)$/.exec(t);
        if (!m) continue;
        const body = t.slice(0, t.length - m[0].length + 1);
        if (/[=≤≥≈<>]/.test(body) || /\t/.test(p.text) || E.looksLikeEquation(t)) max = Math.max(max, +m[1]);
      }
      return max + 1;
    });
  }

  /**
   * Sisipkan rumus di posisi kursor.
   * eq: {ast|omml, number}; opts.mode: 'display' (paragraf sendiri) | 'inline' (di dalam kalimat)
   */
  async function insertEquation(eq, profile, opts) {
    opts = opts || {};
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      if (opts.mode === "inline") {
        const pkg = M().buildEquationPackage(eq, E.equationOptions(profile, { mode: "inline" }));
        const r = sel.insertOoxml(pkg, "Replace");
        await context.sync();
        try { r.select("End"); await context.sync(); } catch (e) { /* abaikan */ }
        return { mode: "inline" };
      }
      const cur = sel.paragraphs.getLast();
      cur.load("text");
      await context.sync();
      const target = (cur.text || "").trim() ? cur.insertParagraph("", "After") : cur;
      const width = await readTextWidthCm(context);
      const pkg = M().buildEquationPackage(eq, E.equationOptions(profile, { number: eq.number, textWidthCm: width || undefined }));
      const eqp = await placeEquation(context, target, pkg);
      try { eqp.getRange("End").select("End"); await context.sync(); } catch (e) { /* abaikan */ }
      return { mode: "display", number: eq.number, textWidthCm: width };
    });
  }

  /** Teks terpilih -> rumus. Seluruh paragraf terpilih = rumus baris (nomor "(n)" di akhir dipertahankan). */
  async function convertSelection(profile, opts) {
    opts = opts || {};
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      const ps = sel.paragraphs;
      ps.load("items/text");
      await context.sync();
      const text = (sel.text || "").trim();
      if (!text) throw new Error("Pilih (blok) dulu teks rumus di dokumen.");
      if (ps.items.length > 1) throw new Error("Pilih satu baris rumus saja.");
      const r = M().parse(text);
      if (r.error) throw new Error(r.error);
      const p = ps.items[0];
      const whole = (p.text || "").trim() === text;
      if (!whole || opts.mode === "inline") {
        const pkg = M().buildEquationPackage({ ast: r.ast }, E.equationOptions(profile, { mode: "inline" }));
        sel.insertOoxml(pkg, "Replace");
        await context.sync();
        return { mode: "inline", tex: r.tex };
      }
      const number = r.number || (opts.number != null && opts.number !== "" ? opts.number : null);
      const width = await readTextWidthCm(context);
      const pkg = M().buildEquationPackage({ ast: r.ast }, E.equationOptions(profile, { number, textWidthCm: width || undefined }));
      await placeEquation(context, p, pkg);
      return { mode: "display", tex: r.tex, number };
    });
  }

  /** Ambil rumus Word yang sedang dipilih -> { ast, tex, number, count } (null bila tidak ada). */
  async function readSelectedEquation() {
    return Word.run(async (context) => {
      const sel = context.document.getSelection();
      let ox = sel.getOoxml();
      const ps = sel.paragraphs; ps.load("items/text");
      await context.sync();
      let xml = ox.value;
      if (!hasMathXml(xml) && ps.items.length) { // kursor di dalam/di samping rumus -> ambil paragrafnya
        ox = ps.items[0].getRange("Whole").getOoxml();
        await context.sync();
        xml = ox.value;
      }
      if (!hasMathXml(xml)) return null;
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
      const all = Array.from(doc.getElementsByTagNameNS(MNS, "oMath"));
      if (!all.length) return null;
      const T = M();
      const asts = all.map((om) => T.ommlToAst(om));
      let ast = asts.length === 1 ? asts[0] : { t: "row", items: asts };
      // nomor: "#(n)" bawaan Word atau "(n)" teks setelah rumus
      let number = null;
      const flat = T.toLatex(ast);
      const nm = /#\\left\(\s*(\d+[a-z]?)\s*\\right\)\s*$/.exec(flat) || /#\(\s*(\d+[a-z]?)\s*\)\s*$/.exec(flat);
      if (nm) {
        number = nm[1];
        // buang "#(n)" dari AST (eqArr satu baris)
        const strip = (n) => { if (n && n.t === "eqarr" && n.rows.length === 1) return strip(n.rows[0]); return n; };
        ast = strip(ast);
        if (ast.t === "row") { const it = ast.items; while (it.length && !(it[it.length - 1].t === "id" && /#/.test(it[it.length - 1].v))) it.pop(); if (it.length) { const last = it[it.length - 1]; last.v = last.v.replace(/#\s*$/, ""); if (!last.v) it.pop(); } }
      } else {
        const tail = xml.slice(xml.lastIndexOf("</m:oMath>"));
        const tailText = (tail.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("");
        const pt = ps.items.length ? (ps.items[0].text || "") : "";
        const m2 = /\(\s*(\d{1,3}[a-z]?)\s*\)\s*$/.exec(tailText) || /\((\d{1,3}[a-z]?)\)\s*$/.exec(pt);
        if (m2) number = m2[1];
      }
      return { ast, tex: T.toLatex(ast), number, count: all.length };
    });
  }

  root.WordOps = { insertEquation, convertSelection, readSelectedEquation, nextEquationNumber, collectMeta, scan, insertSection, reformatSection, clearSection, restoreSection, insertAbstract, clearAbstract, restoreAbstract, countWords, undoLast, _internal: { insertItems, formatNewParagraph, setStyle, placeEquation, readTextWidthCm, formatTable } };
})(typeof self !== "undefined" ? self : this);
