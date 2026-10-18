/* ============================================================
   MOCK BRIDGE — mode demo di browser biasa (tanpa Word).
   Aktif hanya bila URL memuat "?demo". Meniru WordBridge di atas model dokumen
   sederhana (paragraf berisi runs {text, italic}), jadi seluruh alur add-in bisa
   dicoba dan diuji otomatis tanpa Microsoft Word. Lihat demo.html.
   ============================================================ */
(function (root) {
  'use strict';
  if (!/[?&]demo\b/.test(root.location ? root.location.search : '')) return;
  root.CEK_SITASI_DEMO = true;

  var host = (function () { try { return root.parent && root.parent !== root && root.parent.DemoDoc ? root.parent.DemoDoc : null; } catch (e) { return null; } })();
  var doc = host || {
    paragraphs: [
      'Pendahuluan',
      'Pembelajaran mesin berkembang pesat (Smith & Jones, 2020). Menurut Brown et al. (2019), model besar membutuhkan data. Penelitian lain (Lee and Park, 2021; Adams, 2018) menunjukkan hal serupa. Garcia, Lopez, dan Chen (2022) membahas etika. Studi terbaru (Wang, 2023) juga relevan.',
      'Daftar Pustaka',
      'Smith, J., & Jones, K. (2020). Deep learning for text classification. Journal of Machine Learning Research, 21(4), 1-20. https://doi.org/10.1234/jmlr.2020.004',
      'Brown, T., Mann, B., Ryder, N., & Subbiah, M. (2019). Language models are few-shot learners. Advances in Neural Information Processing Systems, 33, 1877-1901.',
      'Lee, H., and Park, S. (2021). Attention mechanisms revisited. Neural Networks, 140, 12–25.',
      'Garcia, M., Lopez, R., & Chen, L. (2022). Ethics in artificial intelligence. AI and Society, 37(2), 400–415. doi:10.1007/s00146-021-01234-5',
      'Adams, P. (2018). Statistical learning theory. Springer.'
    ].map(function (t) { return { runs: [{ text: t, italic: false }], links: [], bookmarks: [] }; }),
    listeners: [], selection: null, track: false, changed: function () {}
  };
  function textOf(p) { return p.runs.map(function (r) { return r.text; }).join(''); }
  function notify() { if (doc.changed) doc.changed(); }

  // ---- operasi runs ----
  function splitAt(p, pos) {
    var acc = 0;
    for (var i = 0; i < p.runs.length; i++) {
      var r = p.runs[i];
      if (pos === acc) return i;
      if (pos < acc + r.text.length) {
        p.runs.splice(i, 1, { text: r.text.slice(0, pos - acc), italic: r.italic }, { text: r.text.slice(pos - acc), italic: r.italic });
        return i + 1;
      }
      acc += r.text.length;
    }
    return p.runs.length;
  }
  function replaceRange(p, start, len, newRuns) {
    var a = splitAt(p, start), b = splitAt(p, start + len);
    var inherit = p.runs[a] ? p.runs[a].italic : false;
    p.runs.splice.apply(p.runs, [a, b - a].concat(newRuns.map(function (r) { return { text: r.text, italic: r.italic == null ? inherit : r.italic }; })));
    p.runs = p.runs.filter(function (r) { return r.text; });
    var delta = newRuns.reduce(function (n, r) { return n + r.text.length; }, 0) - len;
    p.links = p.links.filter(function (l) { return l.start + l.len <= start || l.start >= start + len; }).map(function (l) { return l.start >= start + len ? Object.assign({}, l, { start: l.start + delta }) : l; });
    if (doc.track) (p.revisions = p.revisions || []).push({ start: start, len: len, text: newRuns.map(function (r) { return r.text; }).join('') });
  }
  function setItalic(p, start, len, on) {
    var a = splitAt(p, start), b = splitAt(p, start + len);
    for (var i = a; i < b; i++) p.runs[i].italic = on;
  }
  function isWord(c) { return !!c && /[\p{L}\p{N}_]/u.test(c); }
  function searchIn(text, needle, whole) {
    var out = [], i = 0;
    while (needle) {
      var j = text.indexOf(needle, i); if (j < 0) break;
      if (!whole || (!isWord(text[j - 1]) && !isWord(text[j + needle.length]))) { out.push(j); i = j + needle.length; } else i = j + 1;
    }
    return out;
  }
  function locate(it) {
    var p = doc.paragraphs[it.paraIndex]; if (!p) return null;
    var hits = searchIn(textOf(p), it.text, it.wholeWord);
    if (!hits.length) return null;
    return { p: p, start: hits[Math.min(Math.max(it.occurrence || 0, 0), hits.length - 1)], len: it.text.length };
  }
  function check(snapshot) {
    Object.keys(snapshot || {}).forEach(function (i) { if (!doc.paragraphs[i] || textOf(doc.paragraphs[i]) !== snapshot[i]) throw new Error('Dokumen sudah berubah sejak diperiksa. Klik "Periksa" lalu ulangi.'); });
  }
  function norm(t) { return String(t || '').replace(/^\s*(?:\[\d+\]|\d+\.)\s*/, '').replace(/\s+/g, ' ').trim(); }
  function later(fn) { return new Promise(function (res, rej) { setTimeout(function () { try { res(fn()); notify(); } catch (e) { rej(e); } }, 30); }); }

  var api = {
    // "?demo&old" meniru Word 2016/2019 yang belum punya WordApi 1.4 dan menolak bookmark.
    caps: function () { var old = /[?&]old\b/.test(root.location.search); return { api13: true, api14: !old, api16: false, annotations: false, popup: false }; },
    readParagraphs: function () { return later(function () { return doc.paragraphs.map(function (p) { return { text: textOf(p), isListItem: false }; }); }); },
    select: function (item) { return later(function () { var r = item.kind === 'replace' ? locate({ paraIndex: item.paraIndex, text: item.original, occurrence: item.occurrence }) : null; doc.selection = r ? { paraIndex: item.paraIndex, start: r.start, len: r.len } : { paraIndex: item.paraIndex }; }); },
    applySuggestion: function (s) {
      return later(function () {
        var p = doc.paragraphs[s.paraIndex]; if (!p) throw new Error('Paragraf tidak ditemukan.');
        if (s.kind === 'replace') { var r = locate({ paraIndex: s.paraIndex, text: s.original, occurrence: s.occurrence }); if (!r) throw new Error('Teks asli tidak ditemukan.'); replaceRange(p, r.start, r.len, [{ text: s.replacement }]); return; }
        if (s.kind === 'paragraph') { if (norm(textOf(p)) !== norm(s.compareOriginal || s.original)) throw new Error('Isi referensi sudah berubah.'); p.runs = (s.replacementRuns || [{ text: s.replacement, italic: false }]).map(function (x) { return { text: x.text, italic: !!x.italic }; }); p.links = []; return; }
        if (s.kind === 'append') { var t = textOf(p); if (s.stripTrailingPeriod && /\.$/.test(t)) replaceRange(p, t.length - 1, 1, [{ text: s.appendText, italic: false }]); else replaceRange(p, t.length, 0, [{ text: s.appendText, italic: false }]); return; }
        if (s.kind === 'italic') { s.segments.forEach(function (seg) { var r = locate({ paraIndex: s.paraIndex, text: seg.text, occurrence: seg.occurrence }); if (r) setItalic(p, r.start, r.len, true); }); return; }
        throw new Error('Jenis saran tidak dikenal.');
      });
    },
    applyMany: function (list, opts) {
      var done = 0, failed = 0, chain = Promise.resolve();
      list.slice().sort(function (a, b) { return (b.paraIndex - a.paraIndex) || ((b.offset || 0) - (a.offset || 0)); }).forEach(function (s) {
        chain = chain.then(function () { return api.applySuggestion(s, opts).then(function () { done++; }, function () { failed++; }); });
      });
      return chain.then(function () { return { done: done, failed: failed }; });
    },
    checkItalics: function (plan) {
      return later(function () {
        var bad = {};
        (plan || []).forEach(function (it) {
          it.segments.forEach(function (seg) {
            var r = locate({ paraIndex: it.paraIndex, text: seg.text, occurrence: seg.occurrence }); if (!r) return;
            var acc = 0, ok = true;
            r.p.runs.forEach(function (run) { var s0 = acc, e0 = acc + run.text.length; acc = e0; if (e0 > r.start && s0 < r.start + r.len && !run.italic && run.text.trim()) ok = false; });
            if (!ok) (bad[it.refIndex] = bad[it.refIndex] || { refIndex: it.refIndex, paraIndex: it.paraIndex, segments: [] }).segments.push(seg);
          });
        });
        return Object.keys(bad).map(function (k) { return bad[k]; });
      });
    },
    reorderParagraphs: function (idx, order, snapshot) {
      return later(function () { check(snapshot); var copies = idx.map(function (i) { return JSON.parse(JSON.stringify(doc.paragraphs[i])); }); idx.forEach(function (i, slot) { doc.paragraphs[i] = copies[order[slot]]; }); });
    },
    applyConversion: function (plan, snapshot) {
      return later(function () {
        check(snapshot);
        var applied = 0;
        var edits = plan.edits.filter(function (e) { return !e.crossesParagraph && e.occurrence >= 0; });
        var located = edits.map(function (e) { return locate({ paraIndex: e.paraIndex, text: e.original, occurrence: e.occurrence }); });
        edits.map(function (e, i) { return { e: e, r: located[i] }; }).sort(function (a, b) { return (b.e.paraIndex - a.e.paraIndex) || (b.e.offset - a.e.offset); }).forEach(function (x) {
          if (x.r) { replaceRange(x.r.p, x.r.start, x.r.len, [{ text: x.e.replacement }]); applied++; }
        });
        var idx = plan.references.paraIndexes, lines = plan.references.lines, n = Math.min(idx.length, lines.length);
        for (var k = 0; k < n; k++) doc.paragraphs[idx[k]] = { runs: lines[k].runs.map(function (r) { return { text: r.text, italic: !!r.italic }; }), links: [], bookmarks: [], highlight: !!lines[k].risky, hanging: true };
        (plan.references.flagParaIndexes || []).forEach(function (i) { if (doc.paragraphs[i]) doc.paragraphs[i].highlight = true; });
        var extra = lines.slice(n).map(function (l) { return { runs: l.runs.map(function (r) { return { text: r.text, italic: !!r.italic }; }), links: [], bookmarks: [] }; });
        if (extra.length) doc.paragraphs.splice.apply(doc.paragraphs, [idx[idx.length - 1] + 1, 0].concat(extra));
        for (var d = idx.length - 1; d >= n; d--) doc.paragraphs.splice(idx[d], 1);
        return { applied: applied, skipped: plan.edits.length - applied };
      });
    },
    applyLinks: function (plan, snapshot, opts) {
      var o = opts || {};
      return later(function () {
        check(snapshot);
        if (/[?&]old\b/.test(root.location.search)) { var e = new Error('Word versi ini belum bisa membuat penanda (bookmark) lewat add-in.'); e.code = 'LinkUnsupported'; throw e; }
        plan.bookmarks.forEach(function (b) { doc.paragraphs[b.paraIndex].bookmarks = [b.name]; });
        var items = (o.citations === false ? [] : plan.citationLinks).concat(o.external === false ? [] : plan.externalLinks);
        var linked = 0, missed = 0;
        items.forEach(function (it) { var r = locate(it); if (!r) { missed++; return; } r.p.links.push({ start: r.start, len: r.len, href: it.bookmark ? '#' + it.bookmark : it.href }); linked++; });
        return { linked: linked, missed: missed, bookmarks: plan.bookmarks.length };
      });
    },
    removeLinks: function () {
      return later(function () { var n = 0, b = 0; doc.paragraphs.forEach(function (p) { var before = p.links.length; p.links = p.links.filter(function (l) { return !/^#_CS_REF_/.test(l.href); }); n += before - p.links.length; b += p.bookmarks.length; p.bookmarks = []; }); return { removed: n, bookmarks: b }; });
    },
    insertReference: function (point, runs) {
      return later(function () {
        var para = { runs: runs.map(function (r) { return { text: r.text, italic: !!r.italic }; }), links: [], bookmarks: [], inserted: true };
        if (point.paraIndex == null || point.paraIndex < 0) { doc.paragraphs.push({ runs: [{ text: 'Daftar Pustaka', italic: false }], links: [], bookmarks: [] }, para); return; }
        doc.paragraphs.splice(point.paraIndex + (point.where === 'Before' ? 0 : 1), 0, para);
      });
    },
    // Mode salinan: dokumen tiruan dikemas jadi .docx sungguhan (JSZip), lalu hasil tautan
    // dibaca balik dan ditampilkan sebagai "dokumen baru" di halaman demo.
    canReadFile: function () { return !!root.JSZip; },
    readDocxFile: function () {
      function x(t) { return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
      var body = doc.paragraphs.map(function (p) {
        return '<w:p>' + p.runs.map(function (r) { return '<w:r>' + (r.italic ? '<w:rPr><w:i/></w:rPr>' : '') + '<w:t xml:space="preserve">' + x(r.text) + '</w:t></w:r>'; }).join('') + '</w:p>';
      }).join('');
      var zip = new root.JSZip();
      zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
      zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
      zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
      zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' + body + '</w:body></w:document>');
      return zip.generateAsync({ type: 'uint8array' });
    },
    openNewDocument: function (base64) {
      return root.JSZip.loadAsync(base64, { base64: true }).then(function (zip) {
        return Promise.all([zip.file('word/document.xml').async('string'), zip.file('word/_rels/document.xml.rels').async('string')]);
      }).then(function (parts) {
        var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
        var xml = new DOMParser().parseFromString(parts[0], 'application/xml'), rels = new DOMParser().parseFromString(parts[1], 'application/xml');
        var relMap = {}; Array.prototype.forEach.call(rels.getElementsByTagName('Relationship'), function (r) { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
        var paras = Array.prototype.map.call(xml.getElementsByTagNameNS(W, 'p'), function (p) {
          var out = { runs: [], links: [], bookmarks: [] }, pos = 0;
          (function walk(node, link) {
            Array.prototype.forEach.call(node.childNodes, function (c) {
              if (c.nodeType !== 1) return;
              if (c.localName === 'hyperlink') {
                var href = c.getAttributeNS(W, 'anchor') ? '#' + c.getAttributeNS(W, 'anchor') : relMap[c.getAttributeNS(R, 'id')] || '';
                var start = pos; walk(c, true); out.links.push({ start: start, len: pos - start, href: href }); return;
              }
              if (c.localName === 'bookmarkStart') { out.bookmarks.push(c.getAttributeNS(W, 'name')); return; }
              if (c.localName === 'r') {
                var it = c.getElementsByTagNameNS(W, 'i').length > 0;
                Array.prototype.forEach.call(c.getElementsByTagNameNS(W, 't'), function (t) { out.runs.push({ text: t.textContent, italic: it }); pos += t.textContent.length; });
                return;
              }
              walk(c, link);
            });
          })(p, false);
          return out;
        });
        doc.paragraphs = paras; doc.newDocumentOpened = true;
        if (doc.changed) doc.changed();
      });
    },
    annotate: function () { return Promise.resolve({ supported: false, ids: {} }); },
    clearAnnotations: function () { return Promise.resolve(); },
    onAnnotationEvents: function () { return Promise.resolve(false); },
    onDocumentChanged: function (cb) { doc.listeners.push(cb); return Promise.resolve(); }
  };
  root.DemoDocModel = doc;
  root.DemoTextOf = textOf;
  root.WordBridge = api;
})(typeof window !== 'undefined' ? window : this);
