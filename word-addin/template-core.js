/* ============================================================
   TEMPLATE CORE — Cek Template (setara "Preliminary Check" di cek-sitasi web), dihitung
   dengan aturan hitung kata Microsoft Word.

     - jumlah kata Abstrak, Introduction, Results and Discussion, Conclusion, dan total;
     - "How to cite": rentang halaman vs halaman sebenarnya, kesesuaian judul, plus
       tahun/volume/nomor/DOI/jurnal vs header dan penulis vs byline;
     - kata kunci: jumlah & urutan alfabetis.

   ATURAN HITUNG KATA WORD (dipakai countWords):
     - kata dipisahkan spasi apa pun (spasi, tab, ganti baris, akhir paragraf, spasi tak-putus);
     - en dash (–) dan em dash (—) juga pemisah: "2023–2024" = 2 kata, "policy–practice" = 2;
       dash yang berdiri sendiri (" – ") tidak dihitung;
     - tanda hubung (-), garis miring (/), dan tanda baca lain TIDAK memisah: "self-efficacy",
       "Originality/Value" = 1 kata; token tanda baca yang berdiri sendiri ("-", "&", "…") = 1 kata;
     - kode field, teks terhapus (Track Changes), teks tersembunyi, dan nomor daftar otomatis
       tidak dihitung; hasil field (mis. sitasi Zotero) dihitung;
     - total dokumen = badan (termasuk tabel) + kotak teks + catatan kaki + catatan akhir,
       seperti bilah status Word dengan "Include textboxes, footnotes and endnotes";
       header & footer tidak dihitung.

   Murni: tanpa Office.js. Diuji di test/template.test.js dengan dua naskah contoh.
   ============================================================ */
(function (root) {
  'use strict';

  // ---------- HITUNG KATA ----------
  var SEP = /[\s –—]+/;
  var INVISIBLE = /^[​-‏⁠-⁤﻿­]*$/;
  function countWords(text) {
    var parts = String(text || '').split(SEP), n = 0;
    for (var i = 0; i < parts.length; i++) if (parts[i] && !INVISIBLE.test(parts[i])) n++;
    return n;
  }

  function clean(t) { return String(t || '').replace(/[​-‏⁠-⁤﻿]/g, '').replace(/[\s ]+/g, ' ').trim(); }
  function fold(t) {
    return clean(t).normalize('NFC')
      .replace(/[‘’‚‛′`´]/g, "'").replace(/[“”„″]/g, '"')
      .replace(/[‐-―−]/g, '-').replace(/\s*-\s*/g, '-').replace(/\s*:\s*/g, ': ')
      .replace(/[.\s]+$/, '');
  }
  function words(t) { return fold(t).toLowerCase().split(/[^\p{L}\p{N}']+/u).filter(Boolean); }

  // ---------- STRUKTUR ----------
  var KNOWN = [
    ['introduction', 'Introduction', /^(introduction|pendahuluan|latar\s+belakang)$/i],
    ['literature', 'Literature Review', /^(literature\s+review|theoretical\s+(framework|background)|kajian\s+(pustaka|teori)|tinjauan\s+pustaka|landasan\s+teori)$/i],
    ['method', 'Method', /^((research\s+)?methods?|methodology|research\s+design|materials?\s+and\s+methods?|metode(\s+penelitian)?|metodologi(\s+penelitian)?)$/i],
    ['results_discussion', 'Results and Discussion', /^((research\s+)?results?\s+(and|&)\s+discussions?|findings\s+(and|&)\s+discussions?|hasil(\s+penelitian)?\s+dan\s+pembahasan)$/i],
    ['results', 'Results', /^((research\s+)?results?|findings?|hasil(\s+penelitian)?|temuan(\s+penelitian)?)$/i],
    ['discussion', 'Discussion', /^(discussions?|pembahasan|diskusi)$/i],
    ['conclusion', 'Conclusion', /^(conclusions?(\s+(and|&)\s+(suggestions?|recommendations?|implications?))?|concluding\s+remarks|kesimpulan(\s+dan\s+saran)?|simpulan(\s+dan\s+saran)?|penutup)$/i],
    ['acknowledgment', 'Acknowledgment', /^(acknowledge?ments?|ucapan\s+terima\s*kasih)$/i],
    ['references', 'References', /^(references?|reference\s+list|bibliography|daftar\s+(pustaka|referensi|rujukan)|referensi|rujukan)$/i]
  ];
  var NUM_TOP = /^(?:(?:bab|chapter|section)\s+)?(?:\d{1,2}|[IVXLC]{1,5})(?:[.)]|\s)\s*(?=\p{L})/u;
  var NUM_SUB = /^\d{1,2}\.\d/;
  function headingName(t) { return clean(t).replace(NUM_TOP, '').replace(/[:.]\s*$/, '').trim(); }
  function knownKey(t) {
    var n = headingName(t);
    for (var i = 0; i < KNOWN.length; i++) if (KNOWN[i][2].test(n)) return { key: KNOWN[i][0], label: KNOWN[i][1] };
    return null;
  }
  function isAllCaps(t) { var l = t.replace(/[^\p{L}]/gu, ''); return l.length >= 4 && l === l.toUpperCase() && l !== l.toLowerCase(); }
  function looksHeadingShape(p) {
    var t = clean(p.text);
    if (!t || p.cell) return false;
    var n = countWords(t);
    if (n > 10 || /[.;,]$/.test(t) && !/^\d+\.$/.test(t)) return false;
    if (/^(table|figure|fig\.|tabel|gambar|chart|grafik)\s*\d/i.test(t)) return false;
    return true;
  }
  function headingClass(p) {
    var t = clean(p.text);
    return { caps: isAllCaps(headingName(t)), num: NUM_TOP.test(t) && !NUM_SUB.test(t), style: p.style || null, outline: p.outline, bold: !!p.bold };
  }
  function sameClass(a, b) {
    if (a.outline === 0) return b.outline === 0 || (!!a.style && b.style === a.style);
    if (a.caps) return b.caps && (!a.num || b.num);
    if (a.num) return b.num;
    if (a.style && !/^(normal|[a-z]*body)/i.test(a.style)) return b.style === a.style;
    return b.bold === a.bold && false; // tanpa penanda yang jelas: hanya nama yang dikenal
  }

  // Heading utama: dicari contoh dari "Introduction", lalu semua paragraf pendek yang
  // berformat sama (ALL CAPS / nomor tingkat 1 / gaya Heading 1) dianggap heading utama.
  function detectHeadings(body) {
    var cands = [];
    body.forEach(function (p, i) {
      if (!looksHeadingShape(p)) return;
      var k = knownKey(p.text), c = headingClass(p);
      var signal = c.caps || c.num || c.outline === 0 || (p.style && /heading\s*1|^judul\s*1$|^1$/i.test(p.style)) || p.bold;
      if (k && signal) cands.push({ index: i, key: k.key, label: k.label, cls: c });
    });
    var intro = cands.filter(function (c) { return c.key === 'introduction'; })[0] || cands[0];
    if (!intro) return [];
    var heads = [];
    body.forEach(function (p, i) {
      if (i < intro.index || !looksHeadingShape(p)) return;
      var k = knownKey(p.text), c = headingClass(p);
      var match = sameClass(intro.cls, c);
      var endMatter = k && (k.key === 'references' || k.key === 'acknowledgment');
      if (!match && !endMatter && !(k && (c.caps || c.num || c.outline === 0 || c.bold) && (intro.cls.caps === c.caps))) return;
      if (!k && countWords(p.text) > 8) return;
      heads.push({ index: i, key: k ? k.key : 'other', label: k ? k.label : headingName(p.text), text: clean(p.text) });
    });
    return heads;
  }

  function sum(list, f) { return list.reduce(function (n, x) { return n + f(x); }, 0); }

  function sectionStats(body, heads, opts) {
    return heads.map(function (h, i) {
      var end = i + 1 < heads.length ? heads[i + 1].index : body.length;
      var ps = body.slice(h.index + 1, end);
      var tableWords = sum(ps.filter(function (p) { return p.cell; }), function (p) { return countWords(p.countText); });
      var textWords = sum(ps.filter(function (p) { return !p.cell; }), function (p) { return countWords(p.countText); });
      return { key: h.key, label: h.label, heading: h.text, index: h.index, end: end,
        words: textWords + (opts.countTables !== false ? tableWords : 0), textWords: textWords, tableWords: tableWords };
    });
  }

  // ---------- ABSTRAK & KATA KUNCI ----------
  var ABS_RE = /^\s*(abstract|abstrak)\b\s*[:.\-–—]?\s*/i;
  var KW_RE = /^\s*(key\s*words?|kata[\s-]*kunci|index\s+terms)\b\s*[:.\-–—]?\s*/i;
  var STOP_RE = /^\s*(this\s+is\s+an\s+open[\s-]access|open\s+access|copyright|©|\(c\)\s*\d{4}|received\b|revised\b|accepted\b|published\b|available\s+online|article\s+(history|info)|history\b|to\s+cite|how\s+to\s+cite|cite\s+this|doi\b|https?:\/\/doi|corresponding|correspondence|e-?mail\b|abstract\b|abstrak\b|key\s*words?\b|kata[\s-]*kunci\b|jel\b|introduction\b|pendahuluan\b)/i;

  function cellParas(body, key) { var out = []; body.forEach(function (p, i) { if (p.cell && p.cell.key === key) out.push(i); }); return out; }

  function findAbstract(body, limitIndex) {
    for (var i = 0; i < Math.min(body.length, limitIndex); i++) {
      var t = body[i].text, m = ABS_RE.exec(t.replace(/^[​﻿\s]+/, ''));
      if (!m || countWords(t) > 3 && clean(t.slice(m[0].length)).length < 1) continue;
      var lead = t.length - t.replace(/^[​﻿\s]+/, '').length;
      var head = { index: i, text: clean(t) };
      var parts = [], rest = t.slice(lead + m[0].length);
      if (countWords(rest) >= 5) parts.push({ index: i, offset: lead + m[0].length, text: rest, countText: body[i].countText.slice(lead + m[0].length) });
      var p = body[i], idx;
      function take(list) {
        for (var k = 0; k < list.length; k++) {
          var q = body[list[k]];
          if (!clean(q.text)) { if (parts.length && !p.cell) break; continue; }
          if (STOP_RE.test(clean(q.text)) || (!p.cell && looksHeadingShape(q) && (knownKey(q.text) || isAllCaps(q.text)))) break;
          parts.push({ index: list[k], offset: 0, text: q.text, countText: q.countText });
        }
      }
      if (p.cell) {
        idx = cellParas(body, p.cell.key).filter(function (j) { return j > i; });
        take(idx);
        if (!parts.length) {
          // Judul "ABSTRACT" di sel sendiri → isi ada di sel kolom yang sama pada baris berikutnya.
          var keys = {};
          body.forEach(function (q, j) { if (q.cell && q.cell.table === p.cell.table && q.cell.row > p.cell.row && q.cell.gridCol === p.cell.gridCol && !keys[q.cell.key] && clean(q.text)) keys[q.cell.key] = j; });
          var next = Object.keys(keys).sort(function (a, b) { return keys[a] - keys[b]; })[0];
          if (next) take(cellParas(body, next));
        }
      } else {
        idx = []; for (var j = i + 1; j < body.length; j++) idx.push(j);
        take(idx);
      }
      return { heading: head, parts: parts, words: sum(parts, function (x) { return countWords(x.countText); }) };
    }
    return null;
  }

  function findKeywords(body, limitIndex) {
    for (var i = 0; i < Math.min(body.length, limitIndex); i++) {
      var t = body[i].text, lead = t.length - t.replace(/^[​﻿\s]+/, '').length, m = KW_RE.exec(t.slice(lead));
      if (!m) continue;
      var segs = []; // { index, offset, text }
      var rest = t.slice(lead + m[0].length);
      if (clean(rest)) segs.push({ index: i, offset: lead + m[0].length, text: rest });
      var p = body[i], list = [];
      if (p.cell) list = cellParas(body, p.cell.key).filter(function (j) { return j > i; });
      else if (!segs.length) for (var j = i + 1; j < Math.min(body.length, i + 12); j++) list.push(j);
      for (var k = 0; k < list.length; k++) {
        var q = body[list[k]];
        if (!clean(q.text)) { if (segs.length) break; continue; }
        if (STOP_RE.test(clean(q.text)) || (!p.cell && looksHeadingShape(q) && knownKey(q.text))) break;
        segs.push({ index: list[k], offset: 0, text: q.text });
      }
      var joined = segs.map(function (s) { return s.text; }).join(' ');
      var sep = /;/.test(joined) ? ';' : ',';
      var items = [];
      segs.forEach(function (s) {
        var re = new RegExp('[^' + sep + ']+', 'g'), mm;
        while ((mm = re.exec(s.text))) {
          var raw = mm[0], core = raw.replace(/^[\s ​]+/, ''), off = mm.index + (raw.length - core.length);
          core = core.replace(/[\s .]+$/, '');
          if (!core || !/[\p{L}\p{N}]/u.test(core)) continue;
          items.push({ text: core, index: s.index, offset: s.offset + off });
        }
      });
      return { heading: { index: i, text: clean(t) }, items: items, separator: sep };
    }
    return null;
  }

  var collator = typeof Intl !== 'undefined' ? new Intl.Collator('en', { sensitivity: 'base', numeric: true }) : null;
  function cmp(a, b) {
    var x = a.replace(/^[^\p{L}\p{N}]+/u, ''), y = b.replace(/^[^\p{L}\p{N}]+/u, '');
    return collator ? collator.compare(x, y) : x.toLowerCase().localeCompare(y.toLowerCase());
  }

  // ---------- JUDUL & BYLINE ----------
  var META_RE = /\b(vol(ume)?\.?\s*\d|issn|e-issn|p-issn|https?:\/\/|doi|journal\s+of|jurnal\s+\p{L}+\s*(vol|issn))/iu;
  function findTitle(body, stopIndex) {
    var cands = [];
    for (var i = 0; i < Math.min(body.length, stopIndex); i++) {
      var p = body[i], t = clean(p.text);
      if (!t || p.cell || countWords(t) < 3 || META_RE.test(t) || ABS_RE.test(t) || KW_RE.test(t)) continue;
      if (/@|^\s*[\d,*†]+\s*\p{Lu}|\b(corresponding|correspondence)\b/iu.test(t)) continue;
      cands.push({ index: i, text: t, size: p.size || 0, bold: p.bold });
      if (cands.length >= 4) break;
    }
    if (!cands.length) return null;
    var best = cands[0];
    cands.forEach(function (c) { if (c.size > best.size) best = c; });
    return best;
  }

  var AFFIL_RE = /@|\b(universit\w*|institut\w*|faculty|fakultas|department|departemen|laborator\w*|school|college|academy|akademi|politeknik|polytechnic|center|centre|ministry|kementerian|correspond\w*|e-?mail|orcid|indonesia|india|china|korea|malaysia)\b/i;
  function bylineNames(body, from, to) {
    var names = [];
    for (var i = from; i < to; i++) {
      var t = clean(body[i].text);
      if (t && (STOP_RE.test(t) || /^(article\s+(info|history|information)|info\s+artikel)\b/i.test(t))) break;
      if (!t || AFFIL_RE.test(t) || ABS_RE.test(t) || KW_RE.test(t) || STOP_RE.test(t) || isAllCaps(t) && countWords(t) <= 3 && /article|info|history/i.test(t)) continue;
      t.split(/\s*(?:,|;|\s&\s|\band\b|\bdan\b)\s*/).forEach(function (raw) {
        var n = raw.replace(/[\d*†‡§¶#,]+/g, ' ').replace(/\s+/g, ' ').trim();
        var toks = n.split(' ').filter(Boolean);
        if (toks.length < 1 || toks.length > 5) return;
        if (!toks.every(function (w) { return /^\p{Lu}[\p{L}'’.-]*$/u.test(w); })) return;
        names.push({ name: n, tokens: toks, index: i });
      });
    }
    return names;
  }

  // ---------- HOW TO CITE ----------
  var HTC_RE = /(how\s+to\s+cite(?:\s+this\s+(?:article|paper))?|to\s+cite\s+this\s+(?:article|paper)|cite\s+this\s+(?:article|paper)(?:\s+as)?|cara\s+(?:mengutip|menyitasi|sitasi)(?:\s+artikel\s+ini)?|sitasi\s+artikel\s+ini)\s*:?\s*/i;
  var DASH_CLASS = '\\-\\u2010\\u2011\\u2012\\u2013\\u2014\\u2212';

  function findHowToCite(model) {
    var stories = [];
    if (model.headerFooter) {
      [model.firstPageFooter, model.firstPageHeader, 'footer:default', 'header:default', 'footer:first', 'header:first', 'footer:even', 'header:even'].forEach(function (k) {
        if (model.headerFooter[k] && stories.every(function (s) { return s.story !== k; })) stories.push({ story: k, paras: model.headerFooter[k] });
      });
    }
    stories.unshift({ story: 'body', paras: model.body });
    if (model.textboxes && model.textboxes.length) stories.push({ story: 'textbox', paras: model.textboxes });
    for (var s = 0; s < stories.length; s++) {
      var ps = stories[s].paras;
      for (var i = 0; i < ps.length; i++) {
        var m = HTC_RE.exec(ps[i].text);
        if (!m) continue;
        var start = m.index + m[0].length, text = ps[i].text.slice(start), extra = [];
        for (var j = i + 1; j < ps.length && clean(text).length < 40; j++) { if (clean(ps[j].text)) { text += ' ' + ps[j].text; extra.push(j); } }
        return { story: stories[s].story, index: i, offset: start, label: clean(m[0]), text: text, paraText: ps[i].text, more: extra };
      }
    }
    return null;
  }

  function parseCitation(text, manuscriptTitle, journalHint) {
    var t = String(text || '').replace(/[\s ]+/g, ' ').trim();
    var out = { raw: t };
    var ym = /\((\d{4})[a-z]?(?:,[^)]*)?\)\s*\.?\s*/.exec(t);
    if (!ym) return out;
    out.year = ym[1];
    out.authorsText = t.slice(0, ym.index).trim();
    var rest = t.slice(ym.index + ym[0].length);
    var doi = /\b(10\.\d{4,9}\/[^\s"<>]+)/.exec(rest);
    if (doi) { out.doi = doi[1].replace(/[.,;]+$/, ''); out.doiPrefix = clean(rest.slice(0, doi.index).replace(/^[\s\S]*?(?=(doi|https?:)[^\s]*\s*$)/i, '')); }
    var volRe = new RegExp(',\\s*(\\d+)\\s*(?:\\(\\s*([^)]+?)\\s*\\))?\\s*,\\s*(?:pp?\\.\\s*)?([A-Za-z]?\\d+)\\s*([' + DASH_CLASS + ']+)\\s*([A-Za-z]?\\d+)');
    var vm = volRe.exec(rest), cut;
    if (vm) {
      out.volume = vm[1]; out.issue = vm[2] || null; out.pageStart = vm[3]; out.pageDash = vm[4]; out.pageEnd = vm[5];
      out.pagesText = rest.substr(vm.index, vm[0].length).replace(/^,\s*\d+\s*(?:\([^)]*\))?\s*,\s*/, '');
      out.pagesTextFull = vm[3] + vm[4] + vm[5];
      cut = vm.index;
    } else {
      var art = /,\s*(\d+)\s*(?:\(\s*([^)]+?)\s*\))?\s*,\s*(?:article\s+|art\.\s*|e)?([A-Za-z]?\d+)\s*\.(?!\d)/i.exec(rest);
      var pp = /\bpp?\.\s*([A-Za-z]?\d+)\s*([\-‐-—−]+)\s*([A-Za-z]?\d+)/.exec(rest);
      if (pp) { out.pageStart = pp[1]; out.pageDash = pp[2]; out.pageEnd = pp[3]; out.pagesTextFull = pp[1] + pp[2] + pp[3]; cut = rest.lastIndexOf(',', pp.index); }
      else if (art) { out.volume = art[1]; out.issue = art[2] || null; out.articleNumber = art[3]; cut = art.index; }
    }
    var pre = (cut != null && cut >= 0 ? rest.slice(0, cut) : rest).trim();
    // Judul vs nama jurnal: titik potong di setiap ". " / "? " / "! "; dipilih yang judulnya paling
    // mirip judul naskah (atau yang nama jurnalnya cocok dengan header), bukan potongan pertama.
    var splits = [], re = /([.?!])\s+/g, sm;
    while ((sm = re.exec(pre))) splits.push({ at: sm.index, keep: sm[1] === '?' || sm[1] === '!' ? 1 : 0, next: sm.index + sm[0].length });
    var best = null;
    splits.forEach(function (sp) {
      var title = pre.slice(0, sp.at + sp.keep).trim(), journal = pre.slice(sp.next).replace(/[,\s]+$/, '').trim();
      if (!title || !journal) return;
      var score = 0;
      if (manuscriptTitle) score += similarity(title, manuscriptTitle) * 2;
      if (journalHint) score += similarity(journal, journalHint);
      if (!manuscriptTitle && !journalHint) score = sp.at; // tanpa petunjuk: potongan terakhir
      if (!best || score > best.score) best = { score: score, title: title, journal: journal, offset: 0 };
    });
    if (best) {
      out.title = best.title; out.journal = best.journal;
      var titleAt = t.indexOf(best.title, ym.index);
      out.titleOffset = titleAt;
    } else if (pre) out.title = pre.replace(/[.\s]+$/, '');
    return out;
  }

  // Dice dua-gram kata, 0..1
  function similarity(a, b) {
    var x = words(a), y = words(b);
    if (!x.length || !y.length) return 0;
    var bag = {}, hit = 0;
    x.forEach(function (w) { bag[w] = (bag[w] || 0) + 1; });
    y.forEach(function (w) { if (bag[w]) { bag[w]--; hit++; } });
    return 2 * hit / (x.length + y.length);
  }

  // Perbedaan kata (LCS) untuk ditampilkan: [{op:'eq'|'del'|'ins', text}]
  function diffWords(a, b) {
    var A = clean(a).split(' '), B = clean(b).split(' ');
    function k(w) { return fold(w).toLowerCase().replace(/[^\p{L}\p{N}'-]/gu, ''); }
    var n = A.length, m = B.length, dp = [];
    for (var i = 0; i <= n; i++) { dp.push(new Array(m + 1).fill(0)); }
    for (i = n - 1; i >= 0; i--) for (var j = m - 1; j >= 0; j--) dp[i][j] = k(A[i]) === k(B[j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    var out = []; i = 0; j = 0;
    function push(op, w) { var l = out[out.length - 1]; if (l && l.op === op) l.text += ' ' + w; else out.push({ op: op, text: w }); }
    while (i < n && j < m) {
      if (k(A[i]) === k(B[j])) { push('eq', B[j]); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) push('del', A[i++]);
      else push('ins', B[j++]);
    }
    while (i < n) push('del', A[i++]);
    while (j < m) push('ins', B[j++]);
    return out;
  }

  // Title Case: sebagian besar kata isi (selain kata pertama & sesudah titik dua) berhuruf kapital.
  var MINOR = /^(a|an|the|and|or|but|nor|for|of|in|on|at|to|by|with|from|as|via|vs|per|into|over|than|yang|dan|di|ke|dari|untuk|pada|dengan|dalam)$/i;
  function titleCaseRatio(t) {
    var toks = clean(t).split(' '), caps = 0, total = 0;
    toks.forEach(function (w, i) {
      var prev = toks[i - 1] || '';
      if (i === 0 || /[:?!]$/.test(prev)) return;
      var core = w.replace(/^[^\p{L}]+/u, '');
      if (!core || MINOR.test(core.replace(/[^\p{L}]/gu, '')) || /^\p{Lu}{2,}/u.test(core)) return;
      total++;
      if (/^\p{Lu}/u.test(core)) caps++;
    });
    return total ? caps / total : 0;
  }

  function citationAuthors(text) {
    var s = String(text || '').replace(/\s*(?:&|\band\b|\bdan\b)\s*/g, ', ').replace(/\s+/g, ' ').trim().replace(/[.,\s]+$/, '.');
    var toks = s.split(/\s*,\s*/).filter(Boolean), out = [];
    for (var i = 0; i < toks.length; i++) {
      var surname = toks[i].trim(), init = (toks[i + 1] || '').trim();
      if (/^(?:\p{Lu}\.?[\s-]*){1,4}\.?$/u.test(init) && !/^(?:\p{Lu}\.?[\s-]*){1,4}\.?$/u.test(surname)) { out.push({ surname: surname, initials: init.replace(/[^\p{Lu}]/gu, '') }); i++; }
      else if (/^\.\.\.|…/.test(surname)) continue;
      else return null; // bukan pola "Nama, I." → tidak dibandingkan
    }
    return out;
  }

  function hfText(model) {
    var hf = model.headerFooter || {};
    return Object.keys(hf).map(function (k) { return (hf[k] || []).map(function (p) { return p.text; }).join('\n'); }).join('\n');
  }

  // ---------- ATURAN ----------
  function defaultRules() {
    return {
      abstract: { min: 150, max: 250 },
      introduction: { min: null, max: null },
      results_discussion: { min: null, max: null },
      conclusion: { min: null, max: null },
      total: { min: null, max: null },
      keywords: { min: 3, max: 6 },
      countTables: true, totalIncludesNotes: true
    };
  }
  function within(n, r) {
    if (!r || (r.min == null && r.max == null)) return 'info';
    if (r.min != null && n < r.min) return 'low';
    if (r.max != null && n > r.max) return 'high';
    return 'ok';
  }
  function limitText(r) {
    if (!r || (r.min == null && r.max == null)) return '';
    if (r.min != null && r.max != null) return r.min + '–' + r.max;
    return r.min != null ? '≥ ' + r.min : '≤ ' + r.max;
  }

  // ---------- ANALISIS ----------
  // opts: { rules, pages: { count, source: 'layout'|'file'|'estimate' } }
  function analyze(model, opts) {
    var o = opts || {}, rules = Object.assign(defaultRules(), o.rules || {});
    var body = model.body || [];
    var checks = [];
    function add(c) { checks.push(c); return c; }

    var heads = detectHeadings(body);
    var firstHead = heads.length ? heads[0].index : body.length;
    var sections = sectionStats(body, heads, rules);
    function sec(key) { return sections.filter(function (s) { return s.key === key; })[0] || null; }

    // ----- jumlah kata -----
    var abs = findAbstract(body, firstHead);
    var counts = [];
    function countRow(key, label, n, loc, extra) {
      var r = rules[key], st = within(n, r);
      counts.push(Object.assign({ key: key, label: label, words: n, limit: limitText(r), state: st, location: loc }, extra || {}));
    }
    if (abs) countRow('abstract', 'Abstract', abs.words, { story: 'body', index: abs.heading.index, text: abs.heading.text }, { parts: abs.parts.length });
    else counts.push({ key: 'abstract', label: 'Abstract', words: null, limit: limitText(rules.abstract), state: 'missing' });
    [['introduction', 'Introduction'], ['results_discussion', 'Results and Discussion'], ['conclusion', 'Conclusion']].forEach(function (k) {
      var s = sec(k[0]);
      if (!s && k[0] === 'results_discussion' && (sec('results') || sec('discussion'))) {
        var r = sec('results'), d = sec('discussion');
        var n = (r ? r.words : 0) + (d ? d.words : 0);
        countRow(k[0], k[1], n, { story: 'body', index: (r || d).index, text: (r || d).heading }, { combined: [r, d].filter(Boolean).map(function (x) { return x.label + ' ' + x.words; }).join(' + '), tableWords: (r ? r.tableWords : 0) + (d ? d.tableWords : 0) });
        return;
      }
      if (s) countRow(k[0], k[1], s.words, { story: 'body', index: s.index, text: s.heading }, { heading: s.heading, tableWords: s.tableWords });
      else counts.push({ key: k[0], label: k[1], words: null, limit: limitText(rules[k[0]]), state: 'missing' });
    });
    var main = sum(body, function (p) { return countWords(p.countText); });
    var tb = sum(model.textboxes || [], function (p) { return countWords(p.countText); });
    var fn = sum(model.footnotes || [], function (n) { return sum(n.paragraphs, function (p) { return countWords(p.countText); }); });
    var en = sum(model.endnotes || [], function (n) { return sum(n.paragraphs, function (p) { return countWords(p.countText); }); });
    var total = main + (rules.totalIncludesNotes !== false ? tb + fn + en : 0);
    countRow('total', 'Total naskah', total, null, { main: main, textboxes: tb, footnotes: fn, endnotes: en });

    counts.forEach(function (c) {
      var st = c.state === 'ok' ? 'ok' : c.state === 'info' ? 'info' : c.state === 'missing' ? (c.key === 'total' ? 'info' : 'warn') : 'bad';
      var title = c.state === 'missing' ? c.label + ' tidak ditemukan' : c.label + ': ' + c.words + ' kata' + (c.limit ? ' (batas ' + c.limit + ')' : '');
      var detail = c.state === 'missing' ? (c.key === 'abstract' ? 'Tidak ada judul "Abstract"/"Abstrak" sebelum heading pertama.' : 'Heading "' + c.label + '" tidak terdeteksi sebagai heading utama.')
        : c.state === 'low' ? 'Kurang ' + (rulesFor(c.key).min - c.words) + ' kata dari batas minimal.' : c.state === 'high' ? 'Lebih ' + (c.words - rulesFor(c.key).max) + ' kata dari batas maksimal.' : '';
      if (c.combined) detail = (detail ? detail + ' ' : '') + 'Dijumlahkan dari ' + c.combined + '.';
      if (c.key === 'total') detail = (detail ? detail + ' ' : '') + 'Badan naskah ' + c.main + (c.textboxes ? ' + kotak teks ' + c.textboxes : '') + (c.footnotes ? ' + catatan kaki ' + c.footnotes : '') + (c.endnotes ? ' + catatan akhir ' + c.endnotes : '') + '. Header & footer tidak dihitung (sama seperti Word).';
      else if (c.tableWords && c.state !== 'missing') detail = (detail ? detail + ' ' : '') + (rules.countTables !== false ? 'Termasuk ' + c.tableWords + ' kata di dalam tabel.' : c.tableWords + ' kata di dalam tabel tidak dihitung.');
      add({ id: 'words:' + c.key, group: 'words', status: st, title: title, detail: detail, value: c.words, limit: c.limit, location: c.location });
    });
    function rulesFor(k) { return rules[k] || {}; }

    // ----- judul & how to cite -----
    var absIndex = abs ? abs.heading.index : firstHead;
    var title = findTitle(body, Math.min(absIndex, firstHead));
    var htc = findHowToCite(model);
    var hft = hfText(model);
    var journalHint = (/(journal\s+of[^\n|]*?|jurnal\s+[^\n|]*?)(?=\s*(?:\||\n|e-issn|p-issn|issn|vol\.|$))/i.exec(hft) || [])[1] || null;
    if (journalHint) journalHint = clean(journalHint).replace(/\s*\d+\s*$/, '');
    var cite = htc ? parseCitation(htc.text, title && title.text, journalHint) : null;
    var pageInfo = actualPages(model, o.pages);

    if (!htc) add({ id: 'cite:missing', group: 'cite', status: 'warn', title: 'How to cite tidak ditemukan', detail: 'Tidak ada baris "How to cite" / "To cite this article" di badan naskah, kotak teks, maupun header/footer halaman pertama.' });
    else {
      var where = htc.story === 'body' ? (body[htc.index].cell ? 'tabel halaman pertama' : 'badan naskah') : /footer/.test(htc.story) ? 'footer ' + (/first/.test(htc.story) ? 'halaman pertama' : 'halaman') : /header/.test(htc.story) ? 'header' : 'kotak teks';
      var citeLoc = { story: htc.story, index: htc.index, text: htc.paraText };
      // --- rentang halaman ---
      if (cite && cite.pageStart) {
        var cs = parseInt(cite.pageStart.replace(/\D/g, ''), 10), ce = parseInt(cite.pageEnd.replace(/\D/g, ''), 10);
        var want = pageInfo.count ? pageInfo.start + '–' + pageInfo.end : null;
        var st = 'ok', msgs = [];
        if (!pageInfo.count) { st = 'info'; msgs.push('Jumlah halaman dokumen tidak bisa dibaca, jadi rentang tidak bisa dibandingkan.'); }
        else {
          if (cs !== pageInfo.start) msgs.push('Halaman awal tertulis ' + cs + ', padahal penomoran halaman dokumen dimulai dari ' + pageInfo.start + (pageInfo.startSet ? '' : ' (nomor awal belum diatur)') + '.');
          if (ce !== pageInfo.end) msgs.push('Halaman akhir tertulis ' + ce + ', padahal dokumen ' + pageInfo.count + ' halaman sehingga berakhir di halaman ' + pageInfo.end + '.');
          if (msgs.length) st = pageInfo.source === 'estimate' ? 'warn' : 'bad';
          if (!pageInfo.startSet && ce - cs + 1 === pageInfo.count) {
            msgs = ['Jumlah halaman sudah sama (' + pageInfo.count + '), tetapi penomoran halaman dokumen masih mulai dari 1. Atur di Insert › Page Number › Format Page Numbers › Start at: ' + cs + '.'];
            st = 'warn';
          }
        }
        var dashOk = /^–$/.test(cite.pageDash);
        if (st === 'ok' && msgs.length === 0) msgs.push('Sesuai: dokumen ' + pageInfo.count + ' halaman, halaman ' + pageInfo.start + '–' + pageInfo.end + '.');
        if (!dashOk) msgs.push('Rentang halaman memakai "' + cite.pageDash + '"; APA 7 memakai en dash (–).');
        var fixTo = (st === 'bad' || st === 'warn') && want && pageInfo.startSet !== false ? want : (!dashOk ? cs + '–' + ce : null);
        if (!pageInfo.startSet && ce - cs + 1 === pageInfo.count) fixTo = !dashOk ? cs + '–' + ce : null;
        add({ id: 'cite:pages', group: 'cite', status: st === 'ok' && !dashOk ? 'warn' : st, title: 'Rentang halaman: tertulis ' + cite.pagesTextFull + (want ? ', dokumen ' + want : ''),
          detail: msgs.join(' '), pages: pageInfo, location: citeLoc, where: where,
          fix: fixTo && fixTo !== cite.pagesTextFull ? { story: htc.story, index: htc.index, find: cite.pagesTextFull, replace: fixTo, label: 'Ubah jadi ' + fixTo } : null });
      } else if (cite && cite.articleNumber) {
        add({ id: 'cite:pages', group: 'cite', status: 'info', title: 'How to cite memakai nomor artikel (' + cite.articleNumber + ')', detail: 'Tidak ada rentang halaman yang perlu dibandingkan.', location: citeLoc, where: where, pages: pageInfo });
      } else {
        add({ id: 'cite:pages', group: 'cite', status: 'bad', title: 'Rentang halaman tidak ditemukan di How to cite', detail: 'Pola "Volume(Nomor), awal–akhir" tidak ditemukan. Dokumen ini ' + (pageInfo.count ? pageInfo.count + ' halaman (' + pageInfo.start + '–' + pageInfo.end + ').' : 'jumlah halamannya tidak terbaca.'), location: citeLoc, where: where, pages: pageInfo });
      }
      // "Page 1 of 226" di footer: angka sesudah "of" yang diketik manual harus = halaman akhir.
      var ofm = /\b(?:page|halaman|hal\.)\s*\d+\s*(?:of|dari)\s*(\d+)/i.exec(hft);
      if (ofm && pageInfo.count && parseInt(ofm[1], 10) !== pageInfo.end) {
        add({ id: 'cite:pageof', group: 'cite', status: 'warn', title: 'Footer "' + clean(ofm[0]) + '" tidak sesuai', detail: 'Angka total di footer ' + ofm[1] + ', padahal halaman terakhir dokumen ' + pageInfo.end + '.', location: null });
      }

      // --- judul ---
      if (!title) add({ id: 'cite:title', group: 'cite', status: 'warn', title: 'Judul naskah tidak terdeteksi', detail: 'Judul tidak bisa dibandingkan dengan How to cite.', location: citeLoc });
      else if (!cite || !cite.title) add({ id: 'cite:title', group: 'cite', status: 'warn', title: 'Judul di How to cite tidak terbaca', detail: 'Pastikan susunannya: Penulis (Tahun). Judul. Nama Jurnal, Volume(Nomor), halaman.', location: citeLoc });
      else {
        var a = fold(cite.title), b = fold(title.text);
        var sim = Math.round(similarity(cite.title, title.text) * 100);
        var res = { id: 'cite:title', group: 'cite', location: citeLoc, manuscript: title.text, cited: cite.title, similarity: sim };
        if (a === b) { res.status = 'ok'; res.title = 'Judul sesuai dengan judul naskah'; res.detail = ''; }
        else if (a.toLowerCase() === b.toLowerCase()) {
          var tcC = titleCaseRatio(cite.title), tcM = titleCaseRatio(title.text);
          res.status = 'warn'; res.title = 'Judul sama, tetapi kapitalisasinya berbeda';
          res.detail = (tcC > 0.6 && tcM < 0.4 ? 'How to cite memakai Title Case, sedangkan judul naskah sentence case. APA 7 menulis judul artikel dengan sentence case — samakan dengan judul naskah.'
            : tcM > 0.6 && tcC < 0.4 ? 'Judul naskah memakai Title Case, How to cite sentence case. Periksa gaya yang diminta jurnal.' : 'Huruf kapital pada beberapa kata berbeda.');
          res.diff = diffCase(cite.title, title.text);
          if (title.text.length <= 250 && cite.title.length <= 250) res.fix = { story: htc.story, index: htc.index, find: cite.title, replace: cleanTitleForCite(title.text), label: 'Samakan dengan judul naskah' };
        } else {
          res.status = sim >= 50 ? 'bad' : 'bad';
          res.title = sim >= 50 ? 'Judul berbeda dari judul naskah (' + sim + '% kata sama)' : 'Judul di How to cite tidak cocok dengan judul naskah';
          res.detail = 'Kata yang dicoret hanya ada di How to cite; kata hijau ada di judul naskah.';
          res.diff = diffWords(cite.title, title.text);
          if (sim >= 50 && title.text.length <= 250 && cite.title.length <= 250) res.fix = { story: htc.story, index: htc.index, find: cite.title, replace: cleanTitleForCite(title.text), label: 'Samakan dengan judul naskah' };
        }
        add(res);
      }

      // --- metadata lain ---
      var meta = [];
      var vm = /vol(?:ume)?\.?\s*(\d+)\s*,?\s*(?:no|nomor|issue|iss)\.?\s*(\d+)/i.exec(hft);
      if (cite && cite.volume && vm && (vm[1] !== cite.volume || (cite.issue && vm[2] !== String(cite.issue).replace(/\D/g, '')))) meta.push({ bad: true, text: 'Volume/nomor tertulis ' + cite.volume + '(' + (cite.issue || '-') + '), header menyebut Vol. ' + vm[1] + ' No. ' + vm[2] + '.' });
      var hy = /vol[^\n]{0,40}?((?:19|20)\d{2})/i.exec(hft);
      if (cite && cite.year && hy && hy[1] !== cite.year) meta.push({ bad: true, text: 'Tahun tertulis ' + cite.year + ', header menyebut ' + hy[1] + '.' });
      var hdoi = /\b(10\.\d{4,9}\/[^\s"<>]+)/.exec(hft);
      if (cite && cite.doi && hdoi && hdoi[1].replace(/[.,;]+$/, '').toLowerCase() !== cite.doi.toLowerCase()) meta.push({ bad: true, text: 'DOI berbeda dengan DOI di header (' + hdoi[1].replace(/[.,;]+$/, '') + ').' });
      var vi = cite && cite.doi && /v(\d+)i(\d+)/i.exec(cite.doi);
      if (vi && cite.volume && (vi[1] !== cite.volume || (cite.issue && vi[2] !== String(cite.issue).replace(/\D/g, '')))) meta.push({ bad: true, text: 'DOI memuat v' + vi[1] + 'i' + vi[2] + ', tetapi volume/nomor tertulis ' + cite.volume + '(' + (cite.issue || '-') + ').' });
      if (cite && cite.journal && journalHint && similarity(cite.journal, journalHint) < 0.8) meta.push({ bad: true, text: 'Nama jurnal "' + cite.journal + '" berbeda dengan header ("' + journalHint + '").' });
      if (cite && cite.doi && /\bdoi\s*[.:]\s*https?:/i.test(htc.text)) meta.push({ bad: false, text: 'Hapus awalan "' + (/\bdoi\s*[.:]/i.exec(htc.text) || ['Doi.'])[0] + '" sebelum tautan DOI — APA 7 cukup menulis https://doi.org/….' });
      if (cite && !cite.doi) meta.push({ bad: false, text: 'Tidak ada DOI di How to cite.' });
      // penulis vs byline
      var authors = cite ? citationAuthors(cite.authorsText) : null;
      var byline = title ? bylineNames(body, title.index + 1, Math.min(absIndex, firstHead)) : [];
      if (authors && authors.length && byline.length) {
        var problems = [];
        if (authors.length !== byline.length) problems.push('How to cite menyebut ' + authors.length + ' penulis, byline ' + byline.length + ' (' + byline.map(function (x) { return x.name; }).join(', ') + ').');
        authors.forEach(function (au, i) {
          var by = byline[i]; if (!by) return;
          var sur = fold(au.surname).toLowerCase(), last = fold(by.tokens[by.tokens.length - 1]).toLowerCase(), firstTok = fold(by.tokens[0]).toLowerCase();
          if (sur !== last && sur !== firstTok && fold(by.name).toLowerCase().indexOf(sur) < 0) {
            var elsewhere = byline.findIndex(function (x) { return fold(x.name).toLowerCase().indexOf(sur) >= 0; });
            problems.push(elsewhere >= 0 ? au.surname + ' tertulis di urutan ' + (i + 1) + ', di byline urutan ' + (elsewhere + 1) + '.' : au.surname + ' tidak ada di byline.');
            return;
          }
          var given = by.tokens.filter(function (w) { return fold(w).toLowerCase() !== sur; });
          if (au.initials && given.length && given[0][0].toUpperCase() !== au.initials[0]) problems.push('Inisial ' + au.surname + ' tertulis ' + au.initials[0] + '., byline "' + by.name + '".');
        });
        if (problems.length) meta.push({ bad: true, text: 'Penulis: ' + problems.join(' ') });
      }
      add({ id: 'cite:meta', group: 'cite', status: meta.some(function (x) { return x.bad; }) ? 'bad' : meta.length ? 'warn' : 'ok',
        title: meta.length ? 'Data lain di How to cite perlu dicek' : 'Penulis, tahun, volume/nomor, jurnal & DOI konsisten',
        detail: meta.map(function (x) { return x.text; }).join(' '), items: meta, location: citeLoc,
        checked: { authors: !!(authors && byline.length), header: !!hft.trim() } });
    }

    // ----- kata kunci -----
    var kw = findKeywords(body, firstHead);
    if (!kw || !kw.items.length) add({ id: 'kw', group: 'keywords', status: 'warn', title: 'Kata kunci tidak ditemukan', detail: 'Pastikan ada "Keywords:" / "Kata kunci:" sebelum Introduction.' });
    else {
      var list = kw.items.map(function (x) { return x.text; });
      var sorted = list.slice().sort(cmp);
      var firstBad = -1;
      for (var q = 1; q < list.length; q++) if (cmp(list[q - 1], list[q]) > 0) { firstBad = q; break; }
      var nst = within(list.length, rules.keywords);
      var kwRes = { id: 'kw', group: 'keywords', location: { story: 'body', index: kw.heading.index, text: kw.heading.text }, items: list, sorted: sorted };
      if (firstBad < 0) { kwRes.status = 'ok'; kwRes.title = list.length + ' kata kunci, sudah alfabetis'; kwRes.detail = list.join('; ') + '.'; }
      else {
        kwRes.status = 'bad'; kwRes.title = 'Kata kunci belum alfabetis';
        kwRes.detail = '"' + list[firstBad] + '" seharusnya sebelum "' + list[firstBad - 1] + '". Urutan yang benar: ' + sorted.join('; ') + '.';
        kwRes.fix = { kind: 'keywords', items: kw.items, sorted: sorted, label: 'Urutkan kata kunci' };
      }
      add(kwRes);
      if (nst === 'low' || nst === 'high') add({ id: 'kw:count', group: 'keywords', status: 'warn', title: 'Jumlah kata kunci ' + list.length + ' (batas ' + limitText(rules.keywords) + ')', detail: nst === 'low' ? 'Tambahkan kata kunci.' : 'Kurangi kata kunci.', location: kwRes.location });
    }

    return {
      checks: checks, counts: counts, sections: sections, title: title, howToCite: htc, citation: cite, pages: pageInfo,
      keywords: kw, abstract: abs, rules: rules,
      summary: { bad: checks.filter(function (c) { return c.status === 'bad'; }).length, warn: checks.filter(function (c) { return c.status === 'warn'; }).length }
    };
  }

  function cleanTitleForCite(t) { return clean(t).replace(/[.\s]+$/, ''); }
  function diffCase(a, b) {
    var A = clean(a).split(' '), B = clean(b).split(' ');
    if (A.length !== B.length) return diffWords(a, b);
    return B.map(function (w, i) { return A[i] === w ? { op: 'eq', text: w } : { op: 'case', text: w, was: A[i] }; });
  }

  // Halaman sebenarnya: nomor awal dari Format Page Numbers bagian pertama + jumlah halaman.
  function actualPages(model, given) {
    var start = model.section && model.section.pgStart != null ? model.section.pgStart : 1;
    var startSet = !!(model.section && model.section.pgStart != null);
    var count = null, source = null;
    if (given && given.count) { count = given.count; source = given.source || 'layout'; }
    else if (model.app && model.app.pages) { count = model.app.pages; source = 'file'; }
    else if (model.layoutHints) { count = model.layoutHints.lastRenderedPageBreaks + model.layoutHints.pageBreaks + 1; source = 'estimate'; }
    return { start: start, startSet: startSet, count: count, end: count ? start + count - 1 : null, source: source, restarts: model.section ? model.section.restarts : 0 };
  }

  var api = {
    countWords: countWords, analyze: analyze, defaultRules: defaultRules, detectHeadings: detectHeadings,
    parseCitation: parseCitation, similarity: similarity, diffWords: diffWords, titleCaseRatio: titleCaseRatio,
    findAbstract: findAbstract, findKeywords: findKeywords, findTitle: findTitle, findHowToCite: findHowToCite, compareKeywords: cmp, fold: fold
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TemplateCore = api;
})(typeof window !== 'undefined' ? window : this);
