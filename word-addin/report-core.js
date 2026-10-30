/* ============================================================
   REPORT CORE — "laporan dalam bentuk komentar" (setara fitur web "Naskah Asli dengan
   Highlight + Komentar Word" di cek-sitasi/upload.js).

   Mengubah hasil pemeriksaan (saran Validator, Cek Template, dan Metadata: tahun referensi
   di luar rentang, DOI bermasalah / tanpa DOI, status Scopus) menjadi daftar komentar yang ditempel di lokasinya masing-masing:
     { paraIndex, offset, length, text, occurrence | whole | head/tail, color, comment }
   Bridge Word lalu menyorot rentang itu dan menyisipkan komentar Word asli
   (Tinjau › Komentar), atau DocxCommenter menuliskannya ke salinan .docx.

   Bagian yang dikelola Zotero/Mendeley tidak pernah disorot dan komentarnya tidak
   ditempel DI DALAM field-nya: komentar dipindah ke kata tepat sebelum/sesudah sitasi
   (atau ke judul daftar pustaka), supaya field pengelola sitasi tetap utuh.

   Murni: tanpa Office.js. Diuji di test/report.test.js.
   ============================================================ */
(function (root) {
  'use strict';
  var isNode = typeof module !== 'undefined' && module.exports;
  var AC = isNode ? require('./addin-core.js') : root.AddinCore;
  var CM = isNode ? require('./citation-managers.js') : root.CitationManagers;

  var PREFIX = '[Cek Sitasi]';
  var TEMPLATE_PREFIX = '[Cek Sitasi · Template]';
  var METADATA_PREFIX = '[Cek Sitasi · Metadata]';
  var TERM_PREFIX = '[Cek Sitasi · Istilah]';
  // Skema warna sorotan (sama untuk semua laporan):
  //   MERAH  = perlu diperbaiki · BIRU (Pirus/Turquoise, biru muda) = saran · KUNING = metadata
  //   (DOI bermasalah, tanpa DOI, tahun di luar rentang, Scopus). "Blue" bawaan Word terlalu gelap
  //   untuk teks hitam, jadi biru memakai Turquoise dari palet sorotan Word.
  var COLOR = { error: 'Red', suggestion: 'Turquoise', metadata: 'Yellow', year: 'Yellow', doi: 'Yellow', doiBad: 'Yellow', nodoi: 'Yellow', scopus: 'Yellow' };
  var LEGEND = 'Warna sorotan: merah = perlu diperbaiki, biru = saran, kuning = metadata (DOI, Scopus, tahun).';
  var RANK = { Red: 3, Yellow: 2, Turquoise: 1 };
  var MAX_PART = 700;

  function clip(t, n) { t = String(t || '').replace(/\s+/g, ' ').trim(); return t.length > n ? t.slice(0, n - 1) + '…' : t; }
  function sentence(t) { t = String(t || '').trim(); return t && !/[.!?…]$/.test(t) ? t + '.' : t; }

  function suggestionComment(s) {
    var label = s.severity === 'error' ? 'Perlu diperbaiki' : 'Saran';
    var parts = [sentence(s.title)];
    if (s.reasons && s.reasons.length > 1) parts.push('Temuan: ' + s.reasons.map(function (r) { return r.title; }).join('; ') + '.');
    if (s.description) parts.push(sentence(s.description));
    if (s.replacement) parts.push('Saran perbaikan: ' + clip(s.replacement, 420));
    else if (s.hint) parts.push('Saran perbaikan: ' + clip(s.hint, 300));
    return label + ' — ' + clip(parts.join(' '), MAX_PART);
  }

  // Rentang [offset, offset+length) dipindah keluar dari field/content control pengelola sitasi.
  function anchorOutside(st, managed, item) {
    var span = CM.managedAt(managed, item.paraIndex, item.offset, item.offset + item.length);
    if (!span) return item;
    var text = st.paragraphs[item.paraIndex].text;
    var quoted = clip(text.slice(Math.max(span.start, item.offset), Math.min(span.end, item.offset + item.length)) || text.slice(span.start, span.end), 90);
    var where = span.kind === 'bibliography' ? 'Pada referensi “' + quoted + '” (daftar pustaka ' + CM.nameOf(span.manager) + ')'
      : 'Pada sitasi “' + quoted + '” (dikelola ' + CM.nameOf(span.manager) + ')';
    var how = ' Perbaiki lewat ' + CM.nameOf(span.manager) + ': ' + CM.fixFor(span.manager, span.kind === 'bibliography' ? 'bibliography' : 'citation') + '.';
    function tryRange(p, s0, e0) {
      if (e0 <= s0 || CM.managedAt(managed, p, s0, e0)) return null;
      // Kata di sebelah field bukan bagian field, jadi aman disorot dengan warna temuannya.
      return Object.assign({}, item, { paraIndex: p, offset: s0, length: e0 - s0, managed: span.manager,
        comment: where + ': ' + sentence(item.comment) + how });
    }
    var before = /(\S+)\s*$/.exec(text.slice(0, span.start));
    var got = before ? tryRange(item.paraIndex, before.index, before.index + before[1].length) : null;
    if (!got) {
      var after = /^(\s*)(\S+)/.exec(text.slice(span.end));
      if (after) got = tryRange(item.paraIndex, span.end + after[1].length, span.end + after[1].length + after[2].length);
    }
    if (!got && st.headingIndex >= 0) {
      var h = st.paragraphs[st.headingIndex].text, lead = h.length - h.replace(/^\s+/, '').length;
      got = tryRange(st.headingIndex, lead, h.replace(/\s+$/, '').length);
    }
    return got || Object.assign({}, item, { unanchored: true, managed: span.manager });
  }

  // ---------- Scopus (sama dengan penjelasan di tab Metadata) ----------
  function isPeriodicalRef(ref) {
    return /^(?:journal-article|conference|book-chapter|magazine-article)$/.test(ref.sourceType || '') || (!!ref.journal && ref.sourceType !== 'website' && ref.sourceType !== 'book');
  }
  function scopusExplain(sc, ref) {
    var pct = Math.round((sc.confidence || 0) * 100);
    if (sc.status === 'SCOPUS') return sc.method === 'DOI_EXACT' ? 'Dokumen ditemukan lewat DOI yang cocok persis.' : 'Dokumen ditemukan lewat kecocokan metadata ' + pct + '%.';
    if (sc.status === 'PROBABLE_SCOPUS') return 'Metadata cukup mirip (' + pct + '%) dengan satu dokumen, belum cukup pasti.';
    if (sc.status === 'SCOPUS_SOURCE_ONLY') return 'Jurnal "' + ((sc.matchedSource && sc.matchedSource.title) || (ref.bibliographicModel && ref.bibliographicModel.journal) || ref.journal || '-') + '" terindeks Scopus; dokumen ini belum terverifikasi.' + (sc.discontinuedWarning ? ' Perhatian: jurnal ini berstatus "Discontinued by Scopus".' : '');
    if (sc.method === 'JOURNAL_FOUND_YEAR_NOT_COVERED') return 'Jurnal ada di Scopus, tetapi tahun ' + (ref.year || '-') + ' di luar cakupannya (' + ((sc.matchedSource && sc.matchedSource.coverage) || '-') + ').';
    return 'Belum ditemukan bukti di data Scopus yang dimuat (bukan berarti pasti non-Scopus).';
  }
  // Masalah = tidak ditemukan / tahun di luar cakupan / jurnal discontinued. Sisanya informasi.
  function scopusProblem(sc) { return sc.status === 'UNKNOWN' || !!sc.discontinuedWarning; }
  var SCOPUS_LABEL = { SCOPUS: 'Terindeks Scopus', PROBABLE_SCOPUS: 'Kemungkinan terindeks Scopus', SCOPUS_SOURCE_ONLY: 'Jurnal terindeks Scopus', UNKNOWN: 'Tidak ditemukan di Scopus' };

  // Warna sorotan komentar ringkasan = kategori temuan terberat yang ada di laporan:
  // merah bila ada yang perlu diperbaiki, biru bila hanya saran, kuning bila hanya metadata.
  function summaryColor(d) {
    d = d || {};
    var v = d.validator || {}, t = Object.assign({}, d.template || {}), doi = d.doi || {}, y = d.years || {}, sc = d.scopus || {};
    if (d.term) { t.bad = (t.bad || 0) + d.term.bad; t.warn = (t.warn || 0) + d.term.warn; }
    if (v.errors || t.bad) return COLOR.error;
    if (v.suggestions || t.warn) return COLOR.suggestion;
    if (d.validator || d.template || d.term) {
      if (doi.bad || doi.nodoi || y.out || sc.unknown || sc.discontinued) return COLOR.metadata;
      return COLOR.suggestion;
    }
    return COLOR.metadata; // laporan metadata
  }

  // Ringkasan seluruh pemeriksaan → satu komentar di judul naskah.
  // data: { style, validator: {errors, suggestions}, doi: {ok, bad, nodoi, pending}, years: {percent, from, to, out},
  //         scopus: {indexed, sourceOnly, unknown}, template: {bad, warn, ok, totalWords} }
  function summaryComment(d) {
    var L = [];
    if (d.style) L.push('Gaya sitasi: ' + d.style + '.');
    if (d.validator) L.push('Validator: ' + d.validator.errors + ' perlu diperbaiki, ' + d.validator.suggestions + ' saran.');
    if (d.doi) L.push('DOI (CrossRef): ' + d.doi.ok + ' valid, ' + d.doi.bad + ' bermasalah, ' + d.doi.nodoi + ' tanpa DOI' + (d.doi.pending ? ', ' + d.doi.pending + ' belum selesai dicek' : '') + (d.doi.error ? ', ' + d.doi.error + ' tidak bisa dicek (CrossRef tidak terjangkau)' : '') + '.');
    if (d.years && d.years.percent != null) L.push('Rentang tahun: ' + d.years.percent + '% referensi dari ' + d.years.from + '–' + d.years.to + (d.years.out ? ' (' + d.years.out + ' di luar rentang)' : '') + '.');
    if (d.scopus) L.push('Scopus: ' + d.scopus.indexed + ' terindeks/jurnal terindeks, ' + d.scopus.unknown + ' tidak ditemukan' + (d.scopus.discontinued ? ', ' + d.scopus.discontinued + ' jurnal discontinued' : '') + '.');
    if (d.template) L.push('Template: ' + d.template.bad + ' tidak sesuai, ' + d.template.warn + ' perlu dicek, ' + d.template.ok + ' sesuai' + (d.template.totalWords != null ? '; total ' + d.template.totalWords + ' kata' : '') + '.');
    if (d.term) L.push('Konsistensi istilah: ' + d.term.bad + ' perlu diperbaiki, ' + d.term.warn + ' perlu dicek.');
    if (d.legend !== false) L.push(LEGEND);
    return 'Ringkasan laporan — ' + L.join(' ');
  }

  // opts: { includeErrors, includeSuggestions, includeYears, includeDoi, includeNoDoi, includeScopus, scopusAll,
  //         years, references, online, scopus, summary, styleId, managed, rawTexts, prefix }
  function build(st, suggestions, opts) {
    var o = opts || {};
    var raw = [], counts = { error: 0, suggestion: 0, info: 0, year: 0, doi: 0, nodoi: 0, scopus: 0, template: 0, term: 0, summary: 0, documentLevel: 0 };
    function push(it, what) { if (it.length > 0 && st.paragraphs[it.paraIndex]) { raw.push(it); counts[what]++; } }
    function refSpan(r) { return { paraIndex: r.paraIndex, offset: r.leading, length: Math.max(1, r.text.replace(/\s+$/, '').length - r.leading) }; }

    // 0) Ringkasan semua fitur di judul naskah (paragraf pertama yang berisi teks) — didorong
    // paling awal supaya menjadi butir pertama bila digabung dengan komentar lain di judul.
    if (o.summary) {
      var first = st.paragraphs.findIndex(function (p) { return p.text.trim().length > 0; });
      if (first >= 0) {
        var ft = st.paragraphs[first].text, lead0 = ft.length - ft.replace(/^\s+/, '').length;
        push({ paraIndex: first, offset: lead0, length: ft.trim().length, color: summaryColor(o.summary), comment: summaryComment(o.summary) }, 'summary');
      }
    }

    // 1) Saran Validator (yang sudah diabaikan tidak ikut — disaring pemanggil).
    (suggestions || []).forEach(function (s) {
      var isErr = s.severity === 'error', isInfo = s.severity === 'info';
      // s.countAs (mis. 'template'): sumber lain yang sudah disaring pemanggil — tidak ikut
      // pilihan Perlu diperbaiki/Saran Validator dan dihitung terpisah.
      if (!s.countAs && (isInfo ? o.includeInfo === false : isErr ? o.includeErrors === false : o.includeSuggestions === false)) return;
      var loc = null;
      if (s.paraIndex != null && s.length > 0) loc = { paraIndex: s.paraIndex, offset: s.offset || 0, length: s.length };
      else if (st.headingIndex >= 0) { var h = st.paragraphs[st.headingIndex].text; loc = { paraIndex: st.headingIndex, offset: h.length - h.replace(/^\s+/, '').length, length: h.trim().length }; counts.documentLevel++; }
      if (!loc) return;
      // s.comment / s.color: teks & warna yang sudah disiapkan pemanggil (mis. Cek Template).
      // Temuan online (DOI dari CrossRef) termasuk metadata → kuning.
      var color = s.color !== undefined ? s.color : isInfo ? null : s.category === 'online' ? COLOR.metadata : isErr ? COLOR.error : COLOR.suggestion;
      push(Object.assign(loc, { color: color, comment: s.comment || suggestionComment(s), source: s.id }), s.countAs || (isInfo ? 'info' : isErr ? 'error' : 'suggestion'));
    });

    // 2) Tahun referensi di luar rentang / tidak terdeteksi.
    if (o.includeYears !== false && o.references && o.references.length) {
      var yr = AC.yearRange(o.references, o.years || 10);
      var range = (o.years || 10) + ' tahun terakhir (' + yr.from + '–' + yr.to + ')';
      yr.outRange.forEach(function (x) {
        var idx = AC.refIndexOfRaw(st, x.ref.raw); if (idx < 0) return;
        var r = st.refs[idx], m = new RegExp('(^|[^\\d])(' + x.y + '[a-z]?)(?!\\d)').exec(r.text);
        var loc = m ? { paraIndex: r.paraIndex, offset: m.index + m[1].length, length: m[2].length } : refSpan(r);
        push(Object.assign(loc, { color: COLOR.year, comment: 'Tahun referensi — tahun ' + x.y + ' berada di luar rentang ' + range + '. Pertimbangkan sumber yang lebih baru, atau abaikan bila memang sengaja mengacu sumber klasik/teori dasar.' }), 'year');
      });
      yr.unknown.forEach(function (ref) {
        var idx = AC.refIndexOfRaw(st, ref.raw); if (idx < 0) return;
        var r = st.refs[idx], first = r.text.slice(r.leading).split(/\s+/).slice(0, 4).join(' ');
        push({ paraIndex: r.paraIndex, offset: r.leading, length: first.length, color: COLOR.year,
          comment: 'Tahun referensi — tahun untuk referensi ini tidak terdeteksi otomatis. Periksa manual apakah masih dalam rentang ' + range + '.' }, 'year');
      });
    }

    // 3) DOI bermasalah menurut CrossRef (hanya bila cek online sudah berjalan).
    if (o.includeDoi !== false && o.online && o.references) {
      o.references.forEach(function (ref) {
        var rec = o.online[(o.styleId || '') + '|' + ref.raw];
        if (!rec || (rec.status !== 'mismatch' && rec.status !== 'notfound')) return;
        var idx = AC.refIndexOfRaw(st, ref.raw); if (idx < 0) return;
        var r = st.refs[idx], at = rec.doi ? r.text.toLowerCase().indexOf(String(rec.doi).toLowerCase()) : -1;
        var loc = at >= 0 ? { paraIndex: r.paraIndex, offset: at, length: rec.doi.length } : refSpan(r);
        var diff = (rec.fields || []).filter(function (f) { return !f.ok && !f.info; }).map(function (f) { return f.field + ': naskah “' + clip(f.naskah || '–', 80) + '”, CrossRef “' + clip(f.crossref || '–', 80) + '”'; });
        push(Object.assign(loc, { color: rec.status === 'mismatch' ? COLOR.doiBad : COLOR.doi,
          comment: 'DOI bermasalah — ' + sentence(rec.note || (rec.status === 'mismatch' ? 'DOI tampaknya milik karya lain.' : 'DOI tidak ditemukan di CrossRef.')) + (diff.length ? ' ' + diff.join('; ') + '.' : '') }), 'doi');
      });
    }

    // 3b) Tanpa DOI — hanya jenis sumber yang biasanya punya DOI (artikel/prosiding).
    if (o.includeNoDoi && o.online && o.references) {
      o.references.forEach(function (ref) {
        var rec = o.online[(o.styleId || '') + '|' + ref.raw];
        if (!rec || rec.status !== 'nodoi') return;
        var idx = AC.refIndexOfRaw(st, ref.raw); if (idx < 0) return;
        push(Object.assign(refSpan(st.refs[idx]), { color: COLOR.nodoi,
          comment: 'Tanpa DOI — artikel/prosiding biasanya punya DOI. Cari lewat tombol "Cari DOI" di tab Metadata, atau di situs penerbit.' }), 'nodoi');
      });
    }

    // 3c) Status Scopus (hanya bila data Scopus sudah dimuat). Masalah disorot; yang terindeks
    // dijadikan komentar tanpa sorotan bila scopusAll.
    if (o.includeScopus && o.scopus && o.scopus.length) {
      o.scopus.forEach(function (sc) {
        var ref = sc.ref; if (!ref) return;
        if (!isPeriodicalRef(ref) && sc.status === 'UNKNOWN') return;
        var problem = scopusProblem(sc);
        if (!problem && !o.scopusAll) return;
        var idx = AC.refIndexOfRaw(st, ref.raw); if (idx < 0) return;
        // Nama jurnal dari model bibliografi (sama dengan yang dicocokkan ke Scopus).
        var jName = (ref.bibliographicModel && ref.bibliographicModel.journal) || ref.journal || '';
        var r = st.refs[idx], journal = jName ? r.text.indexOf(jName) : -1;
        var loc = journal >= 0 ? { paraIndex: r.paraIndex, offset: journal, length: jName.length } : refSpan(r);
        push(Object.assign(loc, { color: COLOR.scopus,
          comment: 'Scopus — ' + (sc.discontinuedWarning ? 'Jurnal discontinued' : SCOPUS_LABEL[sc.status] || sc.status) + '. ' + scopusExplain(sc, ref) }), 'scopus');
      });
    }

    // 4) Lindungi bagian Zotero/Mendeley: pindahkan jangkar komentar keluar dari field-nya.
    var unanchored = 0, managedCount = 0;
    var placed = raw.map(function (it) { return o.managed && o.managed.found ? anchorOutside(st, o.managed, it) : it; }).filter(function (it) {
      if (it.managed) managedCount++;
      if (it.unanchored) { unanchored++; return false; }
      return true;
    });

    // 5) Gabungkan komentar yang rentangnya bertumpang tindih (satu kotak komentar per lokasi).
    placed.sort(function (a, b) { return (a.paraIndex - b.paraIndex) || (a.offset - b.offset) || (b.length - a.length); });
    var items = [];
    placed.forEach(function (it) {
      var last = items[items.length - 1];
      if (last && last.paraIndex === it.paraIndex && it.offset < last.offset + last.length) {
        if (last.comments.indexOf(it.comment) < 0) last.comments.push(it.comment);
        last.length = Math.max(last.offset + last.length, it.offset + it.length) - last.offset;
        if ((RANK[it.color] || 0) > (RANK[last.color] || 0)) last.color = it.color;
        if (it.managed) last.managed = last.managed || it.managed;
        return;
      }
      items.push({ paraIndex: it.paraIndex, offset: it.offset, length: it.length, color: it.color || null, comments: [it.comment], managed: it.managed || null });
    });

    // 6) Teks pencarian untuk Word (memakai teks paragraf asli dari Word bila ada).
    items.forEach(function (it) {
      var text = (o.rawTexts && o.rawTexts[it.paraIndex] != null ? o.rawTexts[it.paraIndex] : st.paragraphs[it.paraIndex].text);
      it.text = text.substr(it.offset, it.length);
      it.comment = (o.prefix || PREFIX) + ' ' + (it.comments.length > 1 ? it.comments.map(function (c, i) { return (i + 1) + ') ' + c; }).join('\n') : it.comments[0]);
      var trimmedStart = text.length - text.replace(/^\s+/, '').length, trimmedEnd = text.replace(/\s+$/, '').length;
      it.whole = it.offset <= trimmedStart && it.offset + it.length >= trimmedEnd;
      if (it.whole) return;
      if (it.length <= 240) { it.occurrence = AC.occurrenceIndex(text, it.text, it.offset, false); return; }
      it.head = text.substr(it.offset, 100); it.headOccurrence = AC.occurrenceIndex(text, it.head, it.offset, false);
      var ts = it.offset + it.length - 100;
      it.tail = text.substr(ts, 100); it.tailOccurrence = AC.occurrenceIndex(text, it.tail, ts, false);
    });

    return { items: items, counts: counts, total: raw.length, unanchored: unanchored, managed: managedCount };
  }

  function lead(text) { return String(text || '').replace(/^\s+/, ''); }
  function isReportComment(text) { var t = lead(text); return t.indexOf(PREFIX) === 0 || t.indexOf(TEMPLATE_PREFIX) === 0 || t.indexOf(METADATA_PREFIX) === 0 || t.indexOf(TERM_PREFIX) === 0; }
  function isTermComment(text) { return lead(text).indexOf(TERM_PREFIX) === 0; }
  function isTemplateComment(text) { return lead(text).indexOf(TEMPLATE_PREFIX) === 0; }
  function isMetadataComment(text) { return lead(text).indexOf(METADATA_PREFIX) === 0; }
  // which: 'all' | 'validator' | 'template' | 'metadata' | 'term'. Laporan keseluruhan memakai PREFIX
  // biasa dan mengganti 'all', jadi tidak pernah menumpuk dengan laporan per tab.
  function commentMatches(text, which) {
    if (which === 'template') return isTemplateComment(text);
    if (which === 'metadata') return isMetadataComment(text);
    if (which === 'term') return isTermComment(text);
    if (which === 'validator') return isReportComment(text) && !isTemplateComment(text) && !isMetadataComment(text) && !isTermComment(text);
    return isReportComment(text);
  }

  var api = { build: build, PREFIX: PREFIX, TEMPLATE_PREFIX: TEMPLATE_PREFIX, METADATA_PREFIX: METADATA_PREFIX, TERM_PREFIX: TERM_PREFIX, COLOR: COLOR, LEGEND: LEGEND,
    isReportComment: isReportComment, isTemplateComment: isTemplateComment, isMetadataComment: isMetadataComment, isTermComment: isTermComment, commentMatches: commentMatches,
    suggestionComment: suggestionComment, summaryComment: summaryComment, summaryColor: summaryColor, scopusExplain: scopusExplain, scopusProblem: scopusProblem, isPeriodicalRef: isPeriodicalRef };
  if (isNode) module.exports = api;
  root.ReportCore = api;
})(typeof window !== 'undefined' ? window : this);
