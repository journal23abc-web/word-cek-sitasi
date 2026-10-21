/* ============================================================
   DOCX READER — model dokumen untuk Cek Template, dibaca langsung dari OOXML.

   Kenapa dari OOXML (bukan paragraph.text Office.js):
     - abstrak/kata kunci/"How to cite" sering berada di SEL TABEL halaman pertama, dan
       "How to cite" kadang di FOOTER halaman pertama — keduanya perlu struktur sel/kisah;
     - hitungan kata ala Word mencakup kotak teks, catatan kaki & catatan akhir;
     - nomor halaman awal ("Start at: 74") hanya ada di w:pgNumType bagian pertama.

   Masukan `parts`: { document, footnotes?, endnotes?, rels?, app?, hf?: { 'word/footer3.xml': xml, ... } }
   atau, dari Office.js tanpa file .docx, { document, headers: {first,default,even}, footers: {...} }.

   Paragraf badan dihitung dengan urutan yang sama seperti body.paragraphs di Word: paragraf
   tabel ikut, isi kotak teks dan salinan cadangan (mc:Fallback) tidak.

   Murni (tanpa DOM), bisa diuji di Node.
   ============================================================ */
(function (root) {
  'use strict';

  var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function decode(s) {
    return s.indexOf('&') < 0 ? s : s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (m, e) {
      if (e[0] === '#') { var n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return isNaN(n) ? m : String.fromCodePoint(n); }
      return ENT[e] != null ? ENT[e] : m;
    });
  }
  function attr(attrs, name) {
    var re = new RegExp('(?:^|\\s)' + name.replace(/[.:]/g, '\\$&') + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')');
    var m = re.exec(attrs);
    return m ? decode(m[1] != null ? m[1] : m[2]) : null;
  }

  var TOKEN = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)([A-Za-z_][\w.-]*(?::[\w.-]+)?)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>|([^<]+)/g;

  // Membaca satu "kisah" (story): badan dokumen, catatan kaki, header/footer, isi kotak teks.
  // Hasil: { paragraphs: [{ text, countText, cell, tableIndex, style, outline, bold, size, notes }],
  //          textboxes: [paragraf...], sections: [{ pgStart, titlePg, refs }], breaks }
  // text      = sama seperti paragraph.text di Word (tab \t, ganti baris \u000b, hasil field)
  // countText = teks yang dihitung Word sebagai kata (tanpa teks tersembunyi)
  function readStory(xml, opts) {
    var o = opts || {};
    var src = String(xml || '');
    var paras = [], textboxes = [], sections = [];
    var stack = []; // konteks paragraf (untuk kotak teks bersarang)
    var cur = null;
    var fields = [];
    var tables = [], tableCount = 0;
    var inT = false, inDel = 0, inInstr = false, skipFallback = 0, txbxDepth = 0;
    var run = null; // { hidden, bold, size }
    var inRPr = 0, inPPr = 0, inSectPr = 0, sect = null;
    var lastRendered = 0, pageBreaks = 0, noteKind = o.noteKind || null, curNote = null, notes = [];
    var m;

    function visibleField() { for (var i = 0; i < fields.length; i++) if (fields[i] !== 'result') return false; return true; }
    function newPara() {
      var cell = null, top = tables[tables.length - 1];
      if (top && top.cell) cell = { table: top.index, row: top.row, col: top.col, gridCol: top.gridCol, key: top.index + ':' + top.row + ':' + top.col, nested: tables.length > 1 };
      return { text: '', countText: '', cell: cell, tableIndex: top ? top.index : -1, style: null, outline: null, bold: null, size: null, runs: 0, boldChars: 0, chars: 0, maxSize: 0, sectionBreak: false };
    }
    function add(t, countable) {
      if (!cur || !t) return;
      cur.text += t;
      if (countable !== false && !(run && run.hidden)) cur.countText += t;
      if (/\S/.test(t)) {
        var n = t.replace(/\s/g, '').length;
        cur.chars += n;
        if (run && run.bold) cur.boldChars += n;
        if (run && run.size) cur.maxSize = Math.max(cur.maxSize, run.size);
      }
    }

    while ((m = TOKEN.exec(src))) {
      if (m[5] != null) {
        if (skipFallback) continue;
        if (inT && !inDel && visibleField()) add(decode(m[5]));
        continue;
      }
      if (!m[2]) continue;
      var close = m[1] === '/', name = m[2], attrs = m[3] || '', selfClose = m[4] === '/';

      if (name === 'mc:Fallback') { if (!selfClose) skipFallback += close ? -1 : 1; continue; }
      if (skipFallback) continue;

      switch (name) {
        case 'w:txbxContent':
          if (selfClose) break;
          if (!close) { stack.push({ cur: cur, tables: tables, fields: fields }); cur = null; tables = []; fields = []; txbxDepth++; }
          else { var s = stack.pop(); if (s) { cur = s.cur; tables = s.tables; fields = s.fields; } txbxDepth--; }
          break;
        case 'w:footnote': case 'w:endnote':
          if (close) { if (curNote) notes.push(curNote); curNote = null; }
          else if (!selfClose) { var nt = attr(attrs, 'w:type'); curNote = nt && nt !== 'normal' ? null : { id: attr(attrs, 'w:id'), paragraphs: [] }; }
          break;
        case 'w:p':
          if (close) {
            if (cur) {
              if (txbxDepth) textboxes.push(cur);
              else if (curNote) curNote.paragraphs.push(cur);
              else if (!noteKind) paras.push(cur);
            }
            cur = null;
          } else if (selfClose) {
            var e = newPara();
            if (txbxDepth) textboxes.push(e); else if (curNote) curNote.paragraphs.push(e); else if (!noteKind) paras.push(e);
          } else cur = newPara();
          break;
        case 'w:pPr': if (!selfClose) inPPr += close ? -1 : 1; break;
        case 'w:pStyle': if (inPPr && cur && !inRPr && !close) cur.style = attr(attrs, 'w:val'); break;
        case 'w:outlineLvl': if (inPPr && cur && !close) cur.outline = parseInt(attr(attrs, 'w:val'), 10); break;
        case 'w:r':
          if (!selfClose) run = close ? null : { hidden: false, bold: false, size: 0 };
          break;
        case 'w:rPr': if (!selfClose) inRPr += close ? -1 : 1; break;
        case 'w:vanish': if (inRPr && run && !close) run.hidden = attr(attrs, 'w:val') !== '0' && attr(attrs, 'w:val') !== 'false'; break;
        case 'w:b': if (inRPr && run && !close) run.bold = attr(attrs, 'w:val') !== '0' && attr(attrs, 'w:val') !== 'false'; break;
        case 'w:sz': if (inRPr && run && !close) run.size = parseInt(attr(attrs, 'w:val'), 10) / 2 || 0; break;
        case 'w:t': inT = !close && !selfClose; break;
        case 'm:t': inT = !close && !selfClose; break;
        case 'w:delText': inT = false; break;
        case 'w:del': case 'w:moveFrom': if (!selfClose) inDel += close ? -1 : 1; if (inDel < 0) inDel = 0; break;
        case 'w:instrText': case 'w:delInstrText': inInstr = !close && !selfClose; break;
        case 'w:tab':
          if (run && !close && !inDel && !inPPr && visibleField()) add('\t');
          break;
        case 'w:br':
          if (run && !close && !inDel && visibleField()) { add('\u000b'); if (attr(attrs, 'w:type') === 'page') pageBreaks++; }
          break;
        case 'w:cr': if (run && !close && !inDel && visibleField()) add('\u000b'); break;
        case 'w:noBreakHyphen': if (run && !close && !inDel && visibleField()) add('-'); break;
        case 'w:sym': if (run && !close && !inDel && visibleField()) add('\u25A1'); break; // simbol = satu karakter
        case 'w:lastRenderedPageBreak': if (!close && !txbxDepth && !curNote) lastRendered++; break;
        case 'w:fldChar':
          var ft = attr(attrs, 'w:fldCharType');
          if (ft === 'begin') fields.push('code');
          else if (ft === 'separate' && fields.length) fields[fields.length - 1] = 'result';
          else if (ft === 'end') fields.pop();
          break;
        case 'w:tbl':
          if (selfClose) break;
          if (!close) tables.push({ index: tableCount++, row: -1, col: -1, gridCol: 0, cell: false });
          else tables.pop();
          break;
        case 'w:tr':
          if (!close && !selfClose && tables.length) { var tb = tables[tables.length - 1]; tb.row++; tb.col = -1; tb.gridCol = 0; tb.nextGrid = 0; }
          break;
        case 'w:tc':
          if (!tables.length || selfClose) break;
          var tt = tables[tables.length - 1];
          if (!close) { tt.col++; tt.gridCol = tt.nextGrid || 0; tt.span = 1; tt.cell = true; }
          else { tt.nextGrid = (tt.gridCol || 0) + (tt.span || 1); tt.cell = false; }
          break;
        case 'w:gridSpan':
          if (tables.length && !close) tables[tables.length - 1].span = parseInt(attr(attrs, 'w:val'), 10) || 1;
          break;
        case 'w:sectPr':
          if (selfClose) { sections.push({ pgStart: null, titlePg: false, refs: {}, inPara: !!inPPr }); if (cur) cur.sectionBreak = true; break; }
          if (!close) { inSectPr++; sect = { pgStart: null, titlePg: false, refs: {}, inPara: !!inPPr }; }
          else { inSectPr--; if (sect) { sections.push(sect); if (cur && sect.inPara) cur.sectionBreak = true; } sect = null; }
          break;
        case 'w:pgNumType':
          if (inSectPr && sect && !close) { var st = attr(attrs, 'w:start'); if (st != null && st !== '') sect.pgStart = parseInt(st, 10); sect.pgFmt = attr(attrs, 'w:fmt'); }
          break;
        case 'w:titlePg':
          if (inSectPr && sect && !close) sect.titlePg = attr(attrs, 'w:val') !== '0' && attr(attrs, 'w:val') !== 'false';
          break;
        case 'w:headerReference': case 'w:footerReference':
          if (inSectPr && sect && !close) sect.refs[(name === 'w:headerReference' ? 'header:' : 'footer:') + (attr(attrs, 'w:type') || 'default')] = attr(attrs, 'r:id');
          break;
      }
    }
    if (cur) paras.push(cur);
    paras.concat(textboxes).forEach(finishPara);
    notes.forEach(function (n) { n.paragraphs.forEach(finishPara); });
    return { paragraphs: paras, textboxes: textboxes, notes: notes, sections: sections, lastRenderedPageBreaks: lastRendered, pageBreaks: pageBreaks };
  }
  function finishPara(p) {
    p.bold = p.chars > 0 && p.boldChars >= p.chars * 0.8;
    p.size = p.maxSize || null;
    delete p.boldChars; delete p.maxSize; delete p.runs;
  }

  // Flat OPC (body.getOoxml()) → isi bagian tertentu.
  function partFromFlatOpc(flat, partName) {
    var s = String(flat || '');
    var re = new RegExp('<pkg:part[^>]*pkg:name="' + partName.replace(/[.\/]/g, '\\$&') + '"[^>]*>[\\s\\S]*?<pkg:xmlData>([\\s\\S]*?)</pkg:xmlData>');
    var m = re.exec(s);
    return m ? m[1] : null;
  }
  function relsMap(relsXml) {
    var map = {}, re = /<Relationship\b([^>]*)\/?>/g, m;
    while ((m = re.exec(String(relsXml || '')))) map[attr(m[1], 'Id')] = (attr(m[1], 'Target') || '').replace(/^\/?word\//, '');
    return map;
  }
  function appStats(appXml) {
    var s = String(appXml || '');
    function num(tag) { var m = new RegExp('<' + tag + '>(\\d+)</' + tag + '>').exec(s); return m ? parseInt(m[1], 10) : null; }
    return { pages: num('Pages'), words: num('Words'), application: (/<Application>([^<]*)</.exec(s) || [])[1] || null };
  }

  // Model lengkap dokumen untuk Cek Template.
  function read(parts) {
    var p = parts || {};
    var doc = p.document && /<pkg:package/.test(p.document) ? partFromFlatOpc(p.document, '/word/document.xml') : p.document;
    var body = readStory(doc || '');
    var fnXml = p.footnotes || (p.document && /<pkg:package/.test(p.document) ? partFromFlatOpc(p.document, '/word/footnotes.xml') : null);
    var enXml = p.endnotes || (p.document && /<pkg:package/.test(p.document) ? partFromFlatOpc(p.document, '/word/endnotes.xml') : null);
    var footnotes = fnXml ? readStory(fnXml, { noteKind: 'footnote' }).notes : null;
    var endnotes = enXml ? readStory(enXml, { noteKind: 'endnote' }).notes : null;

    // Bagian pertama: sectPr pertama di dalam paragraf, atau sectPr badan bila hanya satu bagian.
    var sections = body.sections;
    var first = sections[0] || { pgStart: null, titlePg: false, refs: {} };
    var restarts = sections.slice(1).filter(function (s) { return s.pgStart != null; }).length;

    // Header/footer bagian pertama (yang tampil di halaman 1 dan halaman-halaman berikutnya).
    var hf = {};
    function hfStory(xml) { if (!xml) return null; var x = /<pkg:package/.test(xml) ? (partFromFlatOpc(xml, '/word/header1.xml') || partFromFlatOpc(xml, '/word/footer1.xml') || partFromFlatOpc(xml, '/word/document.xml') || xml) : xml; return readStory(x).paragraphs; }
    if (p.headers || p.footers) {
      ['first', 'default', 'even'].forEach(function (t) {
        if (p.headers && p.headers[t]) hf['header:' + t] = hfStory(p.headers[t]);
        if (p.footers && p.footers[t]) hf['footer:' + t] = hfStory(p.footers[t]);
      });
    } else if (p.hf) {
      var rels = relsMap(p.rels);
      Object.keys(first.refs).forEach(function (k) {
        var target = rels[first.refs[k]];
        if (target && p.hf['word/' + target] != null) hf[k] = readStory(p.hf['word/' + target]).paragraphs;
      });
    }

    var stats = appStats(p.app);
    return {
      body: body.paragraphs,
      textboxes: body.textboxes,
      footnotes: footnotes, endnotes: endnotes,
      headerFooter: hf,
      firstPageFooter: first.titlePg ? 'footer:first' : 'footer:default',
      firstPageHeader: first.titlePg ? 'header:first' : 'header:default',
      section: { pgStart: first.pgStart, pgFmt: first.pgFmt || null, titlePg: first.titlePg, count: Math.max(1, sections.length), restarts: restarts },
      app: stats,
      layoutHints: { lastRenderedPageBreaks: body.lastRenderedPageBreaks, pageBreaks: body.pageBreaks }
    };
  }

  // Dari file .docx (JSZip) → parts.
  function partsFromZip(zip) {
    function get(name) { var f = zip.file(name); return f ? f.async('string') : Promise.resolve(null); }
    var hfNames = Object.keys(zip.files).filter(function (n) { return /^word\/(header|footer)\d*\.xml$/.test(n); });
    return Promise.all([get('word/document.xml'), get('word/footnotes.xml'), get('word/endnotes.xml'), get('word/_rels/document.xml.rels'), get('docProps/app.xml')].concat(hfNames.map(get))).then(function (x) {
      var hf = {}; hfNames.forEach(function (n, i) { hf[n] = x[5 + i]; });
      return { document: x[0], footnotes: x[1], endnotes: x[2], rels: x[3], app: x[4], hf: hf };
    });
  }

  var api = { read: read, readStory: readStory, partsFromZip: partsFromZip, partFromFlatOpc: partFromFlatOpc, appStats: appStats };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.DocxReader = api;
})(typeof window !== 'undefined' ? window : this);
