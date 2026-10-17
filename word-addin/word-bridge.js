/* ============================================================
   WORD BRIDGE — semua interaksi Office.js (baca/tulis dokumen Word).
   Lokasi selalu ditentukan sebagai { paraIndex, teks, kemunculan ke-n } lalu dicari
   ulang di Word, jadi tidak bergantung pada offset internal Word.
   ============================================================ */
(function (root) {
  'use strict';

  function supports(v) {
    try { return !!(root.Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported('WordApi', v)); }
    catch (e) { return false; }
  }
  var capsCache = null;
  function caps() {
    if (!capsCache) capsCache = { api13: supports('1.3'), api14: supports('1.4'), api16: supports('1.6'), annotations: supports('1.7'), popup: supports('1.8') };
    return capsCache;
  }

  function escapeSearch(t) { return String(t).replace(/\^/g, '^^'); }

  // ---------- BACA ----------
  function readParagraphs() {
    return Word.run(function (ctx) {
      var ps = ctx.document.body.paragraphs;
      ps.load(caps().api13 ? 'items/text,items/isListItem,items/tableNestingLevel' : 'items/text');
      return ctx.sync().then(function () {
        var lists = [];
        var out = ps.items.map(function (p, i) {
          var o = { text: p.text, isListItem: !!p.isListItem, inTable: p.tableNestingLevel > 0 };
          if (p.isListItem) { var li = p.listItemOrNullObject; li.load('listString'); lists.push([i, li]); }
          return o;
        });
        if (!lists.length) return out;
        return ctx.sync().then(function () {
          lists.forEach(function (x) { if (!x[1].isNullObject) out[x[0]].listString = x[1].listString; });
          return out;
        }).catch(function () { return out; });
      });
    });
  }

  function loadParas(ctx) {
    var ps = ctx.document.body.paragraphs;
    ps.load('items/text');
    return ctx.sync().then(function () { return ps; });
  }

  // Pastikan paragraf di posisi yang direncanakan masih berisi teks yang sama.
  function checkSnapshot(ps, snapshot) {
    if (!snapshot) return;
    var bad = Object.keys(snapshot).filter(function (i) { return !ps.items[i] || ps.items[i].text !== snapshot[i]; });
    if (bad.length) throw new Error('Dokumen sudah berubah sejak diperiksa. Klik "Periksa" lalu ulangi.');
  }

  // Antrekan pencarian untuk banyak item, satu kali sync. item: { paraIndex, text, occurrence, wholeWord }
  function queueSearches(ps, items) {
    return items.map(function (it) {
      var p = ps.items[it.paraIndex];
      if (!p || !it.text || it.text.length > 255) return null;
      var res = p.search(escapeSearch(it.text), { matchCase: true, matchWholeWord: !!it.wholeWord });
      res.load('items/text');
      return res;
    });
  }
  function pick(res, occurrence) {
    if (!res || !res.items || !res.items.length) return null;
    var i = occurrence >= 0 ? occurrence : 0;
    return res.items[Math.min(i, res.items.length - 1)];
  }

  function withTracking(ctx, track, fn) {
    if (!track || !caps().api14) return fn();
    var doc = ctx.document;
    doc.load('changeTrackingMode');
    var prev;
    return ctx.sync().then(function () {
      prev = doc.changeTrackingMode;
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      return ctx.sync();
    }).then(fn).then(function (v) {
      doc.changeTrackingMode = prev; return ctx.sync().then(function () { return v; });
    }, function (err) {
      doc.changeTrackingMode = prev; return ctx.sync().then(function () { throw err; });
    });
  }

  // Tulis ulang isi paragraf dengan runs [{text, italic}] — format paragraf tetap.
  function writeRuns(para, runs) {
    var list = runs.filter(function (r) { return r.text; });
    if (!list.length) { para.insertText('', 'Replace'); return; }
    var first = para.insertText(list[0].text, 'Replace');
    first.font.italic = !!list[0].italic;
    for (var i = 1; i < list.length; i++) {
      var r = para.insertText(list[i].text, 'End');
      r.font.italic = !!list[i].italic;
    }
  }

  function keepBookmarks(ctx, paras) {
    if (!caps().api14) return Promise.resolve(paras.map(function () { return []; }));
    var results = paras.map(function (p) { return p.getRange('Whole').getBookmarks(true, false); });
    return ctx.sync().then(function () { return results.map(function (r) { return (r.value || []).filter(function (n) { return /^_CS_REF_/.test(n); }); }); })
      .catch(function () { return paras.map(function () { return []; }); });
  }
  function restoreBookmarks(paras, names) {
    if (!caps().api14) return;
    paras.forEach(function (p, i) { (names[i] || []).forEach(function (n) { p.getRange('Content').insertBookmark(n); }); });
  }

  // ---------- NAVIGASI ----------
  function select(item) {
    return Word.run(function (ctx) {
      return loadParas(ctx).then(function (ps) {
        var p = ps.items[item.paraIndex];
        if (!p) return;
        if (item.kind === 'replace' && item.original) {
          var res = queueSearches(ps, [{ paraIndex: item.paraIndex, text: item.original, occurrence: item.occurrence }])[0];
          return ctx.sync().then(function () {
            var r = pick(res, item.occurrence);
            (r || p.getRange('Content')).select();
            return ctx.sync();
          });
        }
        p.getRange('Content').select();
        return ctx.sync();
      });
    });
  }

  // ---------- TERAPKAN SARAN ----------
  function applySuggestion(s, opts) {
    var o = opts || {};
    return Word.run(function (ctx) {
      return withTracking(ctx, o.track, function () {
        return loadParas(ctx).then(function (ps) {
          var p = ps.items[s.paraIndex];
          if (!p) throw new Error('Paragraf tidak ditemukan. Klik "Periksa" lalu ulangi.');
          if (s.kind === 'replace') {
            if (p.text.indexOf(s.original) < 0) throw new Error('Teks "' + s.original + '" tidak ditemukan lagi — dokumen mungkin sudah diedit. Klik "Periksa" lalu ulangi.');
            var res = queueSearches(ps, [{ paraIndex: s.paraIndex, text: s.original, occurrence: s.occurrence }])[0];
            return ctx.sync().then(function () {
              var r = pick(res, s.occurrence);
              if (!r) throw new Error('Teks asli tidak dapat ditemukan di Word.');
              r.insertText(s.replacement, 'Replace');
              return ctx.sync();
            });
          }
          if (s.kind === 'paragraph') {
            var expected = s.compareOriginal || s.original;
            if (AddinCoreNorm(p.text) !== AddinCoreNorm(expected) && AddinCoreNorm(p.text) !== AddinCoreNorm(s.original)) throw new Error('Isi referensi sudah berubah. Klik "Periksa" lalu ulangi.');
            return keepBookmarks(ctx, [p]).then(function (names) {
              writeRuns(p, s.replacementRuns || [{ text: s.replacement, italic: false }]);
              restoreBookmarks([p], names);
              return ctx.sync();
            });
          }
          if (s.kind === 'append') {
            if (s.stripTrailingPeriod) {
              var dots = p.search('.', { matchCase: true });
              dots.load('items');
              return ctx.sync().then(function () {
                var last = dots.items[dots.items.length - 1];
                if (last) last.insertText(s.appendText, 'Replace'); else p.insertText(s.appendText, 'End');
                return ctx.sync();
              });
            }
            p.getRange('Content').insertText(s.appendText, 'End');
            return ctx.sync();
          }
          if (s.kind === 'italic') {
            var searches = queueSearches(ps, s.segments.map(function (t) { return { paraIndex: s.paraIndex, text: t, occurrence: 0 }; }));
            return ctx.sync().then(function () {
              searches.forEach(function (res) { var r = pick(res, 0); if (r) r.font.italic = true; });
              return ctx.sync();
            });
          }
          throw new Error('Jenis saran tidak dikenal.');
        });
      });
    });
  }
  function AddinCoreNorm(t) { return String(t || '').replace(/^\s*(?:\[\d+\]|\d+\.)\s*/, '').replace(/[\s ]+/g, ' ').trim(); }

  // Terapkan banyak saran "replace" sekaligus (satu batch; rentang Word tetap valid antar edit).
  function applyMany(list, opts) {
    var o = opts || {};
    var replaces = list.filter(function (s) { return s.kind === 'replace'; });
    var others = list.filter(function (s) { return s.kind !== 'replace'; });
    var done = 0, failed = 0;
    return Word.run(function (ctx) {
      return withTracking(ctx, o.track, function () {
        return loadParas(ctx).then(function (ps) {
          // hindari dua saran yang menimpa teks yang sama
          var taken = {};
          var chosen = replaces.filter(function (s) {
            var k = s.paraIndex + ':' + s.offset;
            for (var x = s.offset; x < s.offset + s.length; x++) if (taken[s.paraIndex + ':' + x]) return false;
            for (var y = s.offset; y < s.offset + s.length; y++) taken[s.paraIndex + ':' + y] = true;
            return !!k;
          });
          failed += replaces.length - chosen.length;
          var searches = queueSearches(ps, chosen.map(function (s) { return { paraIndex: s.paraIndex, text: s.original, occurrence: s.occurrence }; }));
          return ctx.sync().then(function () {
            chosen.forEach(function (s, i) { var r = pick(searches[i], s.occurrence); if (r) { r.insertText(s.replacement, 'Replace'); done++; } else failed++; });
            return ctx.sync();
          });
        });
      });
    }).then(function () {
      var chain = Promise.resolve();
      others.forEach(function (s) { chain = chain.then(function () { return applySuggestion(s, o).then(function () { done++; }, function () { failed++; }); }); });
      return chain.then(function () { return { done: done, failed: failed }; });
    });
  }

  // ---------- CETAK MIRING ----------
  function checkItalics(plan) {
    if (!plan || !plan.length) return Promise.resolve([]);
    return Word.run(function (ctx) {
      return loadParas(ctx).then(function (ps) {
        var flat = [];
        plan.forEach(function (it) { it.segments.forEach(function (seg) { flat.push({ it: it, seg: seg }); }); });
        var res = queueSearches(ps, flat.map(function (f) { return { paraIndex: f.it.paraIndex, text: f.seg, occurrence: 0 }; }));
        return ctx.sync().then(function () {
          var ranges = res.map(function (r) { var x = pick(r, 0); if (x) x.load('font/italic'); return x; });
          return ctx.sync().then(function () {
            var bad = {};
            flat.forEach(function (f, i) {
              var r = ranges[i];
              if (r && r.font.italic !== true) { (bad[f.it.refIndex] = bad[f.it.refIndex] || { refIndex: f.it.refIndex, paraIndex: f.it.paraIndex, segments: [] }).segments.push(f.seg); }
            });
            return Object.keys(bad).map(function (k) { return bad[k]; });
          });
        });
      });
    });
  }

  // ---------- URUTKAN REFERENSI (pertahankan cetak miring/tebal per kata) ----------
  function captureRuns(ctx, paras) {
    if (!caps().api13) return Promise.resolve(paras.map(function (p) { return [{ text: p.text, italic: false }]; }));
    var trs = paras.map(function (p) { var t = p.getRange('Content').getTextRanges([' '], false); t.load('items/text,items/font/italic'); return t; });
    return ctx.sync().then(function () {
      return trs.map(function (t, i) {
        var runs = [];
        function push(text, italic) { if (!text) return; var l = runs[runs.length - 1]; if (l && l.italic === italic) l.text += text; else runs.push({ text: text, italic: italic }); }
        t.items.forEach(function (w) {
          if (w.font.italic === true) push(w.text, true);
          else if (w.font.italic === null) { var m = w.text.match(/^(.*?)([\s.,;:)\]]*)$/); push(m[1], true); push(m[2], false); }
          else push(w.text, false);
        });
        var joined = runs.map(function (r) { return r.text; }).join('');
        return joined.replace(/\s+/g, ' ').trim() === paras[i].text.replace(/\s+/g, ' ').trim() ? runs : [{ text: paras[i].text, italic: false }];
      });
    });
  }

  function reorderParagraphs(paraIndexes, order, snapshot, opts) {
    var o = opts || {};
    return Word.run(function (ctx) {
      return withTracking(ctx, o.track, function () {
        return loadParas(ctx).then(function (ps) {
          checkSnapshot(ps, snapshot);
          var paras = paraIndexes.map(function (i) { return ps.items[i]; });
          return Promise.all([captureRuns(ctx, paras), keepBookmarks(ctx, paras)]).then(function (x) {
            var runs = x[0], marks = x[1];
            paras.forEach(function (p, slot) { writeRuns(p, runs[order[slot]]); });
            restoreBookmarks(paras, order.map(function (src) { return marks[src]; }));
            return ctx.sync();
          });
        });
      });
    });
  }

  // ---------- KONVERSI ----------
  function applyConversion(plan, snapshot, opts) {
    var o = opts || {};
    return Word.run(function (ctx) {
      return withTracking(ctx, o.track, function () {
        return loadParas(ctx).then(function (ps) {
          checkSnapshot(ps, snapshot);
          var edits = plan.edits.filter(function (e) { return !e.crossesParagraph && e.occurrence >= 0; });
          var searches = queueSearches(ps, edits.map(function (e) { return { paraIndex: e.paraIndex, text: e.original, occurrence: e.occurrence }; }));
          var refParas = plan.references.paraIndexes.map(function (i) { return ps.items[i]; });
          if (caps().api13) refParas.forEach(function (p) { p.load('isListItem'); });
          return ctx.sync().then(function () {
            var applied = 0;
            edits.forEach(function (e, i) { var r = pick(searches[i], e.occurrence); if (r) { r.insertText(e.replacement, 'Replace'); applied++; } });
            var lines = plan.references.lines;
            var n = Math.min(lines.length, refParas.length);
            for (var k = 0; k < n; k++) {
              var p = refParas[k];
              if (plan.references.detachList && p.isListItem) { try { p.detachFromList(); } catch (e1) { /* versi lama */ } }
              writeRuns(p, lines[k].runs);
              if (lines[k].risky) { try { p.getRange('Content').font.highlightColor = 'Yellow'; } catch (e2) { /* abaikan */ } }
              if (o.hangingIndent && !(plan.references.keepListNumbering && p.isListItem)) {
                p.leftIndent = plan.target === 'ieee' ? 25 : 36; p.firstLineIndent = plan.target === 'ieee' ? -25 : -36;
              }
            }
            (plan.references.flagParaIndexes || []).forEach(function (i) { if (ps.items[i]) { try { ps.items[i].getRange('Content').font.highlightColor = 'Yellow'; } catch (e3) { /* abaikan */ } } });
            var last = refParas[refParas.length - 1];
            for (var m = n; m < lines.length && last; m++) { last = last.insertParagraph('', 'After'); writeRuns(last, lines[m].runs); }
            for (var d = n; d < refParas.length; d++) refParas[d].delete();
            return ctx.sync().then(function () { return { applied: applied, skipped: plan.edits.length - applied }; });
          });
        });
      });
    });
  }

  // ---------- TAUTAN ----------
  function applyLinks(plan, snapshot, opts) {
    var o = opts || {};
    return Word.run(function (ctx) {
      return loadParas(ctx).then(function (ps) {
        checkSnapshot(ps, snapshot);
        if (!caps().api14) throw new Error('Versi Word ini belum mendukung bookmark lewat add-in (butuh WordApi 1.4).');
        plan.bookmarks.forEach(function (b) { ps.items[b.paraIndex].getRange('Content').insertBookmark(b.name); });
        var items = (o.citations === false ? [] : plan.citationLinks).concat(o.external === false ? [] : plan.externalLinks);
        var searches = queueSearches(ps, items);
        return ctx.sync().then(function () {
          var ranges = items.map(function (it, i) { var r = pick(searches[i], it.occurrence); if (r && o.keepLook !== false) r.load('font/color,font/underline'); return r; });
          return ctx.sync().then(function () {
            var linked = 0, missed = 0;
            items.forEach(function (it, i) {
              var r = ranges[i];
              if (!r) { missed++; return; }
              var color = r.font.color, underline = r.font.underline;
              r.hyperlink = it.bookmark ? '#' + it.bookmark : it.href;
              if (o.keepLook !== false && it.bookmark) { if (color) r.font.color = color; if (underline) r.font.underline = underline; }
              linked++;
            });
            return ctx.sync().then(function () { return { linked: linked, missed: missed, bookmarks: plan.bookmarks.length }; });
          });
        });
      });
    });
  }

  function removeLinks() {
    return Word.run(function (ctx) {
      var whole = ctx.document.body.getRange('Whole');
      var links = whole.getHyperlinkRanges();
      links.load('items/hyperlink,items/text,items/font/color,items/font/underline');
      var marks = caps().api14 ? whole.getBookmarks(true, true) : null;
      return ctx.sync().then(function () {
        var ours = links.items.filter(function (r) { return /#_CS_REF_\d+$/.test(r.hyperlink || ''); });
        ours.forEach(function (r) { r.hyperlink = ''; });
        return ctx.sync().catch(function () {
          // Beberapa versi Word menolak string kosong: tulis ulang teksnya tanpa tautan.
          ours.forEach(function (r) { var c = r.font.color, u = r.font.underline; var n = r.insertText(r.text, 'Replace'); n.font.color = c; n.font.underline = u; });
          return ctx.sync();
        }).then(function () {
          var names = marks ? (marks.value || []).filter(function (n) { return /^_CS_REF_/.test(n); }) : [];
          names.forEach(function (n) { ctx.document.deleteBookmark(n); });
          return ctx.sync().then(function () { return { removed: ours.length, bookmarks: names.length }; });
        });
      });
    });
  }

  // ---------- SISIP REFERENSI BARU ----------
  function insertReference(point, runs, opts) {
    var o = opts || {};
    return Word.run(function (ctx) {
      return withTracking(ctx, o.track, function () {
        return loadParas(ctx).then(function (ps) {
          var anchor;
          if (point.paraIndex == null || point.paraIndex < 0) {
            anchor = ctx.document.body.insertParagraph(o.headingText || 'Daftar Pustaka', 'End');
            anchor = anchor.insertParagraph('', 'After');
          } else {
            anchor = ps.items[point.paraIndex].insertParagraph('', point.where || 'After');
          }
          anchor.getRange('Whole').font.highlightColor = null;
          runs.forEach(function (r) { var x = anchor.insertText(r.text, 'End'); x.font.italic = !!r.italic; });
          anchor.select();
          return ctx.sync();
        });
      });
    });
  }

  // ---------- ANOTASI (garis bawah ala Paperpal, Word M365) ----------
  var COLOR = { error: 'Red', suggestion: 'Blue', online: 'Lavender', info: 'Berry' };
  function annotate(items, useResources) {
    if (!caps().annotations) return Promise.resolve({ supported: false, ids: {} });
    return Word.run(function (ctx) {
      return loadParas(ctx).then(function (ps) {
        var results = [];
        items.forEach(function (it) {
          var p = ps.items[it.paraIndex];
          if (!p || it.length <= 0) return;
          var crit = { colorScheme: COLOR[it.color] || 'Blue', start: it.offset, length: it.length };
          if (useResources && caps().popup && it.popup && it.popup.length) {
            crit.popupOptions = { brandingTextResourceId: 'CS.Brand', titleResourceId: 'CS.Title', subtitleResourceId: 'CS.Subtitle', suggestions: it.popup };
          }
          results.push({ id: it.id, res: p.insertAnnotations({ critiques: [crit] }) });
        });
        return ctx.sync().then(function () {
          var map = {};
          results.forEach(function (r) { (r.res.value || []).forEach(function (annId) { map[annId] = r.id; }); });
          return { supported: true, ids: map };
        });
      });
    }).catch(function (err) {
      if (useResources) return annotate(items, false);
      // Tanpa langganan Microsoft 365 layanan anotasi menolak — jangan dicoba lagi tiap pemeriksaan.
      if (capsCache) capsCache.annotations = false;
      return { supported: false, ids: {}, error: err && err.message };
    });
  }
  function clearAnnotations(ids) {
    if (!caps().annotations || !ids || !ids.length) return Promise.resolve();
    return Word.run(function (ctx) {
      ids.forEach(function (id) { try { var a = ctx.document.getAnnotationById(id); a.delete(); } catch (e) { /* sudah hilang */ } });
      return ctx.sync();
    }).catch(function () { /* anotasi lama mungkin sudah tidak ada */ });
  }
  function onAnnotationEvents(handlers) {
    if (!caps().annotations) return Promise.resolve(false);
    return Word.run(function (ctx) {
      if (handlers.clicked) ctx.document.onAnnotationClicked.add(function (a) { handlers.clicked(a.id); return Promise.resolve(); });
      if (handlers.popup && caps().popup) ctx.document.onAnnotationPopupAction.add(function (a) { handlers.popup(a.id, a.action, a.critiqueSuggestion); return Promise.resolve(); });
      return ctx.sync().then(function () { return true; });
    }).catch(function () { return false; });
  }

  // ---------- DETEKSI PERUBAHAN DOKUMEN ----------
  function onDocumentChanged(cb) {
    if (caps().api16) {
      return Word.run(function (ctx) {
        var h = function () { cb(); return Promise.resolve(); };
        ctx.document.onParagraphChanged.add(h);
        ctx.document.onParagraphAdded.add(h);
        ctx.document.onParagraphDeleted.add(h);
        return ctx.sync();
      }).catch(function () { return fallback(); });
    }
    return fallback();
    function fallback() {
      return new Promise(function (resolve) {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () { cb({ selectionOnly: true }); }, function () { resolve(); });
      });
    }
  }

  root.WordBridge = {
    caps: caps, readParagraphs: readParagraphs, select: select, applySuggestion: applySuggestion, applyMany: applyMany,
    checkItalics: checkItalics, reorderParagraphs: reorderParagraphs, applyConversion: applyConversion,
    applyLinks: applyLinks, removeLinks: removeLinks, insertReference: insertReference,
    annotate: annotate, clearAnnotations: clearAnnotations, onAnnotationEvents: onAnnotationEvents, onDocumentChanged: onDocumentChanged
  };
})(typeof window !== 'undefined' ? window : this);
