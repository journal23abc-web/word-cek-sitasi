/* ============================================================
   TASK PANE UI — Cek Sitasi untuk Word
   Validator (saran ala Paperpal) · Tautan · Konversi APA 7 ↔ IEEE · cek online CrossRef
   ============================================================ */
(function () {
  'use strict';
  var AC = window.AddinCore;
  var WB = window.WordBridge;
  var $ = function (id) { return document.getElementById(id); };

  var CAT_LABEL = { citation: 'Sitasi', reference: 'Referensi', match: 'Kecocokan', online: 'Online', consistency: 'Dokumen' };
  var STYLE_NAME = { apa7: 'APA 7', ieee: 'IEEE' };

  var state = {
    paras: null, st: null, styleId: null, detect: null, result: null,
    suggestions: [], online: {}, onlineRun: 0, onlineTotal: 0, onlineDone: 0,
    filter: 'all', openId: null, busy: false, pending: false, lastHash: null,
    annMap: {}, annIds: [], suppressUntil: 0, convPlan: null, convSnapshot: null, linkPlan: null
  };

  // ---------- pengaturan (tersimpan per perangkat) ----------
  var SETTINGS_KEY = 'cekSitasi.settings.v1';
  var settings = { auto: true, underline: true, online: true, track: false };
  try { var saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}'); Object.keys(saved).forEach(function (k) { settings[k] = saved[k]; }); } catch (e) { /* storage tidak tersedia */ }
  function saveSettings() { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch (e) { /* abaikan */ } }
  var DISMISS_KEY = 'cekSitasi.dismissed.v1';
  var dismissed = new Set();
  try { (JSON.parse(localStorage.getItem(DISMISS_KEY) || '[]')).forEach(function (k) { dismissed.add(k); }); } catch (e) { /* abaikan */ }
  function saveDismissed() { try { localStorage.setItem(DISMISS_KEY, JSON.stringify(Array.from(dismissed).slice(-500))); } catch (e) { /* abaikan */ } }

  // ---------- util ----------
  function esc(t) { var d = document.createElement('div'); d.textContent = t == null ? '' : String(t); return d.innerHTML; }
  function setStatus(msg, kind) {
    var el = $('status');
    el.className = 'status' + (kind === 'err' ? ' err' : '');
    el.innerHTML = (kind === 'busy' ? '<span class="spin"></span>' : '') + esc(msg || '');
  }
  var toastTimer;
  function toast(msg) {
    var t = $('toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2200);
  }
  function hashParas(paras) {
    var h = 0, s = paras.map(function (p) { return p.text; }).join('\u0001');
    for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; }
    return h + ':' + s.length;
  }
  function snapshotOf(indexes) {
    var snap = {};
    indexes.forEach(function (i) { if (state.paras[i]) snap[i] = state.paras[i].text; });
    return snap;
  }
  function errMsg(e) { return (e && (e.message || e.debugInfo && e.debugInfo.message)) || String(e); }

  // ---------- tab ----------
  function showTab(name) {
    document.querySelectorAll('.tab').forEach(function (b) {
      var on = b.dataset.tab === name; b.classList.toggle('active', on); b.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    ['validator', 'link', 'convert'].forEach(function (n) { $('tab-' + n).hidden = n !== name; });
    if (name === 'convert') syncConvertTab();
  }

  // ================= PEMERIKSAAN =================
  function runCheck(opts) {
    var o = opts || {};
    if (state.busy) { state.pending = true; return Promise.resolve(); }
    state.busy = true;
    if (!o.silent) setStatus('Memeriksa dokumen…', 'busy');
    return WB.readParagraphs().then(function (paras) {
      var hash = hashParas(paras);
      if (!o.force && hash === state.lastHash) { if (!o.silent) setStatus(statusLine()); return; }
      state.lastHash = hash;
      state.paras = paras;
      state.st = AC.structure(paras);
      state.detect = AC.detectStyle(state.st);
      var sel = $('styleSel').value;
      state.styleId = sel === 'auto' ? state.detect.styleId : sel;
      state.result = AC.validate(state.st, state.styleId);
      state.suggestions = state.result.suggestions.slice();
      renderBadge();
      return WB.checkItalics(state.result.italicPlan).then(function (bad) {
        bad.forEach(function (b) {
          var r = state.st.refs[b.refIndex];
          state.suggestions.push({
            id: 'it' + b.refIndex, key: 'italic§' + AC.norm(r.line).slice(0, 200), kind: 'italic', category: 'reference', severity: 'suggestion',
            title: 'Cetak miring belum sesuai ' + STYLE_NAME[state.styleId],
            description: 'Bagian berikut seharusnya dicetak miring: ' + b.segments.join('; ') + '.',
            paraIndex: r.paraIndex, offset: r.leading, length: Math.max(1, r.text.trim().length), refIndex: b.refIndex,
            segments: b.segments, original: r.line, canApply: true, actionLabel: 'Cetak miring'
          });
        });
      }).catch(function () { /* versi Word lama: lewati cek cetak miring */ }).then(function () {
        renderAll();
        if (!o.silent) setStatus(statusLine());
        if (settings.underline) refreshAnnotations();
        if (settings.online) runOnline();
      });
    }).catch(function (e) {
      setStatus('Gagal membaca dokumen: ' + errMsg(e), 'err');
    }).then(function () {
      state.busy = false;
      if (state.pending) { state.pending = false; runCheck({ silent: true }); }
    });
  }

  function statusLine() {
    if (!state.st) return '';
    var st = state.st;
    var t = st.refs.length + ' referensi · ' + (state.result ? state.result.stats.citations : 0) + ' sitasi';
    if (st.headingIndex < 0) t = 'Heading daftar pustaka tidak ditemukan';
    return t + ' · diperiksa ' + new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  }

  function renderBadge() {
    var b = $('styleBadge'), d = state.detect;
    b.hidden = false;
    var auto = $('styleSel').value === 'auto';
    b.textContent = auto ? 'Terdeteksi ' + STYLE_NAME[state.styleId] + (d && d.confidence ? ' · ' + d.confidence + '%' : '') : 'Manual: ' + STYLE_NAME[state.styleId];
    b.title = d && d.note ? d.note : '';
  }

  // ---------- cek online (CrossRef) ----------
  function runOnline() {
    if (!state.result) return;
    var run = ++state.onlineRun;
    var refs = state.result.references;
    var styleId = state.styleId;
    var queue = refs.filter(function (r) { return !state.online[styleId + '|' + r.raw]; });
    state.onlineTotal = refs.length; state.onlineDone = refs.length - queue.length;
    renderSummary();
    var st = state.st;
    var workers = [0, 1, 2].map(function () {
      return (function next() {
        if (run !== state.onlineRun || !queue.length) return Promise.resolve();
        var ref = queue.shift();
        var key = styleId + '|' + ref.raw;
        return AC.verifyReference(st, styleId, ref).then(function (res) {
          state.online[key] = res || { kind: 'none' };
        }).catch(function (e) {
          state.online[key] = { kind: 'error', message: errMsg(e) };
        }).then(function () {
          if (run !== state.onlineRun) return;
          state.onlineDone++;
          renderAll(true);
          return next();
        });
      })();
    });
    return Promise.all(workers).then(function () {
      if (run === state.onlineRun && settings.underline) refreshAnnotations();
    });
  }

  // Hasil online disimpan per teks referensi; posisinya dihitung ulang terhadap dokumen terkini.
  function onlineSuggestions() {
    if (!settings.online || !state.result) return { cards: [], verified: 0, errors: 0 };
    var cards = [], verified = 0, errors = 0;
    state.result.references.forEach(function (ref) {
      var res = state.online[state.styleId + '|' + ref.raw];
      if (!res || res.kind === 'none') return;
      if (res.kind === 'error') { errors++; return; }
      if (res.kind === 'verified') { verified++; return; }
      var idx = AC.refIndexOfRaw(state.st, res.original);
      if (idx < 0) return;
      var r = state.st.refs[idx];
      var c = Object.assign({}, res, { refIndex: idx, paraIndex: r.paraIndex, offset: r.leading, length: Math.max(1, r.text.trim().length) });
      cards.push(c);
    });
    return { cards: cards, verified: verified, errors: errors };
  }

  function allSuggestions() {
    var on = onlineSuggestions();
    // Kartu "tenang" (mis. DOI tidak ditemukan) hanya tampil di filter Online.
    return { list: state.suggestions.concat(on.cards).filter(function (s) { return !dismissed.has(s.key) && (!s.quiet || state.filter === 'online'); }), online: on };
  }

  // ================= RENDER VALIDATOR =================
  function renderSummary(online) {
    var all = allSuggestions();
    var list = all.list;
    $('summary').hidden = false;
    $('sumErr').textContent = list.filter(function (s) { return s.severity === 'error'; }).length;
    $('sumSug').textContent = list.filter(function (s) { return s.severity !== 'error'; }).length;
    var on = online || all.online;
    if (!settings.online) $('sumOnl').textContent = 'Mati';
    else if (state.onlineTotal && state.onlineDone < state.onlineTotal) $('sumOnl').textContent = state.onlineDone + '/' + state.onlineTotal;
    else if (on.errors && !on.verified) { $('sumOnl').textContent = 'Gagal'; $('sumOnl').title = 'CrossRef tidak dapat dihubungi'; }
    else $('sumOnl').textContent = on.verified + ' ✓';
    var n = list.length;
    $('tabCount').textContent = n ? (n > 99 ? '99+' : n) : '';
  }

  function renderAll(fromOnline) {
    if (!state.st) return;
    var all = allSuggestions();
    renderSummary(all.online);
    var notice = [];
    if (state.st.headingIndex < 0) notice.push('Heading daftar pustaka (mis. "Daftar Pustaka" / "References") tidak ditemukan di paragraf tersendiri — hanya sitasi di teks yang diperiksa.');
    if (state.detect && state.detect.note && $('styleSel').value === 'auto') notice.push(state.detect.note);
    if (state.result && state.result.failedLines.length) notice.push(state.result.failedLines.length + ' baris daftar pustaka tidak dapat dibaca dan tidak ikut diperiksa.');
    $('notice').hidden = !notice.length; $('notice').innerHTML = notice.map(esc).join('<br>');

    var counts = { all: all.list.length };
    all.list.forEach(function (s) { counts[s.category] = (counts[s.category] || 0) + 1; });
    $('filters').hidden = !all.list.length && !fromOnline;
    document.querySelectorAll('.chip').forEach(function (c) {
      var f = c.dataset.filter;
      c.innerHTML = esc(c.textContent.replace(/\s*\d+$/, '')) + (counts[f] ? ' <span class="n">' + counts[f] + '</span>' : '');
      c.classList.toggle('active', f === state.filter);
    });

    var visible = all.list.filter(function (s) { return state.filter === 'all' || s.category === state.filter; });
    visible.sort(function (a, b) {
      var sa = a.severity === 'error' ? 0 : 1, sb = b.severity === 'error' ? 0 : 1;
      if (a.paraIndex == null && b.paraIndex != null) return 1;
      if (b.paraIndex == null && a.paraIndex != null) return -1;
      return (a.paraIndex - b.paraIndex) || ((a.offset || 0) - (b.offset || 0)) || (sa - sb);
    });
    state.visible = visible;
    var applicable = visible.filter(function (s) { return s.canApply && (s.kind === 'replace' || s.kind === 'paragraph' || s.kind === 'italic' || s.kind === 'append'); });
    $('bulkBar').hidden = applicable.length < 2;
    $('bulkText').textContent = applicable.length + ' saran bisa diterapkan otomatis';

    var host = $('cards');
    // simpan panel pencarian yang sedang terbuka agar tidak hilang saat hasil online masuk
    var openSearch = host.querySelector('.search:not([hidden])');
    var openSearchId = openSearch ? openSearch.closest('.card').dataset.id : null;
    if (openSearchId && fromOnline) return;
    host.innerHTML = visible.map(cardHtml).join('');
    $('empty').hidden = visible.length > 0;
    if (!visible.length) {
      $('empty').innerHTML = all.list.length ? 'Tidak ada saran di kategori ini.' :
        '<span class="big">✓</span>Tidak ada masalah sitasi yang terdeteksi.' + (state.onlineDone < state.onlineTotal ? '<br>Cek online masih berjalan…' : '');
    }
    var vb = $('verifiedBox');
    vb.hidden = !(settings.online && all.online.verified && (state.filter === 'all' || state.filter === 'online'));
    vb.textContent = all.online.verified + ' referensi ber-DOI terverifikasi cocok dengan CrossRef.' + (all.online.errors ? ' ' + all.online.errors + ' tidak bisa dicek (jaringan).' : '');
  }

  function diffHtml(a, b) {
    return AC.wordDiff(a, b).map(function (op) {
      if (op.op === 'eq') return '<span class="ctx">' + esc(op.text) + '</span>';
      return '<span class="' + (op.op === 'del' ? 'del' : 'ins') + '">' + esc(op.text) + '</span>';
    }).join('');
  }
  function runsHtml(runs) { return runs.map(function (r) { return r.italic ? '<i>' + esc(r.text) + '</i>' : esc(r.text); }).join(''); }

  function cardHtml(s) {
    var cls = 'card sev-' + (s.severity || 'suggestion') + ' cat-' + s.category + (state.openId === s.id ? ' open focus' : '');
    var h = '<article class="' + cls + '" data-id="' + esc(s.id) + '">';
    h += '<div class="card-head" data-act="toggle"><span class="cat">' + esc(CAT_LABEL[s.category] || '') + '</span><span class="card-title">' + esc(s.title) + '</span></div>';
    if (s.kind === 'replace' && s.original) {
      h += '<div class="diff compact"><span class="' + (s.replacement ? 'del' : 'ctx') + '">' + esc(s.original) + '</span>';
      if (s.replacement) h += '<span class="to">→</span><span class="ins">' + esc(s.replacement) + '</span>';
      h += '</div>';
    } else if ((s.kind === 'paragraph' || s.kind === 'append') && s.replacement) {
      h += '<div class="diff">' + diffHtml(s.compareOriginal || s.original, s.replacement) + '</div>';
    } else if (s.original && s.kind !== 'sort' && s.kind !== 'restyle') {
      h += '<div class="diff ctx">' + esc(s.original.length > 220 ? s.original.slice(0, 220) + '…' : s.original) + '</div>';
    }
    h += '<div class="more">';
    if (s.description) h += '<p class="desc">' + esc(s.description) + '</p>';
    if (s.reasons && s.reasons.length > 1) h += '<ul class="reasons">' + s.reasons.map(function (r) { return '<li>' + esc(r.title) + (r.description ? ' — ' + esc(r.description) : '') + '</li>'; }).join('') + '</ul>';
    if (s.hint) h += '<p class="desc">' + esc(s.hint) + '</p>';
    if (s.kind === 'paragraph' && s.replacementRuns && s.replacementRuns.some(function (r) { return r.italic; })) h += '<p class="desc">Hasil: ' + runsHtml(s.replacementRuns) + '</p>';
    h += '</div>';
    if (s.blockedReason) h += '<div class="blocked">' + esc(s.blockedReason) + '</div>';
    h += '<div class="actions">';
    if (s.canApply) h += '<button class="accept" data-act="accept">' + esc(s.actionLabel || 'Terima') + '</button>';
    if (s.searchable) h += '<button class="ghost small" data-act="search">Cari referensi online</button>';
    if (s.paraIndex != null) h += '<button class="ghost small" data-act="locate">Lihat</button>';
    h += '<button class="ghost small" data-act="dismiss">Abaikan</button>';
    if (s.link) h += '<a class="link" href="' + esc(s.link) + '" target="_blank" rel="noopener">Buka DOI ↗</a>';
    if (s.searchUrl) h += '<a class="link" href="' + esc(s.searchUrl) + '" target="_blank" rel="noopener">Cari manual ↗</a>';
    if (s.sourceUrl) h += '<a class="link" href="' + esc(s.sourceUrl) + '" target="_blank" rel="noopener">Pedoman ↗</a>';
    h += '</div>';
    if (s.searchable) h += searchPanelHtml(s);
    h += '</article>';
    return h;
  }

  // ---------- panel cari referensi untuk sitasi yang belum ada di daftar pustaka ----------
  var STOP = /^(?:yang|dan|di|ke|dari|pada|untuk|dengan|dalam|ini|itu|juga|atau|oleh|sebagai|adalah|akan|telah|bahwa|the|a|an|of|and|in|on|for|to|with|by|is|are|was|were|that|this|as|at|from|also|et|al|menurut|studi|penelitian|terbaru|lain|hal|serupa|relevan|ada|pula|showed|study|studies|research)$/i;
  function searchPanelHtml(s) {
    var m = (s.original || '').match(/^\(?\s*([\p{L}'’ .-]+?)(?:\s+(?:et\s+al\.?|&|and|dan)\b.*?)?,?\s+(\d{4})/u);
    var author = m ? m[1].trim() : '', year = m ? m[2] : '';
    var para = state.st && state.st.paragraphs[s.paraIndex] ? state.st.paragraphs[s.paraIndex].text : '';
    var before = para.slice(Math.max(0, s.offset - 160), s.offset);
    var sentence = before.split(/[.!?]\s/).pop() || '';
    var kw = (sentence.match(/[\p{L}-]{4,}/gu) || []).filter(function (w) { return !STOP.test(w); }).slice(-6).join(' ');
    return '<div class="search" hidden>' +
      '<div class="grid"><input data-f="author" placeholder="Penulis" value="' + esc(author) + '"><input data-f="year" placeholder="Tahun" value="' + esc(year) + '">' +
      '<input class="kw" data-f="kw" placeholder="Kata kunci judul/topik" value="' + esc(kw) + '"></div>' +
      '<div class="actions"><button class="accept" data-act="dosearch">Cari di CrossRef</button></div><div class="cands"></div></div>';
  }

  function doSearch(card, s) {
    var panel = card.querySelector('.search');
    var q = { author: panel.querySelector('[data-f=author]').value.trim(), year: panel.querySelector('[data-f=year]').value.trim(), keywords: panel.querySelector('[data-f=kw]').value.trim() };
    var out = panel.querySelector('.cands');
    out.innerHTML = '<div class="desc"><span class="spin"></span>Mencari…</div>';
    AC.searchCandidates(q).then(function (cands) {
      if (!cands.length) { out.innerHTML = '<div class="desc">Tidak ada hasil. Coba ubah kata kunci atau kosongkan tahun.</div>'; return; }
      panel._cands = cands;
      out.innerHTML = cands.map(function (c, i) {
        var f = AC.formatCrossrefReference(c.item, state.styleId, state.st.refs.length + 1);
        c.formatted = f;
        return '<div class="cand"><div class="t">' + esc(c.title || '(tanpa judul)') + '</div>' +
          '<div class="m">' + esc([c.authors, c.year, c.container].filter(Boolean).join(' · ')) + '</div>' +
          '<div class="preview">' + runsHtml(f.runs) + '</div>' +
          '<div class="actions"><button class="accept" data-act="insertref" data-i="' + i + '">Sisipkan ke daftar pustaka</button>' +
          (c.doi ? '<a class="link" href="https://doi.org/' + esc(c.doi) + '" target="_blank" rel="noopener">DOI ↗</a>' : '') + '</div></div>';
      }).join('');
    }).catch(function (e) { out.innerHTML = '<div class="desc">Pencarian gagal: ' + esc(errMsg(e)) + '</div>'; });
  }

  // ---------- aksi kartu ----------
  function findSuggestion(id) { return (state.visible || []).find(function (s) { return s.id === id; }) || allSuggestions().list.find(function (s) { return s.id === id; }); }

  function accept(s) {
    var opts = { track: settings.track };
    var st = state.st;
    var p;
    if (s.kind === 'sort') p = WB.reorderParagraphs(st.refs.map(function (r) { return r.paraIndex; }), s.order, snapshotOf(st.refs.map(function (r) { return r.paraIndex; })), opts);
    else if (s.kind === 'restyle') {
      var plan = AC.planConversion(st, state.styleId, state.styleId);
      if (plan.error) return Promise.reject(new Error(plan.error));
      p = WB.applyConversion(plan, snapshotOf(planParas(plan)), Object.assign({ hangingIndent: false }, opts));
    } else p = WB.applySuggestion(s, opts);
    state.suppressUntil = Date.now() + 2500;
    return p.then(function () { toast('Diterapkan'); return runCheck({ force: true, silent: true }); });
  }

  function planParas(plan) {
    var set = {};
    plan.edits.forEach(function (e) { set[e.paraIndex] = 1; });
    plan.references.paraIndexes.forEach(function (i) { set[i] = 1; });
    return Object.keys(set).map(Number);
  }

  function onCardClick(ev) {
    var btn = ev.target.closest('[data-act]');
    var card = ev.target.closest('.card');
    if (!card) return;
    var s = findSuggestion(card.dataset.id);
    if (!s) return;
    var act = btn ? btn.dataset.act : 'toggle';
    if (act === 'toggle') {
      state.openId = state.openId === s.id ? null : s.id;
      document.querySelectorAll('.card').forEach(function (c) { var on = c.dataset.id === state.openId; c.classList.toggle('open', on); c.classList.toggle('focus', on); });
      if (state.openId && s.paraIndex != null) WB.select(s).catch(function () {});
      return;
    }
    if (act === 'locate') { WB.select(s).catch(function (e) { setStatus(errMsg(e), 'err'); }); return; }
    if (act === 'dismiss') { dismissed.add(s.key); saveDismissed(); renderAll(); refreshAnnotations(); toast('Diabaikan'); return; }
    if (act === 'accept') {
      card.classList.add('applying');
      accept(s).catch(function (e) { card.classList.remove('applying'); setStatus(errMsg(e), 'err'); });
      return;
    }
    if (act === 'search') { var panel = card.querySelector('.search'); panel.hidden = !panel.hidden; if (!panel.hidden) doSearch(card, s); return; }
    if (act === 'dosearch') { doSearch(card, s); return; }
    if (act === 'insertref') {
      var c = card.querySelector('.search')._cands[+btn.dataset.i];
      var point = AC.insertionPoint(state.st, state.styleId, c.formatted.apaLine);
      var f = state.styleId === 'ieee' ? AC.formatCrossrefReference(c.item, 'ieee', point.number || state.st.refs.length + 1) : c.formatted;
      btn.disabled = true;
      state.suppressUntil = Date.now() + 2500;
      WB.insertReference(point, f.runs, { track: settings.track, headingText: 'Daftar Pustaka' }).then(function () {
        toast('Referensi disisipkan — periksa kembali isinya');
        return runCheck({ force: true, silent: true });
      }).catch(function (e) { btn.disabled = false; setStatus(errMsg(e), 'err'); });
    }
  }

  function acceptAll() {
    var list = (state.visible || []).filter(function (s) { return s.canApply && (s.kind === 'replace' || s.kind === 'paragraph' || s.kind === 'italic' || s.kind === 'append'); });
    if (!list.length) return;
    setStatus('Menerapkan ' + list.length + ' saran…', 'busy');
    state.suppressUntil = Date.now() + 4000;
    WB.applyMany(list, { track: settings.track }).then(function (r) {
      toast(r.done + ' saran diterapkan' + (r.failed ? ', ' + r.failed + ' dilewati' : ''));
      return runCheck({ force: true });
    }).catch(function (e) { setStatus(errMsg(e), 'err'); });
  }

  // ---------- anotasi di dokumen ----------
  function refreshAnnotations() {
    var old = state.annIds; state.annIds = []; state.annMap = {};
    var clear = WB.clearAnnotations(old);
    if (!settings.underline || !state.st) return clear;
    var items = allSuggestions().list.filter(function (s) { return s.paraIndex != null && s.length > 0; }).slice(0, 250).map(function (s) {
      return { id: s.id, paraIndex: s.paraIndex, offset: s.offset || 0, length: s.length,
        color: s.category === 'online' ? 'online' : s.severity === 'error' ? 'error' : 'suggestion',
        popup: s.kind === 'replace' && s.replacement && s.replacement.length < 120 ? [s.replacement] : null };
    });
    return clear.then(function () { return WB.annotate(items, true); }).then(function (res) {
      if (!res || !res.supported) return;
      state.annMap = res.ids; state.annIds = Object.keys(res.ids);
    }).catch(function () { /* anotasi opsional */ });
  }
  function onAnnotationClicked(annId) {
    var sid = state.annMap[annId]; if (!sid) return;
    showTab('validator');
    state.openId = sid; state.filter = 'all'; renderAll();
    var el = document.querySelector('.card[data-id="' + sid + '"]'); if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
  function onAnnotationPopup(annId, action) {
    var sid = state.annMap[annId]; if (!sid) return;
    var s = findSuggestion(sid); if (!s) return;
    if (action === 'Accept' && s.canApply) accept(s).catch(function (e) { setStatus(errMsg(e), 'err'); });
    else if (action === 'Reject') { dismissed.add(s.key); saveDismissed(); renderAll(); }
  }

  // ================= TAUTAN =================
  function linkOpts() { return { citations: $('lkCit').checked, external: $('lkExt').checked, keepLook: $('lkLook').checked }; }
  function ensureFresh() { return state.st ? runCheck({ silent: true }) : runCheck({ force: true }); }

  function linkPreview() {
    return ensureFresh().then(function () {
      var plan = AC.planLinks(state.st, state.styleId, { external: $('lkExt').checked });
      state.linkPlan = plan;
      var o = linkOpts();
      var h = '<div class="stat-row"><span class="stat"><b>' + (o.citations ? plan.citationLinks.length : 0) + '</b> tautan sitasi</span>' +
        '<span class="stat"><b>' + plan.bookmarks.length + '</b> referensi tujuan</span>' +
        '<span class="stat"><b>' + (o.external ? plan.externalLinks.length : 0) + '</b> DOI/URL</span></div>';
      if (plan.error) h += '<div class="blocked">' + esc(plan.error) + '</div>';
      if (plan.unmatched.length) {
        h += '<div class="h">Tidak ditautkan (' + plan.unmatched.length + ')</div><div class="list">' + plan.unmatched.map(function (u) {
          return '<div class="it warn"><b>' + esc(u.text) + '</b> — ' + esc(u.note || '') + '</div>';
        }).join('') + '</div>';
      }
      if (plan.citationLinks.length && o.citations) {
        h += '<div class="h">Contoh</div><div class="list">' + plan.citationLinks.slice(0, 30).map(function (l) {
          var r = state.st.refs[l.refIndex];
          return '<div class="it"><b>' + esc(l.text) + '</b> → ' + esc(r.line.slice(0, 90)) + (r.line.length > 90 ? '…' : '') + '</div>';
        }).join('') + '</div>';
      }
      $('linkResult').innerHTML = h;
      return plan;
    });
  }
  function linkApply() {
    setStatus('Membuat tautan…', 'busy');
    return linkPreview().then(function (plan) {
      var idx = plan.citationLinks.map(function (l) { return l.paraIndex; }).concat(plan.externalLinks.map(function (l) { return l.paraIndex; }), plan.bookmarks.map(function (b) { return b.paraIndex; }));
      state.suppressUntil = Date.now() + 3000;
      return WB.applyLinks(plan, snapshotOf(idx), linkOpts());
    }).then(function (r) {
      setStatus(r.linked + ' tautan dibuat' + (r.missed ? ', ' + r.missed + ' tidak ditemukan di Word' : '') + '.');
      toast('Tautan dibuat — Ctrl+klik sitasi untuk melompat');
    }).catch(function (e) { setStatus(errMsg(e), 'err'); });
  }
  function unlink() {
    setStatus('Menghapus tautan…', 'busy');
    state.suppressUntil = Date.now() + 3000;
    WB.removeLinks().then(function (r) { setStatus(r.removed + ' tautan sitasi & ' + r.bookmarks + ' bookmark dihapus.'); })
      .catch(function (e) { setStatus(errMsg(e), 'err'); });
  }

  // ================= KONVERSI =================
  function syncConvertTab() {
    var from = state.styleId || 'apa7';
    $('cvFrom').textContent = STYLE_NAME[from];
    if ($('cvTo').value === from) $('cvTo').value = from === 'apa7' ? 'ieee' : 'apa7';
  }
  function convPreview() {
    $('btnConvApply').disabled = true;
    var out = $('convResult');
    return ensureFresh().then(function () {
      syncConvertTab();
      var source = state.styleId, target = $('cvTo').value;
      var st = state.st;
      var meta = Promise.resolve(null);
      if ($('cvEnrich').checked && st.refs.length) {
        out.innerHTML = '<div class="desc">Melengkapi data dari CrossRef…</div><div class="progress"><div id="cvProg"></div></div>';
        meta = AC.buildEnrichmentMetadata(st, source, null, function (d, t) { var b = $('cvProg'); if (b) b.style.width = Math.round(d / t * 100) + '%'; }).catch(function () { return null; });
      }
      return meta.then(function (metadata) {
        if (st !== state.st) throw new Error('Dokumen berubah saat pratinjau dibuat. Ulangi.');
        var plan = AC.planConversion(st, source, target, { metadata: metadata });
        if (plan.error) { out.innerHTML = '<div class="blocked">' + esc(plan.error) + '</div>'; return; }
        plan.enriched = metadata ? metadata.entries.length : 0;
        state.convPlan = plan; state.convSnapshot = snapshotOf(planParas(plan));
        renderConvPlan(plan);
        $('btnConvApply').disabled = !(plan.edits.length || plan.references.lines.length) || source === target && !plan.edits.length;
      });
    }).catch(function (e) { out.innerHTML = '<div class="blocked">' + esc(errMsg(e)) + '</div>'; });
  }
  function renderConvPlan(plan) {
    var risky = plan.references.lines.filter(function (l) { return l.risky; });
    var h = '<div class="stat-row"><span class="stat"><b>' + plan.edits.length + '</b> sitasi diubah</span>' +
      '<span class="stat"><b>' + plan.references.lines.length + '</b> referensi</span>' +
      (plan.enriched ? '<span class="stat"><b>' + plan.enriched + '</b> dilengkapi CrossRef</span>' : '') +
      (plan.unmatched.length ? '<span class="stat"><b>' + plan.unmatched.length + '</b> tak cocok</span>' : '') + '</div>';
    if (plan.skippedEdits) h += '<div class="blocked">' + plan.skippedEdits + ' sitasi melintasi batas paragraf dan akan dilewati.</div>';
    if (risky.length) h += '<div class="blocked">' + risky.length + ' entri gagal dibaca dengan aman (nama penulis hilang: ' + esc(risky.map(function (l) { return l.lostNames.join(', '); }).join('; ')) + '). Entri ini akan ditulis apa adanya dan disorot kuning untuk diperbaiki manual.</div>';
    if (plan.edits.length) {
      h += '<div class="h">Sitasi dalam teks</div><div class="list">' + plan.edits.map(function (e) {
        return '<div class="it"><span class="del">' + esc(e.original) + '</span> <span class="to">→</span> <span class="ins">' + esc(e.replacement) + '</span></div>';
      }).join('') + '</div>';
    }
    h += '<div class="h">Daftar pustaka (' + STYLE_NAME[plan.target] + ')</div><div class="list">' + plan.references.lines.map(function (l) {
      if (l.risky) { l.runs = [{ text: l.original.replace(/^\s*(?:\[\d+\]|\d+\.)\s*/, ''), italic: false }]; if (plan.target === 'ieee' && !plan.references.keepListNumbering) l.runs.unshift({ text: (l.text.match(/^\[\d+\]\s/) || [''])[0], italic: false }); }
      return '<div class="it' + (l.risky ? ' warn' : '') + '">' + runsHtml(l.runs) + (l.wasCited ? '' : ' <em class="ctx">(tidak disitasi)</em>') + '</div>';
    }).join('') + '</div>';
    if (plan.unmatched.length) h += '<div class="h">Tidak diubah</div><div class="list">' + plan.unmatched.map(function (u) { return '<div class="it warn"><b>' + esc(u.raw) + '</b> — ' + esc(u.note || '') + '</div>'; }).join('') + '</div>';
    if (plan.notes.length) h += '<details><summary class="ctx">Catatan konversi (' + plan.notes.length + ')</summary><div class="list">' + plan.notes.map(function (n) { return '<div class="it">' + esc(n.note) + '<br><span class="ctx">' + esc(n.raw.slice(0, 120)) + '</span></div>'; }).join('') + '</div></details>';
    $('convResult').innerHTML = h;
  }
  function convApply() {
    var plan = state.convPlan; if (!plan) return;
    $('btnConvApply').disabled = true;
    setStatus('Menerapkan konversi…', 'busy');
    state.suppressUntil = Date.now() + 5000;
    WB.applyConversion(plan, state.convSnapshot, { track: settings.track, hangingIndent: $('cvHang').checked }).then(function (r) {
      toast('Konversi ke ' + STYLE_NAME[plan.target] + ' selesai');
      $('styleSel').value = plan.target;
      state.convPlan = null;
      $('convResult').innerHTML = '<div class="verified">' + r.applied + ' sitasi dan ' + plan.references.lines.length + ' referensi ditulis ulang ke ' + STYLE_NAME[plan.target] + '.' + (r.skipped ? ' ' + r.skipped + ' sitasi dilewati — cek tab Validator.' : '') + '</div>';
      return runCheck({ force: true }).then(function () { if ($('cvLink').checked) return linkApply(); });
    }).catch(function (e) { setStatus(errMsg(e), 'err'); $('btnConvApply').disabled = false; });
  }

  // ================= INISIALISASI =================
  var debounceTimer;
  function onDocChanged(info) {
    if (!settings.auto || Date.now() < state.suppressUntil) return;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(function () { runCheck({ silent: true }); }, info && info.selectionOnly ? 2500 : 1500);
  }

  function bind() {
    $('btnCheck').addEventListener('click', function () { runCheck({ force: true }); });
    $('styleSel').addEventListener('change', function () { state.online = {}; runCheck({ force: true }); });
    document.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.dataset.tab); }); });
    document.querySelector('.tabs').addEventListener('keydown', function (e) {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      var tabs = Array.from(document.querySelectorAll('.tab')); var i = tabs.indexOf(document.activeElement); if (i < 0) return;
      var n = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length]; n.focus(); showTab(n.dataset.tab);
    });
    $('filters').addEventListener('click', function (e) { var c = e.target.closest('.chip'); if (!c) return; state.filter = c.dataset.filter; renderAll(); });
    $('cards').addEventListener('click', onCardClick);
    $('btnAcceptAll').addEventListener('click', acceptAll);
    $('btnLinkPreview').addEventListener('click', function () { linkPreview().catch(function (e) { setStatus(errMsg(e), 'err'); }); });
    $('btnLinkApply').addEventListener('click', linkApply);
    $('btnUnlink').addEventListener('click', unlink);
    $('btnConvPreview').addEventListener('click', convPreview);
    $('btnConvApply').addEventListener('click', convApply);
    $('cvTo').addEventListener('change', function () { state.convPlan = null; $('btnConvApply').disabled = true; });
    [['setAuto', 'auto'], ['setUnderline', 'underline'], ['setOnline', 'online'], ['setTrack', 'track']].forEach(function (x) {
      var el = $(x[0]); el.checked = !!settings[x[1]];
      el.addEventListener('change', function () {
        settings[x[1]] = el.checked; saveSettings();
        if (x[1] === 'underline') refreshAnnotations();
        if (x[1] === 'online') { renderAll(); if (el.checked) runOnline(); }
      });
    });
  }

  function start(hostOk) {
    bind();
    if (!hostOk) { setStatus('Buka panel ini dari Microsoft Word.', 'err'); return; }
    WB.onAnnotationEvents({ clicked: onAnnotationClicked, popup: onAnnotationPopup });
    WB.onDocumentChanged(onDocChanged);
    runCheck({ force: true });
  }

  if (window.Office && Office.onReady && !window.CEK_SITASI_DEMO) {
    Office.onReady(function (info) {
      try { if (Office.context && Office.context.officeTheme && Office.context.officeTheme.isDarkTheme) document.documentElement.dataset.theme = 'dark'; } catch (e) { /* abaikan */ }
      start(info.host === Office.HostType.Word);
    });
  } else if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { start(true); });
  } else start(true);
})();
