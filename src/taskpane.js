/* UI task pane "Template Paste" */
(function () {
  "use strict";
  const E = window.TplEngine;
  const STORE = "tplpaste.v1";
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmtN = (n) => Number(n || 0).toLocaleString("id-ID");

  const state = {
    inWord: false,
    store: { active: null, profiles: {} },
    profile: null,
    mode: "section",
    pending: {}, // key -> {blocks:[], overrides:{}, items:[], images:0}
    backup: {},  // key -> OOXML isi yang dihapus (hanya selama panel terbuka)
    found: {},
    open: {},
  };

  /* ---------------- penyimpanan profil ---------------- */
  function loadStore() {
    try {
      const raw = localStorage.getItem(STORE);
      if (raw) state.store = JSON.parse(raw);
    } catch (e) { /* localStorage tidak tersedia */ }
    const st = state.store;
    if (!st.profiles) st.profiles = {};
    // migrasi: nama lama JEPSD -> JEPSI
    Object.keys(st.profiles).forEach((n) => {
      const p = st.profiles[n];
      if (p.base === "jepsd") p.base = "jepsi";
      if (n.startsWith("JEPSD")) {
        const nn = n.replace(/^JEPSD/, "JEPSI");
        p.name = nn; st.profiles[nn] = p; delete st.profiles[n];
        if (st.active === n) st.active = nn;
      }
    });
    if (st.presetsAdded) st.presetsAdded = st.presetsAdded.map((b) => (b === "jepsd" ? "jepsi" : b));
    // tandai profil lama yang berasal dari template bawaan
    Object.values(st.profiles).forEach((p) => { if (!p.base) { const pr = E.PRESETS.find((x) => x.name === p.name); p.base = pr ? pr.base : "aieds"; } });
    // tambahkan template bawaan yang belum pernah ditambahkan (yang sudah dihapus pengguna tidak dimunculkan lagi)
    st.presetsAdded = st.presetsAdded || Object.values(st.profiles).map((p) => p.base);
    E.PRESETS.forEach((pr) => {
      if (st.presetsAdded.includes(pr.base) && Object.keys(st.profiles).length) return;
      let n = pr.name, k = 2;
      while (st.profiles[n]) n = pr.name + " " + k++;
      st.profiles[n] = Object.assign(E.clone(pr), { name: n });
      if (!st.presetsAdded.includes(pr.base)) st.presetsAdded.push(pr.base);
    });
    // profil lama: aturan kutipan diganti ke APA 7 (versi lama: italic, tanpa penanganan atribusi)
    Object.values(st.profiles).forEach((p) => { if (p.quote && !p.quote.apa) delete p.quote; });
    // profil lama: tambahkan kartu ABSTRACT bila belum ada
    Object.values(st.profiles).forEach((p) => {
      if (Array.isArray(p.sections) && !p.sections.some((x) => x.kind === "abstract")) {
        p.sections.unshift(E.clone(E.presetFor(p.base).sections.find((x) => x.kind === "abstract")));
      }
    });
    if (!st.active || !st.profiles[st.active]) st.active = Object.keys(st.profiles)[0];
    state.profile = E.mergeProfile(st.profiles[st.active]);
    state.mode = state.profile.layout.insertMode || "section";
  }
  let saveTimer = null;
  function saveStore() {
    state.store.profiles[state.store.active] = state.profile;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { localStorage.setItem(STORE, JSON.stringify(state.store)); } catch (e) { /* abaikan */ } }, 250);
    $("#profileLabel").textContent = state.store.active;
  }

  /* ---------------- toast ---------------- */
  let toastTimer;
  function toast(msg, kind, ms) {
    const t = $("#toast");
    t.textContent = msg;
    t.className = "toast show " + (kind || "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.className = "toast"), ms || (kind === "err" ? 7000 : 3500));
  }
  function errMsg(e) {
    let m = (e && e.message) || String(e);
    if (e && e.debugInfo && e.debugInfo.errorLocation) m += " (" + e.debugInfo.errorLocation + ")";
    return m;
  }
  function needWord() {
    if (!state.inWord) { toast("Buka add-in ini dari Microsoft Word untuk menyisipkan ke dokumen.", "err"); return false; }
    return true;
  }

  /* ---------------- tab ---------------- */
  function initTabs() {
    document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((x) => { x.classList.toggle("active", x === b); x.setAttribute("aria-selected", x === b ? "true" : "false"); });
      $("#content").scrollTop = 0;
      ["whole", "paste", "words", "check", "math", "template"].forEach((k) => $("#tab-" + k).classList.toggle("hidden", k !== b.dataset.tab));
      if (b.dataset.tab === "words") recount();
      if (b.dataset.tab === "check") renderCheckSettings();
    }));
  }

  /* ---------------- mode sisip ---------------- */
  function initMode() {
    const hints = {
      section: "Isi ditambahkan di bawah konten yang sudah ada pada bagian tersebut. Judul bagian dibuat otomatis bila belum ada.",
      cursor: "Isi disisipkan di posisi kursor. Bila kursor di baris kosong, isi masuk tepat di baris itu.",
    };
    const sync = () => {
      document.querySelectorAll("#modeSeg button").forEach((b) => b.classList.toggle("on", b.dataset.mode === state.mode));
      $("#modeHint").textContent = hints[state.mode];
    };
    document.querySelectorAll("#modeSeg button").forEach((b) => b.addEventListener("click", () => {
      state.mode = b.dataset.mode; state.profile.layout.insertMode = state.mode; saveStore(); sync();
    }));
    sync();
  }

  /* ---------------- kartu bagian ---------------- */
  const TYPE_OPTS = [
    ["body", "Isi"], ["heading2", "Heading 2"], ["heading3", "Heading 3"],
    ["quote", "Kutipan"], ["quoteAttr", "Atribusi"], ["equation", "Persamaan"], ["caption", "Caption"], ["note", "Catatan"], ["ref", "Referensi"], ["table", "Tabel"], ["skip", "Lewati"],
  ];
  const typeKey = (it) => (it.type === "heading" ? "heading" + (it.level || 2) : it.type);

  function pend(key) {
    return state.pending[key] || (state.pending[key] = { blocks: [], overrides: {}, items: [], images: 0 });
  }
  function sec(key) { return state.profile.sections.find((s) => s.key === key); }

const isAbs = (key) => { const x = sec(key); return !!x && x.kind === "abstract"; };
  function reclassify(key) {
    const p = pend(key);
    p.items = isAbs(key) ? E.classifyAbstract(p.blocks, sec(key), state.profile) : E.classify(p.blocks, sec(key), state.profile);
    Object.entries(p.overrides).forEach(([i, v]) => {
      const it = p.items[i];
      if (!it) return;
      if (v.startsWith("heading")) { it.type = "heading"; it.level = +v.slice(7); }
      else it.type = v;
    });
  }

  function runsHtml(runs, max) {
    let out = "", n = 0;
    for (const r of runs || []) {
      if (n >= max) break;
      let t = r.text.slice(0, max - n); n += t.length;
      t = esc(t);
      if (r.sub) t = "<sub>" + t + "</sub>";
      if (r.sup) t = "<sup>" + t + "</sup>";
      if (r.i) t = "<i>" + t + "</i>";
      if (r.b) t = "<b>" + t + "</b>";
      out += t;
    }
    return out + (n >= max ? "…" : "");
  }

  function itemWords(items) {
    return items.filter((i) => i.type !== "skip").reduce((a, i) => a + (i.type === "table" ? 0 : E.countWords((i.label ? i.label + " " : "") + E.plain(i.runs))), 0);
  }

  function renderSections() {
    const host = $("#sections");
    host.innerHTML = "";
    for (const s of state.profile.sections) {
      const p = pend(s.key);
      const live = p.items.filter((i) => i.type !== "skip");
      const card = document.createElement("section");
      card.className = "sec" + (state.open[s.key] ? " open" : "");
      card.dataset.key = s.key;
      const abs = s.kind === "abstract";
      const docPill = state.inWord ? (state.found[s.key] ? '<span class="pill indoc" title="Judul bagian ditemukan di dokumen">ada</span>'
        : abs ? '<span class="pill new" title="Judul ABSTRACT tidak ditemukan; gunakan mode Posisi kursor">tidak ada</span>' : '<span class="pill new" title="Judul bagian akan dibuat otomatis">baru</span>') : "";
      const words = itemWords(p.items), absLim = +state.profile.limits.abstractWords || 0;
      const readyPill = !live.length ? "" : abs
        ? `<span class="pill ready${absLim && words > absLim ? " over" : ""}">${fmtN(words)}${absLim ? " / " + fmtN(absLim) : ""} kata</span>`
        : `<span class="pill ready">${live.length} siap · ${fmtN(words)} kata</span>`;
      card.innerHTML = `
        <div class="sec-head">
          <svg class="chev" viewBox="0 0 16 16"><path d="M6 3l5 5-5 5" fill="none" stroke="currentColor" stroke-width="1.8"/></svg>
          <div class="sec-title">${esc(s.title)}</div>
          ${readyPill}
          ${docPill}
        </div>
        <div class="sec-body">
          <div class="drop" contenteditable="true" spellcheck="false" data-key="${esc(s.key)}">Klik di sini lalu tekan Ctrl+V (⌘V) untuk menempel ${abs ? "abstrak (menggantikan abstrak yang ada di dokumen)" : "isi " + esc(titleCase(s.title))}</div>
          ${p.images ? `<p class="warnline">${p.images} gambar di tempelan diabaikan — sisipkan gambar langsung di Word.</p>` : ""}
          <div class="preview ${p.items.length ? "" : "hidden"}"></div>
          <div class="sec-meta ${p.items.length ? "" : "hidden"}">${summary(p.items)}</div>
          <div class="sec-actions">
            <button class="btn primary small act-insert" ${live.length ? "" : "disabled"}>${abs ? "Ganti abstrak di dokumen" : "Sisipkan ke dokumen"}</button>
            ${abs ? "" : '<button class="btn small act-reformat" title="Terapkan format template ke isi yang sudah ada di bagian ini">Rapikan isi yang ada</button>'}
            <button class="btn small ghost act-clear" ${p.items.length ? "" : "disabled"} title="Kosongkan tempelan di panel ini (dokumen tidak berubah)">Kosongkan tempelan</button>
          </div>
          <div class="sec-actions danger-zone">
            <button class="btn small danger act-wipe" ${state.inWord && !state.found[s.key] ? "disabled" : ""} title="Hapus semua isi bagian ini di dokumen; judul bagian tetap ada">${abs ? "Kosongkan abstrak di dokumen" : "Hapus isi bagian ini di dokumen"}</button>
            ${state.backup[s.key] ? `<button class="btn small act-restore" title="Kembalikan isi yang terakhir dihapus">↶ Kembalikan isi yang dihapus</button>` : ""}
          </div>
        </div>`;
      host.appendChild(card);
      renderPreview(card, s.key);

      card.querySelector(".sec-head").addEventListener("click", () => { state.open[s.key] = !state.open[s.key]; card.classList.toggle("open"); });
      const drop = card.querySelector(".drop");
      drop.addEventListener("paste", (ev) => onPaste(ev, s.key, drop));
      drop.addEventListener("keydown", (ev) => { if (!(ev.ctrlKey || ev.metaKey) && ev.key.length === 1) ev.preventDefault(); });
      drop.addEventListener("drop", (ev) => ev.preventDefault());
      card.querySelector(".act-insert").addEventListener("click", () => doInsert([s.key]));
      if (!abs) card.querySelector(".act-reformat").addEventListener("click", () => doReformat(s.key));
      card.querySelector(".act-clear").addEventListener("click", () => { state.pending[s.key] = null; renderSections(); });
      card.querySelector(".act-wipe").addEventListener("click", () => doWipe([s.key]));
      const rb = card.querySelector(".act-restore");
      if (rb) rb.addEventListener("click", () => doRestore(s.key));
    }
  }

  function summary(items) {
    if (items.some((i) => i.type === "abstract")) {
      const lab = items.filter((i) => i.type === "abstract" && i.label).length;
      const sk = items.filter((i) => i.type === "skip" && i.why === "keywords").length;
      return (lab ? `Abstrak terstruktur: ${lab} label dibuat bold.` : "Abstrak satu bagian.") + (sk ? " Baris Keywords tidak dimasukkan." : "") + " Abstrak lama di dokumen akan diganti (bisa dikembalikan).";
    }
    const c = (t, l) => items.filter((i) => i.type === t && (l == null || i.level === l)).length;
    const parts = [];
    if (c("body")) parts.push(c("body") + " paragraf");
    if (c("heading")) parts.push(c("heading") + " sub-heading");
    if (c("ref")) parts.push(c("ref") + " referensi");
    if (c("caption")) parts.push(c("caption") + " caption");
    if (c("table")) parts.push(c("table") + " tabel");
    if (c("note")) parts.push(c("note") + " catatan");
    if (c("quote")) parts.push(c("quote") + " kutipan");
    if (c("quoteAttr")) parts.push(c("quoteAttr") + " atribusi (digabung ke kutipan)");
    if (c("equation")) parts.push(c("equation") + " persamaan");
    return parts.join(" · ") + (parts.length ? ". Ubah jenis lewat menu di kiri bila deteksi keliru." : "");
  }

  function renderPreview(card, key) {
    const p = pend(key);
    const box = card.querySelector(".preview");
    box.innerHTML = "";
    p.items.forEach((it, i) => {
      const row = document.createElement("div");
      row.className = "prow t-" + it.type + (it.level === 3 ? " l3" : "");
      const sel = document.createElement("select");
      const opts = isAbs(key) ? [["abstract", "Abstrak"], ["skip", "Lewati"]] : TYPE_OPTS;
      opts.forEach(([v, l]) => {
        if (v === "table" && it.type !== "table") return;
        if (it.type === "table" && !["table", "skip"].includes(v)) return;
        const o = document.createElement("option"); o.value = v; o.textContent = l; sel.appendChild(o);
      });
      sel.value = typeKey(it);
      sel.addEventListener("change", () => {
        p.overrides[i] = sel.value; reclassify(key);
        const open = state.open[key]; renderSections(); state.open[key] = open;
      });
      const txt = document.createElement("div");
      txt.className = "txt";
      const mAst = it.type === "equation" && it.math && it.math.ast && window.TplMath;
      if (mAst) txt.innerHTML = `<math>${TplMath.toMathML(it.math.ast)}</math>` + (it.math.number ? ` <span class="eqnum">(${esc(it.math.number)})</span>` : "");
      else txt.innerHTML = it.type === "abstract" && it.label ? "<b>" + esc(it.label) + " –</b> " + runsHtml(it.runs, 160) : it.type === "table" ? esc(`Tabel ${it.rows.length} baris × ${Math.max(...it.rows.map((r) => r.length))} kolom: ${it.rows[0].join(" | ")}`).slice(0, 200) : runsHtml(it.runs, 180);
      row.append(sel, txt);
      box.appendChild(row);
    });
  }

  function titleCase(s) { return E.titleCase(String(s).toLowerCase()); }

  function onPaste(ev, key, drop) {
    ev.preventDefault();
    const cd = ev.clipboardData || window.clipboardData;
    const clip = { html: cd.getData("text/html") || "", text: cd.getData("text/plain") || cd.getData("Text") || "" };
    const parsed = E.parseInput(clip, sec(key), state.profile);
    if (!parsed.blocks.length) { toast("Tidak ada teks yang terbaca dari clipboard.", "err"); return; }
    const p = pend(key);
    p.blocks = p.blocks.concat(parsed.blocks);
    p.images += parsed.images || 0;
    reclassify(key);
    state.open[key] = true;
    renderSections();
    const d = document.querySelector(`.drop[data-key="${CSS.escape(key)}"]`);
    if (d) { d.classList.add("flash"); d.textContent = `✓ ${parsed.blocks.length} blok ditambahkan — tempel lagi untuk menambah`; setTimeout(() => d.classList.remove("flash"), 1500); }
  }

  /* ---------------- Tempel keseluruhan ---------------- */
  function onWholePaste(ev) {
    ev.preventDefault();
    const cd = ev.clipboardData || window.clipboardData;
    const clip = { html: cd.getData("text/html") || "", text: cd.getData("text/plain") || cd.getData("Text") || "" };
    if (!clip.html && !clip.text.trim()) { toast("Clipboard kosong.", "err"); return; }
    splitWholeNow(clip, false);
  }
  function splitWholeNow(clip, quiet) {
    const r = E.splitWhole(clip, state.profile);
    if (!r.order.length) {
      state.whole = null; renderWhole();
      if (!quiet) toast("Judul bagian (INTRODUCTION, METHOD, …) tidak ditemukan. Pastikan judul bagian ditulis KAPITAL seperti di template.", "err", 8000);
      return;
    }
    state.wholeClip = clip;
    // seluruh tempelan bagian sebelumnya diganti dengan hasil pecahan ini
    state.profile.sections.forEach((x) => { state.pending[x.key] = null; });
    r.order.forEach((k) => { const p = pend(k); p.blocks = r.parts[k].blocks; p.images = 0; reclassify(k); });
    state.whole = r;
    renderWhole(); renderSections();
    const d = $("#wholeDrop");
    if (!quiet && d) { d.classList.add("flash"); d.textContent = `✓ Dipecah ke ${r.order.length} bagian — tempel lagi untuk mengganti`; setTimeout(() => d.classList.remove("flash"), 1500); }
  }
  function renderWhole() {
    const host = $("#wholeResult"), r = state.whole;
    $("#wholeActions").hidden = !r;
    if (!r) { host.innerHTML = ""; return; }
    const rows = state.profile.sections.map((s) => {
      const p = state.pending[s.key];
      if (!p || !p.items.length) return `<li class="wr miss"><span class="wt">${esc(s.title)}</span><span class="wn">tidak ada di tempelan</span></li>`;
      const its = p.items.filter((i) => i.type !== "skip"), c = (t) => its.filter((i) => i.type === t).length;
      const bits = [];
      if (s.kind === "references") bits.push(c("ref") + " referensi");
      else if (s.kind === "abstract") bits.push(fmtN(itemWords(its)) + " kata");
      else {
        bits.push(its.length + " blok", fmtN(itemWords(its)) + " kata");
        if (c("heading")) bits.push(c("heading") + " sub-heading");
        if (c("table")) bits.push(c("table") + " tabel");
        if (c("equation")) bits.push(c("equation") + " rumus");
      }
      return `<li class="wr ok"><span class="wt">${esc(s.title)}</span><span class="wn">${bits.join(" · ")}</span></li>`;
    }).join("");
    const notes = [];
    if (r.front) notes.push(`${r.front} baris sebelum judul bagian pertama (judul artikel, penulis, afiliasi) dilewati.`);
    r.unknown.forEach((u) => {
      const sec = u.key && state.profile.sections.find((x) => x.key === u.key);
      notes.push(sec ? `Judul "${u.title}" tidak ada di template → dijadikan sub-heading di ${sec.title}.` : `Judul "${u.title}" tidak ada di template → isinya dilewati.`);
    });
    if (r.images) notes.push(`${r.images} gambar diabaikan — sisipkan gambar langsung di Word.`);
    host.innerHTML = `<ul class="wlist">${rows}</ul>` + (notes.length ? `<div class="wnotes">${notes.map((n) => `<p>${esc(n)}</p>`).join("")}</div>` : "");
  }
  async function wholeInsert() {
    if (!state.whole) return;
    const prev = state.mode;
    state.mode = "section"; // tempel keseluruhan selalu di akhir tiap bagian
    try { await doInsert(state.profile.sections.map((s) => s.key)); } finally { state.mode = prev; }
    renderWhole();
  }
  /** Rapikan semua bagian yang ada di dokumen, satu per satu. */
  async function tidyAll() {
    if (!needWord()) return;
    setBusy(true);
    const host = $("#tidyResult");
    host.innerHTML = '<span class="muted">Merapikan…</span>';
    const rows = [], dups = [];
    let ok = 0, fail = 0, skipped = 0;
    try {
      await refreshScan();
      for (const s of state.profile.sections) {
        if (s.kind === "abstract") continue;
        if (!state.found[s.key]) { skipped++; continue; }
        try {
          const r = await WordOps.reformatSection(s.key, state.profile);
          ok++;
          let d;
          if (s.kind === "references") {
            d = `${r.formatted - (r.dupRemoved || 0)} referensi` + (r.dupRemoved ? `, ${r.dupRemoved} ganda dihapus` + (r.renumbered ? `, ${r.citations} sitasi disesuaikan` : "") : "");
            (r.dupList || []).forEach((x) => dups.push(x));
          } else {
            const bits = [`${r.formatted} paragraf`];
            if (r.headings) bits.push(`${r.headings} sub-heading`);
            if (r.tables) bits.push(`${r.tables} tabel`);
            if (r.quotes) bits.push(`${r.quotes} kutipan`);
            if (r.equations) bits.push(`${r.equations} rumus`);
            if (r.tableErrors) bits.push(`${r.tableErrors} tabel gagal`);
            d = bits.join(" · ");
          }
          rows.push(`<div class="tr"><span>${esc(s.title)}</span><span class="tn">${esc(d)}</span></div>`);
        } catch (e) {
          fail++;
          rows.push(`<div class="tr err"><span>${esc(s.title)}</span><span class="tn">gagal: ${esc(errMsg(e))}</span></div>`);
        }
      }
      host.innerHTML = rows.join("") + (dups.length ? `<details><summary>${dups.length} referensi ganda yang dihapus</summary>${dups.map((x) => `<p>${esc(x)}…</p>`).join("")}</details>` : "");
      toast(ok ? `✓ ${ok} bagian dirapikan${fail ? `, ${fail} gagal` : ""}${dups.length ? `, ${dups.length} referensi ganda dihapus` : ""}.` + (skipped ? ` ${skipped} bagian tidak ada di dokumen.` : "") : "Tidak ada bagian template yang ditemukan di dokumen.", fail || !ok ? "err" : "ok", 7000);
    } catch (e) { host.innerHTML = ""; toast("Gagal merapikan: " + errMsg(e), "err"); }
    finally { setBusy(false); }
  }
  function initWhole() {
    $("#tidyAll").addEventListener("click", tidyAll);
    const d = $("#wholeDrop");
    d.addEventListener("paste", onWholePaste);
    d.addEventListener("keydown", (ev) => { if (!(ev.ctrlKey || ev.metaKey) && ev.key.length === 1) ev.preventDefault(); });
    d.addEventListener("drop", (ev) => ev.preventDefault());
    $("#wholeInsert").addEventListener("click", wholeInsert);
    $("#wholeReview").addEventListener("click", () => document.querySelector('.tab[data-tab="paste"]').click());
    $("#wholeClear").addEventListener("click", () => {
      state.profile.sections.forEach((x) => { state.pending[x.key] = null; });
      state.whole = null; state.wholeClip = null;
      $("#wholeDrop").textContent = "Klik di sini lalu tekan Ctrl+V (⌘V) untuk menempel seluruh artikel";
      renderWhole(); renderSections();
    });
  }

  /* ---------------- aksi Word ---------------- */
  async function doInsert(keys) {
    if (!needWord()) return;
    const todo = keys.filter((k) => pend(k).items.some((i) => i.type !== "skip"));
    if (!todo.length) { toast("Belum ada isi yang ditempel."); return; }
    if (state.mode === "cursor" && todo.length > 1) { toast("Mode kursor hanya untuk satu bagian sekaligus. Pilih 'Akhir bagian'.", "err"); return; }
    setBusy(true);
    try {
      let last = null, total = 0, eqText = 0, afNo = false;
      for (const k of todo) {
        const items = pend(k).items;
        if (isAbs(k)) {
          const r = await WordOps.insertAbstract(items, state.profile, state.mode);
          if (r.backup) state.backup[k] = r.backup;
          total += r.inserted;
          const lim = +state.profile.limits.abstractWords || 0;
          if (lim && r.words > lim) toast(`Abstrak ${fmtN(r.words)} kata — melebihi batas ${fmtN(lim)}.`, "err");
          state.pending[k] = null;
          if (todo.length === 1) { await refreshScan(); renderSections(); toast(`✓ Abstrak diganti (${fmtN(r.words)} kata${lim ? " / " + fmtN(lim) : ""}). Abstrak lama bisa dikembalikan.`, lim && r.words > lim ? "err" : "ok", 6000); return; }
          continue;
        }
        last = await WordOps.insertSection(k, items, state.profile, state.mode);
        total += last.inserted; eqText += last.eqFallback || 0; if (last.autofitUnsupported) afNo = true;
        if (last.missingStyles && last.missingStyles.length) toast("Gaya tidak ditemukan, dipakai Normal: " + last.missingStyles.join(", "), "err");
        state.pending[k] = null;
      }
      await refreshScan();
      renderSections();
      const c = last.count;
      const over = c.limit && c.total > c.limit;
      toast(`✓ ${total} blok disisipkan. Total isi ${fmtN(c.total)}${c.limit ? " / " + fmtN(c.limit) : ""} kata${over ? " — MELEBIHI BATAS" : ""}.`
        + (afNo ? " AutoFit to contents belum didukung Word versi ini — lebar tabel dibiarkan (atur lewat Layout → AutoFit)." : "")
        + (eqText ? ` ${eqText} rumus ditolak Word sebagai persamaan dan disisipkan sebagai teks — blok barisnya lalu pakai Rumus → "Ubah teks terpilih jadi rumus".` : ""), over || eqText ? "err" : "ok", eqText ? 10000 : 6000);
      if (!$("#tab-words").classList.contains("hidden")) renderCounts(c);
    } catch (e) {
      console.error(e);
      toast("Gagal menyisipkan: " + errMsg(e), "err");
    } finally { setBusy(false); }
  }

  /** Hapus isi bagian di dokumen (judul tetap). */
  async function doWipe(keys) {
    if (!needWord()) return;
    await refreshScan();
    const targets = keys.filter((k) => state.found[k]);
    if (!targets.length) { toast("Judul bagian tidak ditemukan di dokumen."); return; }
    const words = targets.reduce((a, k) => {
      const r = state.lastCount && state.lastCount.sections.find((x) => x.key === k);
      return a + (r ? r.words : 0);
    }, 0);
    const names = targets.map((k) => sec(k).title).join(", ");
    const msg = targets.length === 1 && isAbs(targets[0])
      ? "Kosongkan teks abstrak di dokumen? Judul ABSTRACT dan isi kotak lain (keywords, article history) tidak berubah."
      : targets.length === 1
      ? `Hapus SEMUA isi ${names} di dokumen (±${fmtN(words)} kata, termasuk sub-heading, tabel & gambar)? Judul bagian tetap ada.`
      : `Hapus SEMUA isi ${targets.length} bagian (${names}) — ±${fmtN(words)} kata, termasuk tabel & gambar? Judul-judul bagian tetap ada.`;
    if (!(await ask(msg, null, { ok: "Hapus", danger: true }))) return;
    setBusy(true);
    try {
      let total = 0;
      // dari bagian paling bawah ke atas
      const order = state.profile.sections.map((x) => x.key).filter((k) => targets.includes(k)).reverse();
      for (const k of order) {
        const r = isAbs(k) ? await WordOps.clearAbstract(state.profile) : await WordOps.clearSection(k, state.profile);
        if (r.ooxml) state.backup[k] = r.ooxml;
        total += r.deleted;
      }
      await refreshScan();
      renderSections();
      toast(total ? `✓ Isi dihapus (${total} paragraf). Bisa dipulihkan dengan "Kembalikan" atau Ctrl+Z.` : "Bagian sudah kosong.", total ? "ok" : "", 6000);
    } catch (e) { toast("Gagal menghapus: " + errMsg(e), "err"); } finally { setBusy(false); }
  }

  async function doRestore(key) {
    if (!needWord() || !state.backup[key]) return;
    setBusy(true);
    try {
      if (isAbs(key)) await WordOps.restoreAbstract(state.backup[key], state.profile);
      else await WordOps.restoreSection(key, state.backup[key], state.profile);
      delete state.backup[key];
      await refreshScan(); renderSections();
      toast("✓ Isi " + sec(key).title + " dikembalikan.", "ok");
    } catch (e) { toast("Gagal mengembalikan: " + errMsg(e) + ". Coba Ctrl+Z.", "err"); } finally { setBusy(false); }
  }

  async function doReformat(key) {
    if (!needWord()) return;
    setBusy(true);
    try {
      const r = await WordOps.reformatSection(key, state.profile);
      if (sec(key) && sec(key).kind === "references") {
        toast(`✓ ${r.formatted - (r.dupRemoved || 0)} referensi dirapikan.` + (r.dupRemoved ? ` ${r.dupRemoved} referensi ganda dihapus${r.renumbered ? `, nomor diurutkan ulang dan ${r.citations} sitasi di teks disesuaikan` : ""}.` : " Tidak ada referensi ganda."), "ok", r.dupRemoved ? 9000 : 5000);
        return;
      }
      const warn = r.tableErrors || r.equationErrors || r.mathType;
      toast(`✓ ${r.formatted} paragraf dirapikan (${r.headings} sub-heading${r.quotes ? `, ${r.quotes} kutipan` : ""}${r.tables ? `, ${r.tables} tabel` : ""}${r.equations ? `, ${r.equations} rumus diubah jadi persamaan Word` : ""}${r.inlineMath ? `, ${r.inlineMath} rumus dalam kalimat dirapikan` : ""}).`
        + (r.tableErrors ? ` ${r.tableErrors} tabel gagal dirapikan (mungkin sel gabungan rumit).` : "")
        + (r.autofitUnsupported ? " AutoFit to contents belum didukung Word versi ini — lebar tabel dibiarkan." : "")
        + (r.equationErrors ? ` ${r.equationErrors} rumus gagal diubah — coba blok barisnya lalu pakai tab Rumus.` : "")
        + (r.mathType ? ` ${r.mathType} rumus berupa objek MathType/Equation 3.0 — tidak bisa diubah otomatis; klik dua kali objeknya lalu simpan ulang sebagai persamaan Word.` : ""), warn ? "err" : "ok", warn ? 9000 : 5000);
    } catch (e) { toast(errMsg(e), "err"); } finally { setBusy(false); }
  }

  async function doUndo() {
    if (!needWord()) return;
    try {
      const ok = await WordOps.undoLast();
      toast(ok ? "✓ Sisipan terakhir dihapus." : "Tidak ada sisipan yang bisa dihapus.", ok ? "ok" : "");
      await refreshScan(); renderSections();
    } catch (e) { toast(errMsg(e), "err"); }
  }

  function setBusy(b) {
    document.querySelectorAll("button").forEach((x) => {
      if (b) { x.dataset.wasDisabled = x.disabled ? "1" : ""; x.disabled = true; }
      else if (x.dataset.wasDisabled !== undefined) { x.disabled = x.dataset.wasDisabled === "1"; delete x.dataset.wasDisabled; }
    });
    document.body.style.cursor = b ? "progress" : "";
  }

  async function refreshScan() {
    if (!state.inWord) return;
    try { const r = await WordOps.scan(state.profile); state.found = r.found; state.lastCount = r.count; } catch (e) { console.warn(e); }
  }

  /* ---------------- hitung kata ---------------- */
  async function recount() {
    renderLimits();
    if (!state.inWord) { renderCounts(null); return; }
    try { state.lastCount = await WordOps.countWords(state.profile); renderCounts(state.lastCount); }
    catch (e) { toast(errMsg(e), "err"); }
  }

  /** Aturan hitung kata — tersimpan per profil jurnal. */
  function renderLimits() {
    const L = state.profile.limits;
    L.excludeKeys = L.excludeKeys || [];
    const host = $("#limitSettings");
    const unset = !L.bodyWords && !L.abstractWords && !L.minBodyWords;
    const num = (k, label, step) => `<label for="lim_${k}">${label}</label><input type="number" min="0" step="${step}" id="lim_${k}" data-k="${k}" value="${+L[k] || 0}">`;
    const chk = (id, label, on, attrs) => `<label for="${id}">${label}</label><input type="checkbox" id="${id}" ${attrs} ${on ? "checked" : ""}>`;
    host.innerHTML = `<fieldset class="limits"><details ${unset || state.limitsOpen ? "open" : ""}>
      <summary>Aturan hitung kata · ${esc(state.store.active.split(" (")[0])}</summary>
      <div class="fgrid">
        ${num("bodyWords", "Maksimal kata isi", 100)}
        ${num("minBodyWords", "Minimal kata isi", 100)}
        ${num("abstractWords", "Maksimal kata abstrak", 10)}
        <div class="fdesc">0 = tidak dibatasi. Isi sesuai ketentuan jurnal; tersimpan untuk profil ini saja.</div>
        ${chk("lim_tables", "Teks di dalam tabel ikut dihitung", L.countTables, 'data-b="countTables"')}
        <div class="subhead">Bagian yang ikut dihitung</div>
        ${state.profile.sections.filter((s) => s.kind !== "abstract").map((s) => {
          const on = s.kind === "references" ? L.countReferences : !L.excludeKeys.includes(s.key);
          return chk("lim_s_" + s.key, esc(s.title), on, `data-sec="${esc(s.key)}" data-ref="${s.kind === "references" ? 1 : ""}"`);
        }).join("")}
      </div></details></fieldset>
      ${unset ? '<p class="warnline">Batas kata untuk jurnal ini belum diatur.</p>' : ""}`;
    host.querySelector("details").addEventListener("toggle", (e) => { state.limitsOpen = e.target.open; });
    host.querySelectorAll("input").forEach((inp) => inp.addEventListener("change", () => {
      if (inp.dataset.k) L[inp.dataset.k] = Math.max(0, parseInt(inp.value, 10) || 0);
      else if (inp.dataset.b) L[inp.dataset.b] = inp.checked;
      else if (inp.dataset.sec) {
        if (inp.dataset.ref) L.countReferences = inp.checked;
        else {
          L.excludeKeys = L.excludeKeys.filter((k) => k !== inp.dataset.sec);
          if (!inp.checked) L.excludeKeys.push(inp.dataset.sec);
        }
      }
      state.limitsOpen = true;
      saveStore();
      recount();
    }));
  }

  function renderCounts(c) {
    const T = $("#wordTotals"), W = $("#wordTable");
    if (!c) { T.innerHTML = '<p class="hint">Buka di Word untuk menghitung kata dokumen.</p>'; W.innerHTML = ""; return; }
    const card = (label, n, lim, min) => {
      const over = lim && n > lim, under = min && n < min;
      const pct = lim ? Math.min(100, (n / lim) * 100) : 0;
      let info = lim ? "/ " + fmtN(lim) + " kata" : "kata";
      if (over) info += " · lebih " + fmtN(n - lim);
      else if (lim) info += " · sisa " + fmtN(lim - n);
      if (under) info += " · kurang " + fmtN(min - n) + " dari minimal " + fmtN(min);
      return `<div class="total-card ${over ? "over" : under ? "under" : ""}"><div class="of">${label}</div>
        <div><span class="big">${fmtN(n)}</span> <span class="of">${info}</span></div>
        ${lim ? `<div class="bar ${over ? "over" : ""}"><span style="width:${pct}%"></span></div>` : ""}</div>`;
    };
    T.innerHTML = card("Isi artikel (bagian yang dicentang)", c.total, c.limit, c.minLimit) + card("Abstrak", c.abstract, c.abstractLimit, 0);
    W.innerHTML = `<table class="wtable">${c.sections.map((s) => `<tr class="${s.counted ? "" : "nc"}"><td>${esc(s.title)}</td><td class="n">${fmtN(s.words)}</td></tr>`).join("")}</table>
      ${c.sections.length ? "" : '<p class="hint">Belum ada judul bagian yang dikenali di dokumen.</p>'}`;
  }

  /* ---------------- cek kesesuaian ---------------- */
  function renderCheckSettings() {
    const P = state.profile;
    P.check = Object.assign({ maxMonths: 3, minDays: 0, requireRevised: true }, P.check || {});
    P.journal = Object.assign({ name: "", abbrev: "" }, P.journal || {});
    const host = $("#checkSettings");
    host.innerHTML = `<fieldset class="limits"><details ${state.checkOpen ? "open" : ""}>
      <summary>Aturan cek · ${esc(state.store.active.split(" (")[0])}</summary>
      <div class="fgrid">
        <label for="ck_journal" class="full">Nama jurnal (untuk dicocokkan dengan cite)</label>
        <input type="text" id="ck_journal" class="full" value="${esc(P.journal.name)}">
        <label for="ck_max">Maks. Received → Accepted (bulan)</label><input type="number" id="ck_max" min="0" step="1" value="${+P.check.maxMonths || 0}">
        <label for="ck_min">Min. Received → Accepted (hari)</label><input type="number" id="ck_min" min="0" step="1" value="${+P.check.minDays || 0}">
        <label for="ck_rev">Tanggal Revised wajib ada</label><input type="checkbox" id="ck_rev" ${P.check.requireRevised ? "checked" : ""}>
        <div class="fdesc">0 bulan = jarak tidak dibatasi. Tersimpan untuk profil jurnal ini.</div>
      </div></details></fieldset>`;
    host.querySelector("details").addEventListener("toggle", (e) => { state.checkOpen = e.target.open; });
    const save = () => {
      P.journal.name = $("#ck_journal").value.trim();
      P.check.maxMonths = Math.max(0, parseInt($("#ck_max").value, 10) || 0);
      P.check.minDays = Math.max(0, parseInt($("#ck_min").value, 10) || 0);
      P.check.requireRevised = $("#ck_rev").checked;
      saveStore();
      if (state.lastMeta) renderCheck(TplCheck.run(state.lastMeta, state.profile));
    };
    host.querySelectorAll("input").forEach((i) => i.addEventListener("change", save));
  }

  async function runCheck() {
    if (!needWord()) return;
    setBusy(true);
    try {
      state.lastMeta = await WordOps.collectMeta(state.profile);
      renderCheck(TplCheck.run(state.lastMeta, state.profile));
    } catch (e) { toast("Gagal memeriksa: " + errMsg(e), "err"); } finally { setBusy(false); }
  }

  function renderCheck(res) {
    const F = res.findings;
    const n = (l) => F.filter((f) => f.level === l).length;
    $("#checkSummary").innerHTML = `<div class="chips">
      <span class="chip error">✕ ${n("error")} perlu diperbaiki</span>
      <span class="chip warn">⚠ ${n("warn")} perlu dicek</span>
      <span class="chip ok">✓ ${n("ok")} sesuai</span></div>`;
    const ICON = { ok: "✓", warn: "!", error: "✕", info: "i" };
    const order = { error: 0, warn: 1, info: 2, ok: 3 };
    const groups = [...new Set(F.map((f) => f.group))];
    const p = res.parsed, c = p.cite || {};
    const H = p.history || {};
    const hd = (k) => (H[k] ? esc(H[k].text) + (H[k].date && !H[k].date.invalid ? ` <span class="of">(${TplCheck.fmtDate(H[k].date)})</span>` : "") : "—");
    $("#checkResults").innerHTML = groups.map((g) => {
      const list = F.filter((f) => f.group === g).sort((a, b) => order[a.level] - order[b.level]);
      const bad = list.filter((f) => f.level === "error" || f.level === "warn").length;
      return `<section class="cgroup"><h4>${esc(g)}<span class="cnt">${bad ? bad + " temuan" : "aman"}</span></h4>${list.map((f) =>
        `<div class="finding ${f.level}"><span class="ic">${ICON[f.level]}</span><div><div class="ft">${esc(f.title)}</div>${f.detail ? `<div class="fd">${esc(f.detail)}</div>` : ""}</div></div>`).join("")}</section>`;
    }).join("") + `<section class="cgroup"><details><summary style="padding:8px 10px;font-weight:600;font-size:12.5px;cursor:pointer">Data yang terbaca dari dokumen</summary>
      <table class="parsed"><tbody>
        <tr><td>Judul</td><td>${esc(p.title || "—")}</td></tr>
        <tr><td>Penulis</td><td>${esc(p.authors.join("; ") || "—")}</td></tr>
        <tr><td>Cite: penulis</td><td>${esc(c.authorsStr || "—")}</td></tr>
        <tr><td>Cite: judul</td><td>${esc(c.title || "—")}</td></tr>
        <tr><td>Cite: jurnal</td><td>${esc(c.journal || "—")}</td></tr>
        <tr><td>Cite: vol(no), hlm</td><td>${esc((c.volume || "—") + "(" + (c.issue || "—") + "), " + (c.pageStart || "—") + "–" + (c.pageEnd || "—"))}</td></tr>
        <tr><td>Cite: DOI</td><td>${esc(c.doi || "—")}</td></tr>
        <tr><td>Header</td><td>${p.issue ? esc(`Vol. ${p.issue.volume}, No. ${p.issue.issue} ${p.issue.monthText || ""} ${p.issue.year || ""}`) : "—"}${p.headerDoi ? "<br>" + esc(p.headerDoi) : ""}</td></tr>
        <tr><td>Halaman awal</td><td>${p.startPage || "tidak terbaca"}</td></tr>
        <tr><td>Received</td><td>${hd("received")}</td></tr>
        <tr><td>Revised</td><td>${hd("revised")}</td></tr>
        <tr><td>Accepted</td><td>${hd("accepted")}</td></tr>
      </tbody></table></details></section>`;
  }

  /* ---------------- editor template ---------------- */
  const ALIGN = [["Left", "Kiri"], ["Centered", "Tengah"], ["Right", "Kanan"], ["Justified", "Justify"]];
  const paraFields = (opt) => [
    ["size", "Ukuran font (pt)", "number", { step: 0.5, min: 6 }],
    ...(opt.bold ? [["bold", "Tebal (bold)", "bool"], ["italic", "Miring (italic)", "bool"]] : []),
    ...(opt.caps ? [["caps", "HURUF KAPITAL", "bool"]] : []),
    ["align", "Perataan", "select", { options: ALIGN }],
    ...(opt.indent ? [["firstIndentCm", "Indentasi baris pertama (cm)", "number", { step: 0.01 }]] : []),
    ...(opt.hanging ? [["hangingCm", "Indentasi gantung / hanging (cm)", "number", { step: 0.01 }]] : []),
    ["before", "Spasi sebelum (pt)", "number", { step: 1, min: 0 }],
    ["after", "Spasi sesudah (pt)", "number", { step: 1, min: 0 }],
    ["line", "Spasi baris (1 = tunggal)", "number", { step: 0.05, min: 0.5 }],
    ["style", "Gaya Word (style)", "text", { desc: "Biarkan Normal kecuali jurnal memakai gaya khusus; tulis persis seperti di Word." }],
  ];
  const GROUPS = [
    { title: "Umum", fields: [
      ["font", "Font", "text"], ["color", "Warna teks", "color"],
      ["layout.blankBeforeSection", "Sisakan baris kosong sebelum judul bagian berikutnya", "bool"],
    ] },
    { title: "Judul bagian (INTRODUCTION, METHOD, …)", prefix: "sectionHeading", fields: paraFields({ bold: true, caps: true }) },
    { title: "Sub-heading tingkat 2", prefix: "h2", fields: paraFields({ bold: true }) },
    { title: "Sub-heading tingkat 3", prefix: "h3", fields: paraFields({ bold: true }) },
    { title: "Paragraf isi", prefix: "body", fields: paraFields({ indent: true }) },
    { title: "Caption tabel & gambar", prefix: "caption", fields: [["labelBold", "Label \"Table 1.\" tebal", "bool"], ...paraFields({})] },
    { title: "Catatan tabel (Note.)", prefix: "note", fields: [["labelItalic", "Label \"Note.\" miring", "bool"], ["labelBold", "Label \"Note.\" tebal", "bool"], ["italic", "Seluruh catatan miring", "bool"], ...paraFields({})] },
    { title: "Abstrak", prefix: "abstract", fields: [
      ["size", "Ukuran font (pt)", "number", { step: 0.5, min: 6 }],
      ["italic", "Miring (italic)", "bool"],
      ["align", "Perataan", "select", { options: ALIGN }],
      ["before", "Spasi sebelum (pt)", "number", { step: 1, min: 0 }],
      ["after", "Spasi sesudah (pt)", "number", { step: 1, min: 0 }],
      ["line", "Spasi baris (1 = tunggal)", "number", { step: 0.01, min: 0.5 }],
      ["labelBold", "Label abstrak terstruktur bold (Purpose – …)", "bool"],
      ["labelSep", "Pemisah label (kosong = ikuti tempelan)", "text", { desc: "Contoh: \" – \" (AIEDS) atau \": \"" }],
      ["joinLabeled", "Komponen terstruktur", "select", { options: [["linebreak", "1 paragraf, dipisah line break"], ["paragraphs", "Paragraf terpisah"]] }],
      ["labels", "Label yang dikenali (pisahkan koma)", "list"],
    ] },
    { title: "Kutipan langsung (wawancara)", prefix: "quote", fields: [
      ["minBlockWords", "Kutipan blok bila ≥ … kata (APA: 40)", "number", { step: 1, min: 0 }],
      ["removeMarks", "Hapus tanda kutip pada kutipan blok (APA)", "bool"],
      ["attribution", "Atribusi partisipan", "select", { options: [["paren", "Dalam kurung setelah kutipan (APA)"], ["keep", "Biarkan seperti tempelan"]] }],
      ["mergeAttribution", "Gabungkan baris atribusi terpisah ke kutipan", "bool"],
      ["italic", "Miring (italic) — APA: tidak", "bool"],
      ["leftCm", "Indentasi kiri (cm)", "number", { step: 0.01, min: 0 }],
      ["rightCm", "Indentasi kanan (cm)", "number", { step: 0.01, min: 0 }],
      ...paraFields({}),
    ] },
    { title: "Rumus / persamaan", prefix: "equation", fields: [
      ["autoConvert", "Ubah baris rumus yang ditempel jadi persamaan Word", "bool", { desc: "Contoh: \"Y = β0 + β1X1 + ε (1)\" menjadi persamaan asli. Rumus dari Word/ChatGPT/Wikipedia selalu diubah." }],
      ["numbering", "Penomoran", "select", { options: [["tab", "(1) rata kanan (tab)"], ["after", "(1) langsung setelah rumus"], ["native", "Penomoran Word 365 (#)"], ["none", "Tanpa nomor"]] }],
      ["align", "Posisi rumus", "select", { options: [["Centered", "Tengah"], ["Left", "Kiri"]] }],
      ["textWidthCm", "Lebar area teks (cm) — cadangan bila tidak terbaca dari dokumen", "number", { step: 0.01, min: 5 }],
      ["size", "Ukuran rumus (pt)", "number", { step: 0.5, min: 6 }],
      ["before", "Spasi sebelum (pt)", "number", { step: 1, min: 0 }],
      ["after", "Spasi sesudah (pt)", "number", { step: 1, min: 0 }],
    ] },
    { title: "Tabel yang ditempel", prefix: "table", fields: [
      ["cellAlign", "Perataan isi sel", "select", { options: ALIGN }],
      ["size", "Ukuran font tabel (pt)", "number", { step: 0.5 }], ["headerBold", "Baris judul tebal", "bool"],
      ["apaBorders", "Garis gaya APA (atas, bawah, bawah judul)", "bool"], ["firstColLeft", "Kolom pertama rata kiri", "bool"],
      ["autoFit", "Lebar tabel", "select", { options: [["content", "AutoFit to contents"], ["window", "AutoFit to window (selebar halaman)"], ["none", "Biarkan"]] }],
      ["repeatHeader", "Ulangi baris judul kolom di tiap halaman (Repeat as header row)", "bool"],
      ["align", "Posisi tabel", "select", { options: [["Centered", "Tengah"], ["Left", "Kiri"]] }],
    ] },
    { title: "Referensi", prefix: "references", fields: [
      ...paraFields({ hanging: true }),
      ["citeStyle", "Gaya referensi", "select", { options: [["apa", "APA (Nama, A. (2020). …)"], ["ieee", "IEEE bernomor ([1] A. Nama, \"Judul,\" …)"]] }],
      ["autoItalic", "Italic otomatis (APA: nama jurnal + volume / judul buku; IEEE: nama jurnal / judul buku) bila tempelan tanpa italic", "bool"],
      ["linkDoi", "Jadikan DOI/URL hyperlink", "bool"],
      ["doiFormat", "Penulisan DOI", "select", { options: [["url", "https://doi.org/10.… (APA)"], ["keep", "Biarkan seperti tempelan (doi: 10.…)"]] }],
      ["stripNumbering", "Hapus penomoran [1] / 1. (matikan untuk IEEE)", "bool"],
      ["sort", "Urutkan A–Z (yang ditempel)", "bool"],
    ] },
    { title: "Deteksi otomatis", prefix: "detect", fields: [
      ["maxHeadingWords", "Maks. kata untuk sub-heading", "number", { step: 1, min: 3 }],
      ["stripHeadingNumbers", "Hapus nomor heading (2.1, A.)", "bool"],
      ["titleCaseHeadings", "Ubah sub-heading ke Title Case", "bool"],
      ["keepItalic", "Pertahankan italic dari tempelan", "bool"],
      ["keepBold", "Pertahankan bold di dalam paragraf", "bool"],
      ["keepSupSub", "Pertahankan superscript/subscript", "bool"],
      ["mergeBrokenLines", "Sambung baris terputus (salinan PDF)", "bool"],
      ["detectQuotes", "Kenali kutipan langsung (diawali tanda kutip “…”)", "bool"],
      ["requireCapsSectionTitle", "Judul bagian di dokumen ditulis KAPITAL", "bool", { desc: "Membedakan judul bagian RESULTS dari sub-heading Results." }],
    ] },
  ];

  const getPath = (o, p) => p.split(".").reduce((a, k) => (a == null ? a : a[k]), o);
  const setPath = (o, p, v) => { const ks = p.split("."); let a = o; ks.slice(0, -1).forEach((k) => (a = a[k] = a[k] || {})); a[ks[ks.length - 1]] = v; };

  function renderTemplate() {
    const sel = $("#profileSelect");
    const opts = Object.keys(state.store.profiles).map((n) => `<option ${n === state.store.active ? "selected" : ""}>${esc(n)}</option>`).join("");
    sel.innerHTML = opts;
    $("#journalSelect").innerHTML = opts;
    $("#journalSelectWhole").innerHTML = opts;
    const form = $("#templateForm");
    form.innerHTML = "";
    GROUPS.forEach((g, gi) => {
      const fs = document.createElement("fieldset");
      const det = document.createElement("details");
      if (gi === 0) det.open = true;
      det.innerHTML = `<summary>${esc(g.title)}</summary>`;
      const grid = document.createElement("div");
      grid.className = "fgrid";
      g.fields.forEach(([k, label, type, opt]) => {
        opt = opt || {};
        const path = g.prefix ? g.prefix + "." + k : k;
        const id = "f_" + path.replace(/\./g, "_");
        const val = getPath(state.profile, path);
        const lab = document.createElement("label"); lab.htmlFor = id; lab.textContent = label;
        let inp;
        if (type === "select") {
          inp = document.createElement("select");
          opt.options.forEach(([v, l]) => { const o = document.createElement("option"); o.value = v; o.textContent = l; inp.appendChild(o); });
          inp.value = val;
        } else {
          inp = document.createElement("input");
          inp.type = type === "bool" ? "checkbox" : type;
          if (type === "list") inp.type = "text";
          if (type === "bool") inp.checked = !!val; else inp.value = val == null ? "" : Array.isArray(val) ? val.join(", ") : val;
          if (opt.step) inp.step = opt.step;
          if (opt.min != null) inp.min = opt.min;
        }
        inp.id = id;
        inp.addEventListener("change", () => {
          let v = type === "bool" ? inp.checked : type === "number" ? parseFloat(inp.value) : type === "list" ? inp.value.split(",").map((x) => x.trim()).filter(Boolean) : inp.value;
          if (type === "number" && isNaN(v)) v = 0;
          setPath(state.profile, path, v);
          saveStore();
          Object.keys(state.pending).forEach((key) => state.pending[key] && reclassify(key));
          renderSections();
        });
        grid.append(lab, inp);
        if (opt.desc) { const d = document.createElement("div"); d.className = "fdesc"; d.textContent = opt.desc; grid.appendChild(d); }
      });
      det.appendChild(grid);
      fs.appendChild(det);
      form.appendChild(fs);
    });
    form.appendChild(renderSectionEditor());
  }

  function renderSectionEditor() {
    const fs = document.createElement("fieldset");
    const det = document.createElement("details");
    det.innerHTML = "<summary>Daftar bagian artikel</summary>";
    const box = document.createElement("div");
    box.className = "secedit";
    const list = state.profile.sections;
    list.forEach((s, i) => {
      const el = document.createElement("div");
      el.className = "secitem";
      el.innerHTML = `
        <span class="lab">Judul di dokumen</span><input type="text" data-f="title" value="${esc(s.title)}">
        <span class="lab">Nama lain (pisahkan koma)</span><input type="text" data-f="aliases" value="${esc((s.aliases || []).join(", "))}">
        <span class="lab">Sub-heading tingkat 2 khusus (mis. Results, Discussion)</span><input type="text" data-f="majors" value="${esc((s.majorSubheadings || []).join(", "))}">
        <span class="lab">Ukuran judul bagian ini (pt) — kosong = ikut pengaturan umum</span><input type="number" step="0.5" min="6" data-f="hsize" value="${s.heading && s.heading.size ? s.heading.size : ""}">
        <div class="tools">${s.kind === "abstract" ? '<span class="lab">Abstrak (mengganti isi kotak abstrak)</span>' : `<label><input type="checkbox" data-f="refs" ${s.kind === "references" ? "checked" : ""}> Bagian referensi</label>`}
          <span><button type="button" class="icon-btn" data-a="up" title="Naik">↑</button> <button type="button" class="icon-btn" data-a="down" title="Turun">↓</button> <button type="button" class="icon-btn" data-a="del" title="Hapus">✕</button></span></div>`;
      el.querySelectorAll("input").forEach((inp) => inp.addEventListener("change", () => {
        const f = inp.dataset.f;
        if (f === "title") s.title = inp.value.trim() || s.title;
        if (f === "aliases") s.aliases = inp.value.split(",").map((x) => x.trim()).filter(Boolean);
        if (f === "majors") s.majorSubheadings = inp.value.split(",").map((x) => x.trim()).filter(Boolean);
        if (f === "hsize") { const v = parseFloat(inp.value); if (v > 0) s.heading = Object.assign({}, s.heading, { size: v }); else if (s.heading) delete s.heading.size; }
        if (f === "refs") { if (inp.checked) s.kind = "references"; else delete s.kind; }
        saveStore(); renderSections();
      }));
      el.querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
        const a = b.dataset.a;
        if (a === "del") { ask(`Hapus bagian ${s.title}?`).then((ok) => { if (!ok) return; list.splice(i, 1); saveStore(); renderTemplate(); renderSections(); }); return; }
        if (a === "up" && i > 0) [list[i - 1], list[i]] = [list[i], list[i - 1]];
        if (a === "down" && i < list.length - 1) [list[i + 1], list[i]] = [list[i], list[i + 1]];
        saveStore(); renderTemplate(); renderSections();
        $("#templateForm details:last-of-type").open = true;
      }));
      box.appendChild(el);
    });
    const add = document.createElement("button");
    add.type = "button"; add.className = "btn small"; add.textContent = "+ Tambah bagian";
    add.addEventListener("click", () => {
      const key = "s" + Date.now().toString(36);
      list.splice(Math.max(0, list.findIndex((x) => x.kind === "references")), 0, { key, title: "NEW SECTION", aliases: [] });
      saveStore(); renderTemplate(); renderSections();
      $("#templateForm details:last-of-type").open = true;
    });
    box.appendChild(add);
    det.appendChild(box); fs.appendChild(det);
    return fs;
  }

  // Dialog sendiri — window.confirm/prompt diblokir di Office
  function ask(title, defVal, opt) {
    opt = opt || {};
    return new Promise((resolve) => {
      const d = $("#askDialog"), inp = $("#askInput");
      $("#askTitle").textContent = title;
      $("#askOk").textContent = opt.ok || "OK";
      $("#askOk").classList.toggle("danger-solid", !!opt.danger);
      inp.classList.toggle("hidden", defVal == null);
      inp.value = defVal == null ? "" : defVal;
      const done = (v) => { d.close(); $("#askOk").onclick = $("#askCancel").onclick = null; resolve(v); };
      $("#askOk").onclick = () => done(defVal == null ? true : inp.value.trim());
      $("#askCancel").onclick = () => done(defVal == null ? false : null);
      d.showModal();
      if (defVal != null) inp.focus();
    });
  }

  function initProfiles() {
    const switchProfile = (name) => {
      state.store.active = name;
      state.profile = E.mergeProfile(state.store.profiles[name]);
      saveStore(); renderTemplate();
      // isi yang sudah ditempel tetap ada, diklasifikasi ulang dengan aturan jurnal baru
      Object.keys(state.pending).forEach((k) => { if (state.pending[k]) { if (sec(k)) reclassify(k); else state.pending[k] = null; } });
      renderSections(); renderLimits(); refreshScan().then(renderSections);
      if (state.wholeClip) splitWholeNow(state.wholeClip, true); // tempelan keseluruhan dipecah ulang dengan judul bagian jurnal baru
      toast("Template aktif: " + name.split(" (")[0]);
    };
    $("#profileSelect").addEventListener("change", (e) => switchProfile(e.target.value));
    $("#journalSelect").addEventListener("change", (e) => switchProfile(e.target.value));
    $("#journalSelectWhole").addEventListener("change", (e) => switchProfile(e.target.value));
    $("#profDup").addEventListener("click", () => {
      let n = state.store.active + " (salinan)", k = 2;
      while (state.store.profiles[n]) n = state.store.active + " (salinan " + k++ + ")";
      const p = E.clone(state.profile); p.name = n;
      state.store.profiles[n] = p; state.store.active = n; state.profile = p;
      saveStore(); renderTemplate(); toast("Profil diduplikat: " + n);
    });
    $("#profRename").addEventListener("click", async () => {
      const cur = state.store.active;
      const n = await ask("Nama profil baru:", cur);
      if (!n || n === cur) return;
      if (state.store.profiles[n]) { toast("Nama sudah dipakai.", "err"); return; }
      state.store.profiles[n] = state.profile; delete state.store.profiles[cur];
      state.profile.name = n; state.store.active = n; saveStore(); renderTemplate();
    });
    $("#profDelete").addEventListener("click", async () => {
      if (Object.keys(state.store.profiles).length < 2) { toast("Minimal harus ada satu profil."); return; }
      if (!(await ask("Hapus profil " + state.store.active + "?"))) return;
      delete state.store.profiles[state.store.active];
      state.store.active = Object.keys(state.store.profiles)[0];
      state.profile = E.mergeProfile(state.store.profiles[state.store.active]);
      saveStore(); renderTemplate(); renderSections();
    });
    $("#profReset").addEventListener("click", async () => {
      const pr = E.presetFor(state.profile.base);
      if (!(await ask("Kembalikan profil ini ke pengaturan bawaan " + pr.name.split(" (")[0] + "?"))) return;
      const keep = state.store.active;
      state.profile = E.clone(pr); state.profile.name = keep;
      saveStore(); renderTemplate(); renderSections(); renderLimits(); toast("Profil dikembalikan ke bawaan.");
    });
    const dlg = $("#jsonDialog");
    $("#profExport").addEventListener("click", () => {
      $("#jsonTitle").textContent = "Ekspor profil — salin & simpan teks ini";
      $("#jsonText").value = JSON.stringify(state.profile, null, 2);
      $("#jsonApply").classList.add("hidden");
      dlg.showModal();
    });
    $("#profImport").addEventListener("click", () => {
      $("#jsonTitle").textContent = "Impor profil — tempel JSON profil di bawah";
      $("#jsonText").value = "";
      $("#jsonApply").classList.remove("hidden");
      dlg.showModal();
    });
    $("#jsonCopy").addEventListener("click", () => { $("#jsonText").select(); try { document.execCommand("copy"); toast("Disalin."); } catch (e) { /* */ } });
    $("#jsonClose").addEventListener("click", () => dlg.close());
    $("#jsonApply").addEventListener("click", () => {
      try {
        const p = E.mergeProfile(JSON.parse($("#jsonText").value));
        let n = p.name || "Profil impor", k = 2;
        while (state.store.profiles[n]) n = (p.name || "Profil impor") + " " + k++;
        p.name = n;
        state.store.profiles[n] = p; state.store.active = n; state.profile = p;
        saveStore(); renderTemplate(); renderSections(); dlg.close(); toast("Profil diimpor: " + n, "ok");
      } catch (e) { toast("JSON tidak valid: " + e.message, "err"); }
    });
  }

  /* ---------------- tab Rumus ---------------- */
  const SYMBOLS = [
    ["α", "\\alpha "], ["β", "\\beta "], ["γ", "\\gamma "], ["δ", "\\delta "], ["ε", "\\varepsilon "], ["θ", "\\theta "], ["λ", "\\lambda "], ["μ", "\\mu "], ["σ", "\\sigma "], ["χ", "\\chi "], ["Σ", "\\Sigma "],
    ["x²", "^{2}"], ["xᵢ", "_{i}"], ["a⁄b", "\\frac{}{}"], ["√", "\\sqrt{}"], ["∑ⁿ", "\\sum_{i=1}^{n} "], ["x̄", "\\bar{}"], ["ŷ", "\\hat{}"],
    ["×", "\\times "], ["·", "\\cdot "], ["±", "\\pm "], ["≤", "\\leq "], ["≥", "\\geq "], ["≠", "\\neq "], ["≈", "\\approx "], ["∞", "\\infty "], ["( )", "\\left(  \\right)"],
  ];
  const mathState = { mode: "display", last: null };
  const TM = () => window.TplMath;

  function eqCurrent() {
    const raw = $("#eqInput").value;
    if (!raw.trim()) return null;
    if (mathState.grabbed && mathState.grabbed.src === raw) return mathState.grabbed.res;
    return TM().parse(raw);
  }
  function renderEqPreview() {
    const box = $("#eqPreview");
    const r = eqCurrent();
    if (!r) { box.innerHTML = '<span class="muted">Pratinjau rumus muncul di sini.</span>'; return; }
    if (r.error) { box.innerHTML = `<span class="err">${esc(r.error)}</span>`; return; }
    if (r.number && !$("#eqNumber").dataset.touched) $("#eqNumber").value = r.number;
    const num = mathState.mode === "display" && numberingChoice() !== "none" && $("#eqNumber").value.trim() ? `<span class="eqnum">(${esc($("#eqNumber").value.trim().replace(/^\((.*)\)$/, "$1"))})</span>` : "";
    box.innerHTML = `<div class="eq-render ${mathState.mode}"><math display="${mathState.mode === "display" ? "block" : "inline"}">${TM().toMathML(r.ast)}</math>${num}</div><div class="eq-tex" title="LaTeX">${esc(TM().toLatex(r.ast))}</div>`;
  }
  function numberingChoice() {
    const v = $("#eqNumbering").value;
    return v === "profile" ? ((state.profile.equation || {}).numbering || "tab") : v;
  }
  async function copyText(t, what) {
    try { await navigator.clipboard.writeText(t); toast(what + " disalin.", "ok"); return; } catch (e) { /* WebView Office: pakai cara lama */ }
    const ta = document.createElement("textarea");
    ta.value = t; ta.style.position = "fixed"; ta.style.opacity = "0"; document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
    ta.remove();
    toast(ok ? what + " disalin." : "Tidak bisa menyalin otomatis — salin manual dari kotak rumus.", ok ? "ok" : "err");
  }
  function insertAtCaret(ta, snippet) {
    const a = ta.selectionStart, b = ta.selectionEnd, v = ta.value;
    const sel = v.slice(a, b);
    let ins = snippet, caret;
    const hole = snippet.indexOf("{}");
    if (hole >= 0) { ins = snippet.slice(0, hole + 1) + sel + snippet.slice(hole + 1); caret = a + hole + 1 + sel.length; }
    else if (snippet.includes("  ")) { const k = snippet.indexOf("  ") + 1; ins = snippet.slice(0, k) + sel + snippet.slice(k); caret = a + k + sel.length; }
    else caret = a + ins.length;
    ta.value = v.slice(0, a) + ins + v.slice(b);
    ta.focus(); ta.setSelectionRange(caret, caret);
    mathState.grabbed = null;
    renderEqPreview();
  }
  /** Tempel ke kotak rumus: ambil LaTeX dari HTML (Word/KaTeX/MathJax/Wikipedia) bila ada. */
  function onEqPaste(ev) {
    const cd = ev.clipboardData || window.clipboardData;
    const html = cd && cd.getData("text/html");
    if (!html || !/msEquation|<math|katex|mjx-container|mwe-math/i.test(html)) return; // teks biasa: biarkan
    const parsed = E.parseHtml(html);
    const m = parsed.blocks.map((b) => b.math).filter(Boolean)[0];
    if (!m || !m.ast) return;
    ev.preventDefault();
    const tex = m.tex || TM().toLatex(m.ast);
    $("#eqInput").value = tex + (m.number ? " \\tag{" + m.number + "}" : "");
    mathState.grabbed = m.omml ? { src: $("#eqInput").value, res: { ast: m.ast, tex, number: m.number, omml: m.omml } } : null;
    renderEqPreview();
    toast("Rumus dari " + (m.source === "word" ? "Word" : "web") + " terbaca.", "ok");
  }
  function eqForInsert() {
    const r = eqCurrent();
    if (!r) { toast("Tulis atau tempel rumus dulu.", "err"); return null; }
    if (r.error) { toast(r.error, "err"); return null; }
    const numbering = numberingChoice();
    const number = numbering === "none" ? null : ($("#eqNumber").value.trim() || null);
    return { eq: { ast: r.ast, omml: r.omml, number }, numbering };
  }
  async function eqInsert() {
    if (!needWord()) return;
    const x = eqForInsert(); if (!x) return;
    const prof = Object.assign(E.clone(state.profile), { equation: Object.assign({}, state.profile.equation, { numbering: x.numbering }) });
    setBusy(true);
    try {
      await WordOps.insertEquation(x.eq, prof, { mode: mathState.mode });
      toast(`✓ Rumus disisipkan${x.eq.number && mathState.mode === "display" ? " (" + x.eq.number + ")" : ""}.`, "ok");
      bumpNumber();
    } catch (e) { toast("Gagal menyisipkan rumus: " + errMsg(e), "err"); } finally { setBusy(false); }
  }
  function bumpNumber() {
    const inp = $("#eqNumber"), v = inp.value.trim();
    if (mathState.mode === "display" && /^\d+$/.test(v)) { inp.value = String(+v + 1); inp.dataset.touched = "1"; }
  }
  async function eqConvert() {
    if (!needWord()) return;
    const numbering = numberingChoice();
    const prof = Object.assign(E.clone(state.profile), { equation: Object.assign({}, state.profile.equation, { numbering }) });
    setBusy(true);
    try {
      const r = await WordOps.convertSelection(prof, { number: numbering === "none" ? null : $("#eqNumber").value.trim() });
      toast(r.mode === "inline" ? "✓ Teks terpilih diubah jadi rumus (dalam kalimat)." : `✓ Baris diubah jadi rumus${r.number ? " (" + r.number + ")" : ""}.`, "ok");
      if (r.mode === "display" && r.number) bumpNumber();
    } catch (e) { toast(errMsg(e), "err"); } finally { setBusy(false); }
  }
  async function eqGrab() {
    if (!needWord()) return;
    setBusy(true);
    try {
      const r = await WordOps.readSelectedEquation();
      if (!r) { toast("Tidak ada rumus di posisi kursor. Klik di dalam rumus Word lalu coba lagi.", "err"); return; }
      $("#eqInput").value = r.tex;
      if (r.number) { $("#eqNumber").value = r.number; $("#eqNumber").dataset.touched = "1"; }
      mathState.grabbed = null;
      renderEqPreview();
      toast(`✓ Rumus diambil${r.count > 1 ? ` (${r.count} rumus digabung)` : ""} — bisa disalin sebagai LaTeX/MathML atau diedit.`, "ok");
    } catch (e) { toast("Gagal membaca rumus: " + errMsg(e), "err"); } finally { setBusy(false); }
  }
  async function eqAutoNumber() {
    if (!needWord()) return;
    try { const n = await WordOps.nextEquationNumber(); $("#eqNumber").value = String(n); $("#eqNumber").dataset.touched = "1"; renderEqPreview(); toast("Nomor berikutnya: (" + n + ")"); }
    catch (e) { toast(errMsg(e), "err"); }
  }
  function renderLibrary() {
    const host = $("#eqLibrary");
    const groups = {};
    TM().LIBRARY.forEach((x, i) => (groups[x.group] = groups[x.group] || []).push([x, i]));
    host.innerHTML = Object.entries(groups).map(([g, xs]) => `<div class="libgroup"><div class="libtitle">${esc(g)}</div><div class="chips">${xs.map(([x, i]) => `<button type="button" class="chip" data-i="${i}" title="${esc(x.tex)}">${esc(x.name)}</button>`).join("")}</div></div>`).join("");
    host.querySelectorAll(".chip").forEach((b) => b.addEventListener("click", () => {
      $("#eqInput").value = TM().LIBRARY[+b.dataset.i].tex;
      mathState.grabbed = null;
      renderEqPreview();
      $("#eqInput").focus();
    }));
  }
  function initMath() {
    if (!TM()) return;
    const ta = $("#eqInput");
    let t;
    ta.addEventListener("input", () => { mathState.grabbed = null; clearTimeout(t); t = setTimeout(renderEqPreview, 150); });
    ta.addEventListener("paste", onEqPaste);
    $("#eqSymbols").innerHTML = SYMBOLS.map(([l, s], i) => `<button type="button" class="sym" data-i="${i}" title="${esc(s.trim())}">${esc(l)}</button>`).join("");
    $("#eqSymbols").querySelectorAll(".sym").forEach((b) => b.addEventListener("click", () => insertAtCaret(ta, SYMBOLS[+b.dataset.i][1])));
    document.querySelectorAll("#eqModeSeg button").forEach((b) => b.addEventListener("click", () => {
      mathState.mode = b.dataset.mode;
      document.querySelectorAll("#eqModeSeg button").forEach((x) => x.classList.toggle("on", x === b));
      $("#eqNumRow").classList.toggle("hidden", mathState.mode === "inline");
      renderEqPreview();
    }));
    $("#eqNumbering").addEventListener("change", renderEqPreview);
    $("#eqNumber").addEventListener("input", () => { $("#eqNumber").dataset.touched = $("#eqNumber").value ? "1" : ""; renderEqPreview(); });
    $("#eqInsert").addEventListener("click", eqInsert);
    $("#eqConvert").addEventListener("click", eqConvert);
    $("#eqGrab").addEventListener("click", eqGrab);
    $("#eqAutoNum").addEventListener("click", eqAutoNumber);
    const need = (f) => () => { const r = eqCurrent(); if (!r || r.error) { toast("Belum ada rumus yang valid.", "err"); return; } f(r); };
    $("#eqCopyTex").addEventListener("click", need((r) => copyText(TM().toLatex(r.ast), "LaTeX")));
    $("#eqCopyMml").addEventListener("click", need((r) => copyText(`<math xmlns="http://www.w3.org/1998/Math/MathML" display="block">${TM().toMathML(r.ast)}</math>`, "MathML")));
    $("#eqCopyUni").addEventListener("click", need((r) => copyText(TM().toUnicode(r.ast).trim(), "Teks rumus")));
    renderLibrary();
  }

  /* ---------------- start ---------------- */
  function boot(inWord) {
    state.inWord = inWord;
    loadStore();
    $("#profileLabel").textContent = state.store.active;
    $("#notInWord").classList.toggle("hidden", inWord);
    state.open[state.profile.sections[0] && state.profile.sections[0].key] = true;
    initTabs(); initMode(); initProfiles(); initMath(); initWhole();
    renderTemplate(); renderSections(); renderLimits(); renderCounts(null);
    $("#insertAll").addEventListener("click", () => doInsert(state.profile.sections.map((s) => s.key)));
    $("#undoLast").addEventListener("click", doUndo);
    $("#wipeAll").addEventListener("click", () => doWipe(state.profile.sections.filter((x) => x.kind !== "abstract").map((x) => x.key)));
    $("#recount").addEventListener("click", recount);
    $("#runCheck").addEventListener("click", runCheck);
    if (inWord) refreshScan().then(renderSections);
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    let booted = false;
    Office.onReady((info) => {
      const inWord = !!(info && info.host === Office.HostType.Word);
      if (!booted) { booted = true; boot(inWord); return; }
      // Word terlambat siap (mis. > 2,5 detik saat pertama dibuka): aktifkan mode Word
      if (inWord && !state.inWord) {
        state.inWord = true;
        $("#notInWord").classList.add("hidden");
        refreshScan().then(renderSections);
      }
    });
    // di luar Office, onReady kadang tidak terpanggil — tampilkan pratinjau, tetapi tetap tunggu Word
    setTimeout(() => { if (!booted) { booted = true; boot(false); } }, 2500);
  } else {
    document.addEventListener("DOMContentLoaded", () => boot(false));
  }

  // untuk pengujian
  window.__tplState = state;
})();
