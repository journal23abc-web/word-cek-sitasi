/*
 * TplCheck — pemeriksa kesesuaian naskah (tanpa Office.js).
 *   - "To cite this article" vs judul, penulis, jurnal, volume/nomor, tahun, halaman, DOI
 *   - Article history (Received / Revised / Accepted), default maks. 3 bulan received → accepted
 *   - Header & footer: placeholder, running head penulis/judul
 *   - Batas kata
 * Masukan `meta` = { infos:[{text,inTable,row,cell}], hf:[{kind,type,text}], startPage, today? }
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./engine.js"));
  else root.TplCheck = factory(root.TplEngine);
})(typeof self !== "undefined" ? self : this, function (E) {
  "use strict";

  /* ---------------- util ---------------- */
  const WS = /[\s ​ -   　]+/g;
  const clean = (s) => String(s || "").replace(WS, " ").trim();
  const norm = (s) => clean(s).toLowerCase().normalize("NFKD").replace(/[̀-ͯ]/g, "")
    .replace(/[‐-―−]/g, "-").replace(/[’‘`]/g, "'");
  const words = (s) => norm(s).replace(/[^\p{L}\p{N}]+/gu, " ").trim().split(" ").filter(Boolean);
  const fmtN = (n) => Number(n || 0).toLocaleString("id-ID");

  function lev(a, b) {
    a = norm(a); b = norm(b);
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  }

  /** Selisih kata (LCS): {missing: kata di A tidak ada di B, extra: kata di B tidak ada di A} */
  function wordDiff(a, b) {
    const A = words(a), B = words(b);
    const L = Array.from({ length: A.length + 1 }, () => new Array(B.length + 1).fill(0));
    for (let i = A.length - 1; i >= 0; i--) for (let j = B.length - 1; j >= 0; j--)
      L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const missing = [], extra = [];
    let i = 0, j = 0;
    while (i < A.length && j < B.length) {
      if (A[i] === B[j]) { i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) missing.push(A[i++]); else extra.push(B[j++]);
    }
    while (i < A.length) missing.push(A[i++]);
    while (j < B.length) extra.push(B[j++]);
    return { same: !missing.length && !extra.length, missing, extra };
  }

  /* ---------------- tanggal ---------------- */
  const MONTHS = {
    jan: 1, january: 1, januari: 1, feb: 2, february: 2, februari: 2, peb: 2, mar: 3, march: 3, maret: 3,
    apr: 4, april: 4, may: 5, mei: 5, jun: 6, june: 6, juni: 6, jul: 7, july: 7, juli: 7,
    aug: 8, august: 8, agu: 8, agt: 8, agustus: 8, sep: 9, sept: 9, september: 9,
    oct: 10, october: 10, okt: 10, oktober: 10, nov: 11, november: 11, nop: 11, dec: 12, december: 12, des: 12, desember: 12,
  };
  const MONTH_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const monthOf = (t) => { const k = t.toLowerCase().replace(/\.$/, ""); return MONTHS[k] || (k.length >= 3 ? MONTHS[k.slice(0, 3)] : 0) || 0; };
  const dim = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  function mk(y, m, d, raw, fmt) {
    if (!(m >= 1 && m <= 12) || !(d >= 1 && d <= dim(y, m))) return { invalid: true, raw };
    return { y, m, d, t: Date.UTC(y, m - 1, d), raw, fmt };
  }
  function parseDate(s) {
    const raw = clean(s);
    const t = raw.replace(/(\d)(st|nd|rd|th)\b/gi, "$1").replace(/\.$/, "");
    let m;
    if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t))) return mk(+m[1], +m[2], +m[3], raw, "ymd");
    if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(t))) {
      let d = +m[1], mo = +m[2];
      if (mo > 12 && d <= 12) [d, mo] = [mo, d];
      return mk(+m[3], mo, d, raw, "dmy-num");
    }
    let y = 0, mo = 0, d = 0;
    for (const tok of t.replace(/[,]/g, " ").split(/\s+/).filter(Boolean)) {
      if (/^\d{4}$/.test(tok)) y = +tok;
      else if (/^\d{1,2}$/.test(tok)) d = +tok;
      else if (monthOf(tok)) mo = monthOf(tok);
    }
    if (y && mo && d) return mk(y, mo, d, raw, "name");
    return null;
  }
  const fmtDate = (x) => (x && x.t != null ? `${x.d} ${MONTH_ID[x.m]} ${x.y}` : "—");
  const addMonths = (x, n) => { const t = new Date(Date.UTC(x.y, x.m - 1 + n, 1)); const y = t.getUTCFullYear(), m = t.getUTCMonth() + 1; return Date.UTC(y, m - 1, Math.min(x.d, dim(y, m))); };
  const DAY = 86400000;

  const HIST_KEYS = [
    ["received", /^(received|submitted|diterima|dikirim|tanggal masuk)$/i],
    ["revised", /^(revised|revision|revisi|direvisi|diperbaiki)$/i],
    ["accepted", /^(accepted|disetujui|diterima untuk diterbitkan)$/i],
    ["published", /^(published|available online|diterbitkan|terbit|online)$/i],
  ];
  function extractHistory(lines) {
    const out = {};
    for (const ln of lines) {
      for (const part0 of String(ln).split(/[;|\n]/)) {
        const part = clean(part0).replace(/^[(\[]+/, "").replace(/[)\]]+$/, ""); // "(Received: …; … Accepted: …)"
        const m = /^\s*([A-Za-z][A-Za-z ]{2,30}?)\s*(?:date|tanggal)?\s*[:：]\s*(.+?)\s*$/i.exec(part);
        if (!m) continue;
        const key = HIST_KEYS.find(([, re]) => re.test(m[1].trim()));
        if (key && !out[key[0]]) out[key[0]] = Object.assign({ text: m[2] }, { date: parseDate(m[2]) });
      }
    }
    return out;
  }

  /* ---------------- cite ---------------- */
  const RE_CITE = /(to cite this article|how to cite(?: this article)?|cite this article|please cite|cara (?:mengutip|sitasi)|sitasi artikel)\s*[:：]?\s*/i;
  const RE_INITIALS = /^(?:\p{Lu}[\p{Ll}]?\.?\s*-?\s*){1,6}\.*$/u;

  function parseAuthors(str) {
    const s = clean(str).replace(/\s*,?\s*&\s*/g, ", ").replace(/,?\s+and\s+(?=\p{Lu})/gu, ", ").replace(/,\s*$/, "");
    const tok = s.split(/\s*,\s*/).filter(Boolean);
    const out = [];
    for (let i = 0; i < tok.length; i++) {
      const t = tok[i];
      if (/^et al\.?$/i.test(t)) { out.push({ etal: true }); continue; }
      const n = tok[i + 1];
      if (n && RE_INITIALS.test(n)) { out.push({ surname: t, initials: n, raw: t + ", " + n }); i++; }
      else out.push({ surname: t, initials: "", raw: t, org: true });
    }
    return out;
  }

  function parseCite(text, journalName) {
    let t = clean(text);
    const m0 = RE_CITE.exec(t);
    if (m0) t = t.slice(m0.index + m0[0].length);
    const res = { raw: t };
    const y = /\(\s*(\d{4}[a-z]?|n\.\s?d\.|in press)\s*\)\s*\.?/i.exec(t);
    if (!y) return Object.assign(res, { error: "Tahun dalam kurung (2026) tidak ditemukan." });
    res.authorsStr = t.slice(0, y.index).trim();
    res.authors = parseAuthors(res.authorsStr);
    res.year = y[1];
    let rest = t.slice(y.index + y[0].length).trim();
    const doiM = /(https?:\/\/(?:dx\.)?doi\.org\/\S+|\bdoi:\s*\S+)/i.exec(rest);
    if (doiM) { res.doi = doiM[1].replace(/^doi:\s*/i, "https://doi.org/"); rest = rest.slice(0, doiM.index).trim(); }
    const vm = /,\s*(\d+|[Xx]+)\s*(?:\(\s*([\w\-–]+)\s*\))?\s*,\s*(?:pp?\.\s*)?([A-Za-z]?\d+|[Xx]+)\s*(?:[–\-—]+\s*([A-Za-z]?\d+|[Xx]+))?\s*\.?\s*$/.exec(rest);
    let before = rest;
    if (vm) {
      res.volume = vm[1]; res.issue = vm[2] || ""; res.pageStart = vm[3]; res.pageEnd = vm[4] || "";
      res.pageDash = (/(\d)\s*([–\-—]+)\s*\S/.exec(vm[0]) || [])[2] || "";
      before = rest.slice(0, vm.index);
    }
    before = before.trim();
    const jn = journalName ? norm(journalName) : "";
    const ib = norm(before).lastIndexOf(jn);
    if (jn && ib > 0) {
      // norm() tidak mengubah panjang huruf latin biasa — cari posisi dengan pencarian case-insensitive
      const pos = before.toLowerCase().lastIndexOf(journalName.toLowerCase().slice(0, 12));
      const cut = pos > 0 ? pos : ib;
      res.title = before.slice(0, cut).trim().replace(/[.\s]+$/, "");
      res.journal = before.slice(cut).trim().replace(/[,.\s]+$/, "");
    } else {
      const k = before.lastIndexOf(". ");
      res.title = (k > 0 ? before.slice(0, k) : before).trim();
      res.journal = (k > 0 ? before.slice(k + 2) : "").trim().replace(/[,.\s]+$/, "");
    }
    return res;
  }

  /* ---------------- halaman depan ---------------- */
  const RE_AFFIL = /(universit|institut|faculty|fakultas|department|departemen|jurusan|program studi|school|college|sekolah|politeknik|polytechnic|academy|akademi|hospital|rumah sakit|ministry|kementerian|centre|center|laborator|campus|kampus|indonesia|malaysia|india|germany|nigeria|philippines|thailand|vietnam|china|japan|australia|united|kingdom|usa\b|e-?mail|@|correspond|orcid)/i;
  const RE_SKIP_FRONT = /(article (history|info)|abstract|abstrak|keywords?|kata kunci|received|revised|accepted|open access|licen[cs]e|cite this|how to cite)/i;
  const PARTICLES = new Set(["bin", "binti", "van", "von", "de", "der", "den", "da", "di", "du", "del", "la", "le", "al", "el", "ibn", "binte", "a/l", "a/p"]);

  function frontMatter(infos, profile) {
    const heads = E.findSectionHeadings(infos, profile);
    const firstHead = heads.length ? heads[0].index : Math.min(infos.length, 60);
    const titleIdx = infos.findIndex((p, i) => i < firstHead && !p.inTable && clean(p.text).length > 3);
    return { firstHead, titleIdx, title: titleIdx >= 0 ? clean(infos[titleIdx].text) : "", lines: infos.slice(0, firstHead).map((p, i) => Object.assign({ i }, p)) };
  }

  function frontAuthors(fm) {
    const names = [];
    for (const p of fm.lines) {
      if (p.i === fm.titleIdx) continue;
      const t = clean(p.text);
      if (!t || t.length > 160 || E.countWords(t) > 14) continue;
      if (RE_AFFIL.test(t) || RE_SKIP_FRONT.test(t) || /[:;]\s*$|[.;]\s*$/.test(t) || /:/.test(t)) continue;
      const letters = t.replace(/[^\p{L}]/gu, "");
      if (letters.length >= 4 && letters === letters.toUpperCase()) continue;
      for (let part of t.split(/\s*(?:,|&|;|\band\b|\bdan\b)\s*/)) {
        part = part.replace(/[\d*†‡§¶#✉🖂📧✉️]+/gu, " ").replace(WS, " ").trim();
        const ws = part.split(" ").filter(Boolean);
        if (ws.length < 2 || ws.length > 6) continue;
        if (!ws.every((w) => /^\p{Lu}/u.test(w) || PARTICLES.has(w.toLowerCase()))) continue;
        if (!ws.every((w) => /^[\p{L}'’.\-]+$/u.test(w))) continue;
        names.push({ full: part, line: p.i });
      }
    }
    return names;
  }

  const initialsOf = (given) => given.split(/[\s\-]+/).filter((w) => w && !PARTICLES.has(w.toLowerCase())).map((w) => w[0].toUpperCase()).join("");
  const initialsLetters = (s) => String(s || "").replace(/[^\p{Lu}]/gu, "");

  /* ---------------- header / footer ---------------- */
  function issueInfo(hf) {
    const order = ["FirstPage", "Primary", "EvenPages"];
    const hs = hf.filter((h) => h.kind === "header").sort((a, b) => order.indexOf(a.type) - order.indexOf(b.type));
    for (const h of hs) {
      const m = /Vol(?:ume)?\.?\s*(\d+)\s*[,;]?\s*(?:No|Issue|Iss|Nomor|Number)\.?\s*(\d+)\s*[,;]?\s*\(?\s*(?:([A-Za-z]+)\.?\s+)?(\d{4})?\s*\)?/i.exec(clean(h.text));
      if (m) return { volume: m[1], issue: m[2], month: m[3] ? monthOf(m[3]) : 0, monthText: m[3] || "", year: m[4] || "", where: h.type };
    }
    return null;
  }
  function headerDoi(hf) {
    for (const h of hf) for (const ln of String(h.text).split("\n")) {
      if (RE_CITE.test(ln)) continue;
      const m = /https?:\/\/(?:dx\.)?doi\.org\/[^\s]+/i.exec(ln);
      if (m) return m[0].replace(/[.,;]+$/, "");
    }
    return "";
  }
  const HF_NAME = { header: "Header", footer: "Footer" };
  const TYPE_NAME = { FirstPage: "halaman pertama", Primary: "halaman ganjil", EvenPages: "halaman genap" };

  /* ---------------- pemeriksaan utama ---------------- */
  function run(meta, profile, opts) {
    opts = opts || {};
    meta = Object.assign({}, meta, { startPage: meta.startPage != null && meta.startPage !== "" ? parseInt(meta.startPage, 10) || null : null });
    const P = profile;
    const C = Object.assign({ maxMonths: 3, minDays: 0, requireRevised: true }, P.check || {});
    const today = opts.today ? new Date(opts.today) : new Date();
    const todayT = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
    const F = [];
    const seen = new Set();
    const add = (group, level, title, detail) => {
      const k = group + "|" + title + "|" + (detail || "");
      if (seen.has(k)) return;
      seen.add(k);
      F.push({ group, level, title, detail: detail || "" });
    };
    const infos = meta.infos || [], hf = meta.hf || [];
    const fm = frontMatter(infos, P);
    const issue = issueInfo(hf);
    const hdoi = headerDoi(hf);
    const jName = (P.journal && P.journal.name) || "";

    /* ===== Cite ===== */
    const G1 = "How to cite";
    let citeText = "", citeWhere = "";
    infos.some((p, i) => { if (RE_CITE.test(p.text || "")) { citeText = p.text; citeWhere = p.inTable ? "tabel halaman depan" : "badan dokumen"; return true; } return false; });
    if (!citeText) for (const h of hf) for (const ln of String(h.text).split("\n")) if (!citeText && RE_CITE.test(ln)) { citeText = ln; citeWhere = `${HF_NAME[h.kind]} ${TYPE_NAME[h.type] || ""}`.trim(); }
    let cite = null;
    if (!citeText) add(G1, C.requireCite === false ? "info" : "error", "Teks \"To cite this article\" tidak ditemukan", C.requireCite === false ? "Template jurnal ini tidak memakai cite di halaman depan." : "Dicari di badan dokumen, tabel, header, dan footer.");
    else {
      cite = parseCite(citeText, jName);
      add(G1, "info", "Ditemukan di " + citeWhere, clean(citeText).slice(0, 400));
      if (cite.error) add(G1, "error", "Format cite tidak terbaca", cite.error);
    }
    const names = frontAuthors(fm);
    if (cite && !cite.error) {
      // tanda baca
      const raw = cite.raw;
      const punct = [];
      if (/[^.]\.\.(?!\.)/.test(raw) || /\.\.\s*\(/.test(raw)) punct.push("titik ganda \"..\"" + ((/(\S{0,12}\.\.)/.exec(raw) || [])[1] ? ` (di "${(/(\S{0,12}\.\.)/.exec(raw) || [])[1]}")` : ""));
      if (/\s[,.;:]/.test(raw.replace(/\s(?=https?:)/g, ""))) punct.push("spasi sebelum tanda baca");
      if (/ {2,}/.test(String(citeText).replace(/^[\s\S]*?:/, "").replace(/\s+$/, ""))) punct.push("spasi ganda");
      if (/,(?=\S)/.test(raw.replace(/https?:\/\/\S+/g, ""))) punct.push("tidak ada spasi setelah koma");
      if (/[Xx]{3,}|\?\?/.test(raw)) punct.push("placeholder (xxx/XXX) belum diganti");
      if (punct.length) add(G1, "warn", "Tanda baca/penulisan cite perlu dicek", punct.join("; ") + ".");

      // penulis
      const cas = (cite.authors || []).filter((a) => !a.etal);
      if (cas.length >= 2 && !/&/.test(cite.authorsStr) && !(cite.authors || []).some((a) => a.etal)) add(G1, "warn", "Tidak ada \"&\" sebelum penulis terakhir (APA)", cite.authorsStr);
      if (!names.length) add(G1, "info", "Daftar penulis di halaman depan tidak terbaca otomatis", "Cek manual nama & inisial penulis di cite.");
      else {
        const used = new Set();
        const problems = [], okNames = [];
        cas.forEach((ca, k) => {
          const sn = norm(ca.surname);
          const idx = names.findIndex((n, j) => !used.has(j) && (norm(n.full) === sn || norm(n.full).endsWith(" " + sn)));
          if (idx < 0) {
            // cari yang mirip (salah ketik)
            let best = -1, bd = 99;
            names.forEach((n, j) => { if (used.has(j)) return; const last = n.full.split(" ").slice(-Math.max(1, ca.surname.split(" ").length)).join(" "); const d = lev(last, ca.surname); if (d < bd) { bd = d; best = j; } });
            if (best >= 0 && bd <= 2) { used.add(best); problems.push(`"${ca.surname}" berbeda ejaan dengan nama di artikel "${names[best].full}"`); }
            else problems.push(`"${ca.raw}" tidak ada di daftar penulis artikel`);
            return;
          }
          used.add(idx);
          if (idx !== k) problems.push(`urutan "${ca.surname}" berbeda (di artikel urutan ke-${idx + 1}, di cite ke-${k + 1})`);
          const full = names[idx].full;
          const given = full.slice(0, full.length - ca.surname.length).trim();
          const exp = initialsOf(given), got = initialsLetters(ca.initials);
          if (!ca.org && exp && got && exp !== got) problems.push(`inisial ${ca.surname}: cite "${ca.initials}", seharusnya "${exp.split("").join(". ")}." (dari "${full}")`);
          else okNames.push(ca.surname);
        });
        names.forEach((n, j) => { if (!used.has(j) && !(cite.authors || []).some((a) => a.etal)) problems.push(`penulis "${n.full}" belum ada di cite`); });
        if (problems.length) add(G1, "error", "Penulis di cite tidak sesuai artikel", problems.join("; ") + ".");
        else add(G1, "ok", `Penulis sesuai (${cas.length} orang, urutan & inisial cocok)`, cas.map((a) => a.raw).join("; "));
      }

      // judul
      if (fm.title && cite.title) {
        const d = wordDiff(fm.title, cite.title);
        if (d.same) add(G1, "ok", "Judul di cite sama dengan judul artikel");
        else add(G1, "error", "Judul di cite berbeda dari judul artikel",
          (d.missing.length ? `Kata di artikel yang tidak ada di cite: ${d.missing.slice(0, 8).join(", ")}. ` : "") +
          (d.extra.length ? `Kata di cite yang tidak ada di artikel: ${d.extra.slice(0, 8).join(", ")}.` : ""));
      }
      // jurnal
      if (jName) {
        if (cite.journal && norm(cite.journal).replace(/[^a-z0-9]+/g, "") === norm(jName).replace(/[^a-z0-9]+/g, "")) add(G1, "ok", "Nama jurnal sesuai", cite.journal);
        else add(G1, "error", "Nama jurnal di cite tidak sesuai", `Cite: "${cite.journal || "—"}" · seharusnya: "${jName}"`);
      }
      // volume / nomor
      if (issue) {
        const okV = String(cite.volume) === issue.volume, okI = String(cite.issue) === issue.issue;
        if (okV && okI) add(G1, "ok", `Volume & nomor sesuai: ${cite.volume}(${cite.issue})`, `Header: Vol. ${issue.volume}, No. ${issue.issue}`);
        else add(G1, "error", "Volume/nomor di cite tidak sesuai header", `Cite: ${cite.volume || "—"}(${cite.issue || "—"}) · Header: Vol. ${issue.volume}, No. ${issue.issue}`);
        if (issue.year && cite.year && cite.year.slice(0, 4) !== issue.year) add(G1, "error", "Tahun di cite berbeda dengan tahun terbitan", `Cite: ${cite.year} · Header: ${issue.year}`);
        else if (issue.year) add(G1, "ok", "Tahun sesuai: " + cite.year);
      } else if (cite.volume) add(G1, "info", "Volume/nomor di header tidak terbaca", `Cite: ${cite.volume}(${cite.issue}). Cek manual.`);
      // halaman
      if (!cite.pageStart) add(G1, "warn", "Nomor halaman tidak terbaca di cite");
      else if (/x/i.test(cite.pageStart + cite.pageEnd)) add(G1, "warn", "Nomor halaman masih placeholder", `${cite.pageStart}–${cite.pageEnd}`);
      else {
        const s = parseInt(cite.pageStart, 10), e = parseInt(cite.pageEnd || cite.pageStart, 10);
        const msgs = [];
        if (e < s) msgs.push(`halaman akhir (${e}) lebih kecil dari awal (${s})`);
        if (meta.startPage && s !== meta.startPage) msgs.push(`halaman awal ${s}, tetapi penomoran halaman dokumen dimulai dari ${meta.startPage}`);
        if (cite.pageDash && cite.pageDash !== "–") msgs.push(`pemisah halaman "${cite.pageDash}" sebaiknya en dash "–"`);
        if (msgs.length) add(G1, "warn", "Halaman perlu dicek", msgs.join("; ") + ".");
        else add(G1, "ok", `Halaman ${s}–${e}` + (meta.startPage ? " (halaman awal cocok dengan penomoran dokumen)" : ""), meta.startPage ? "" : "Halaman awal dokumen tidak terbaca; cek manual.");
      }
      // DOI
      if (!cite.doi) add(G1, "warn", "DOI tidak ada di cite", hdoi ? "DOI di header: " + hdoi : "");
      else {
        const doi = cite.doi.replace(/[.,;]+$/, "");
        const msgs = [];
        if (/\.$/.test(cite.doi) && /\.(v\d+i\d+)\.$/i.test(cite.doi)) msgs.push(`DOI tampak belum lengkap (berakhir "${cite.doi.slice(-8)}", nomor artikel belum ada)`);
        else if (!/^https:\/\/doi\.org\/10\.\d{4,9}\/\S+$/i.test(doi)) msgs.push("format DOI tidak standar (harus https://doi.org/10.xxxx/…)");
        const vi = /v(\d+)i(\d+)/i.exec(doi);
        if (vi && cite.volume && (vi[1] !== String(cite.volume) || vi[2] !== String(cite.issue))) msgs.push(`DOI memuat v${vi[1]}i${vi[2]}, tetapi cite ${cite.volume}(${cite.issue})`);
        if (hdoi && norm(hdoi) !== norm(doi)) msgs.push(`berbeda dengan DOI di header (${hdoi})`);
        if (msgs.length) add(G1, "error", "DOI perlu diperbaiki", msgs.join("; ") + ".");
        else add(G1, "ok", "DOI valid" + (hdoi ? " dan sama dengan header" : ""), doi);
      }
    }

    /* ===== Article history ===== */
    const G2 = "Article history";
    const H = extractHistory(fm.lines.map((p) => p.text).concat(infos.slice(fm.firstHead).filter((p) => p.inTable).slice(0, 0).map((p) => p.text)));
    const R = H.received && H.received.date, V = H.revised && H.revised.date, A = H.accepted && H.accepted.date;
    const label = { received: "Received", revised: "Revised", accepted: "Accepted" };
    ["received", "revised", "accepted"].forEach((k) => {
      if (!H[k]) { if (k !== "revised" || C.requireRevised) add(G2, k === "revised" ? "warn" : "error", `${label[k]} tidak ditemukan`); }
      else if (!H[k].date) add(G2, "error", `Tanggal ${label[k]} tidak terbaca`, `"${H[k].text}"`);
      else if (H[k].date.invalid) add(G2, "error", `Tanggal ${label[k]} tidak valid`, `"${H[k].text}"`);
    });
    const good = (x) => x && !x.invalid && x.t != null;
    if (good(R) && good(A)) {
      const days = Math.round((A.t - R.t) / DAY);
      const limitT = addMonths(R, C.maxMonths);
      const detail = `${fmtDate(R)} → ${fmtDate(A)} = ${days} hari (±${(days / 30.44).toFixed(1)} bulan)`;
      if (A.t < R.t) add(G2, "error", "Accepted lebih awal dari Received", detail);
      else if (C.maxMonths && A.t > limitT) add(G2, "warn", `Jarak Received → Accepted melebihi ${C.maxMonths} bulan`, `${detail}. Batas: ${fmtDate({ y: new Date(limitT).getUTCFullYear(), m: new Date(limitT).getUTCMonth() + 1, d: new Date(limitT).getUTCDate(), t: limitT })}; lebih ${Math.round((A.t - limitT) / DAY)} hari.`);
      else if (C.minDays && days < C.minDays) add(G2, "warn", `Jarak Received → Accepted kurang dari ${C.minDays} hari`, detail);
      else add(G2, "ok", `Jarak Received → Accepted wajar (≤ ${C.maxMonths} bulan)`, detail);
    }
    if (good(V)) {
      if (good(R) && V.t < R.t) add(G2, "error", "Revised lebih awal dari Received", `${fmtDate(R)} vs ${fmtDate(V)}`);
      if (good(A) && V.t > A.t) add(G2, "error", "Revised setelah Accepted", `${fmtDate(V)} vs ${fmtDate(A)}`);
    }
    if (good(A) && issue && issue.year) {
      const iy = +issue.year, im = issue.month || 12;
      const endT = Date.UTC(iy, im, 0);
      if (A.t > endT) add(G2, "warn", "Accepted setelah bulan terbit", `${fmtDate(A)}, terbitan ${issue.monthText ? issue.monthText + " " : ""}${iy}`);
    }
    [["received", R], ["revised", V], ["accepted", A]].forEach(([k, x]) => { if (good(x) && x.t > todayT) add(G2, "info", `Tanggal ${label[k]} di masa depan`, fmtDate(x)); });
    const fmts = [R, V, A].filter(good).map((x) => x.fmt);
    if (new Set(fmts).size > 1) add(G2, "info", "Format penulisan tanggal tidak seragam", ["received", "revised", "accepted"].filter((k) => H[k]).map((k) => `${label[k]}: ${H[k].text}`).join(" · "));
    if (cite && cite.year && good(A) && String(A.y) !== cite.year.slice(0, 4) && !(issue && issue.year === cite.year.slice(0, 4))) add(G2, "info", "Tahun accepted berbeda dengan tahun di cite", `${A.y} vs ${cite.year}`);

    /* ===== Header & footer ===== */
    const G3 = "Header & footer";
    const firstAuthor = cite && cite.authors && cite.authors[0] && cite.authors[0].surname ? cite.authors[0].surname : (names[0] ? names[0].full.split(" ").slice(-1)[0] : "");
    let anyHf = false;
    for (const h of hf) {
      for (const ln of String(h.text).split("\n")) {
        const t = clean(ln);
        if (!t || RE_CITE.test(t)) continue;
        anyHf = true;
        const where = `${HF_NAME[h.kind]} ${TYPE_NAME[h.type] || h.type}`;
        const ph = [];
        if (/(^|[\s|│])(Author|Authors|Penulis|Nama Penulis)(\s*(et al\.?)?)(?=\s*[│|\t]|\s*$)/i.test(t)) ph.push("\"Author\"");
        if (/(^|[\s|│])(Title|Judul|Article Title)\s*(…|\.{2,})?\.*(?=\s*$|\s*[│|])/i.test(t)) ph.push("\"Title…\"");
        if (/X{4}|Y{4}|\bVol\.?\s*X\b|\bNo\.?\s*X\b/i.test(t)) ph.push((t.match(/[A-Z\-]*ISSN[^;|]*X{4}[^;|]*/i) || t.match(/\S*X{4}\S*/i) || [""])[0] || "XXXX");
        if (ph.length) add(G3, "warn", `Placeholder belum diganti di ${where}`, `${ph.join(", ")} — "${t.slice(0, 120)}"`);
        // running head penulis
        const rh = /^(?:.*?\b)?([\p{Lu}][\p{L}'’\-]+(?:\s[\p{Lu}][\p{L}'’\-]+)?)(?:,\s*(?:\p{Lu}\.\s*)+,?)?\s+et\s+al\.?/u.exec(t) || /^([\p{Lu}][\p{L}'’\-]+)\s*(?:&|and|dan)\s*[\p{Lu}][\p{L}'’\-]+\s*[│|]/u.exec(t);
        if (rh && firstAuthor) {
          const sn = rh[1].split(" ").slice(-firstAuthor.split(" ").length).join(" ");
          if (norm(sn) !== norm(firstAuthor) && !norm(rh[1]).endsWith(norm(firstAuthor))) add(G3, "error", `Running head penulis salah di ${where}`, `Tertulis "${rh[0].trim()}", seharusnya "${firstAuthor} et al." (penulis pertama: ${firstAuthor})`);
          else add(G3, "ok", `Running head penulis sesuai (${where})`, rh[0].trim());
        }
        // running title
        const rt = /[│|]\s*(.+?)\s*(?:…|\.{2,})\.?\s*$/.exec(t);
        if (rt && fm.title && !/^(title|judul)$/i.test(rt[1].trim())) {
          const short = words(rt[1]), full = words(fm.title);
          const ok = short.length && short.every((w, k) => full[k] === w);
          if (!ok) add(G3, "error", `Running title tidak sama dengan awal judul (${where})`, `"${rt[1]}" vs "${fm.title.slice(0, 80)}"`);
          else add(G3, "ok", `Running title sesuai (${where})`, rt[1]);
        }
        // nama jurnal di header
        if (jName && h.kind === "header" && /journal|sciences|jurnal/i.test(t) && !/ISSN|https?:/i.test(t) && t.length < 120) {
          const m = norm(t).includes(norm(jName));
          if (!m && lev(t.replace(/\s*Vol[\s\S]*$/i, ""), jName) > 3) add(G3, "warn", `Nama jurnal di ${where} berbeda`, `"${t}" vs "${jName}"`);
        }
      }
    }
    if (!anyHf) add(G3, "info", "Header/footer tidak terbaca");
    // rentang halaman di header ("Page 57-71") vs penomoran halaman dokumen
    for (const h of hf) {
      const m = /\b(?:Page|Pages|pp?\.|Hal(?:aman)?\.?)\s*(\d+)\s*[-–—]\s*(\d+)/i.exec(clean(h.text));
      if (!m) continue;
      const a = +m[1], b = +m[2];
      if (b < a) add(G3, "warn", "Rentang halaman di header terbalik", m[0]);
      if (meta.startPage && a !== meta.startPage) add(G3, "warn", `Header menulis halaman ${a}–${b}, tetapi penomoran dokumen dimulai dari ${meta.startPage}`, "Atur Page Number Format → Start at: " + a + ", atau perbaiki rentang di header.");
      if (cite && cite.pageStart && /^\d+$/.test(cite.pageStart) && (+cite.pageStart !== a || (cite.pageEnd && +cite.pageEnd !== b))) add(G3, "error", "Rentang halaman di header berbeda dengan cite", `Header: ${a}–${b} · Cite: ${cite.pageStart}–${cite.pageEnd}`);
      break;
    }
    if (issue) add(G3, "info", `Terbitan: Vol. ${issue.volume}, No. ${issue.issue}${issue.monthText ? ", " + issue.monthText : ""}${issue.year ? " " + issue.year : ""}`, hdoi ? "DOI header: " + hdoi : "");

    /* ===== Jumlah kata ===== */
    const G4 = "Jumlah kata";
    const cnt = E.countDocument(infos, E.findSectionHeadings(infos, P), P);
    if (cnt.limit) {
      if (cnt.total > cnt.limit) add(G4, "error", `Isi ${fmtN(cnt.total)} kata — melebihi batas ${fmtN(cnt.limit)}`, `Lebih ${fmtN(cnt.total - cnt.limit)} kata.`);
      else add(G4, "ok", `Isi ${fmtN(cnt.total)} / ${fmtN(cnt.limit)} kata`);
    } else add(G4, "info", `Isi ${fmtN(cnt.total)} kata`, "Batas kata jurnal belum diatur (tab Hitung kata).");
    if (cnt.minLimit && cnt.total < cnt.minLimit) add(G4, "warn", `Isi kurang dari minimal ${fmtN(cnt.minLimit)} kata`, `Kurang ${fmtN(cnt.minLimit - cnt.total)} kata.`);
    if (cnt.abstractLimit) {
      if (cnt.abstract > cnt.abstractLimit) add(G4, "error", `Abstrak ${fmtN(cnt.abstract)} kata — melebihi ${fmtN(cnt.abstractLimit)}`);
      else add(G4, "ok", `Abstrak ${fmtN(cnt.abstract)} / ${fmtN(cnt.abstractLimit)} kata`);
    } else if (cnt.abstract) add(G4, "info", `Abstrak ${fmtN(cnt.abstract)} kata`, "Batas kata abstrak belum diatur.");

    return { findings: F, parsed: { title: fm.title, authors: names.map((n) => n.full), cite, history: H, issue, headerDoi: hdoi, startPage: meta.startPage || null, count: cnt } };
  }

  return { run, parseDate, parseCite, parseAuthors, extractHistory, frontAuthors, frontMatter, issueInfo, wordDiff, fmtDate };
});
