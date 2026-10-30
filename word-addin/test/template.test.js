// Tes Node: Cek Template (hitung kata ala Word, How to cite, kata kunci) — memakai dua naskah
// contoh di test/fixtures bila ada, plus naskah sintetis.
//   node word-addin/test/template.test.js
var assert = require('assert');
var fs = require('fs');
var path = require('path');
var JSZip = require('../vendor/jszip.min.js');
var DR = require('../docx-reader.js');
var TC = require('../template-core.js');

var passed = 0, failed = 0, queue = [];
function test(name, fn) { queue.push([name, fn]); }

function clone(m) { return JSON.parse(JSON.stringify(m)); }
// Ganti teks pada paragraf model (text & countText) — mensimulasikan suntingan penulis.
function edit(model, story, find, repl) {
  var list = story === 'body' ? model.body : model.headerFooter[story];
  var p = list.find(function (x) { return x.text.indexOf(find) >= 0; });
  assert.ok(p, 'teks untuk diedit tidak ditemukan: ' + find);
  p.text = p.text.split(find).join(repl); p.countText = p.countText.split(find).join(repl);
  return model;
}
function check(r, id) { return r.checks.find(function (c) { return c.id === id; }); }

// ---------- hitung kata ----------
test('aturan Word: en/em dash memisah, dash berdiri sendiri tidak dihitung', function () {
  assert.strictEqual(TC.countWords('343–44'), 2);
  assert.strictEqual(TC.countWords('up–to–date'), 3);
  assert.strictEqual(TC.countWords('up – to – date'), 3);
  assert.strictEqual(TC.countWords('Purpose — Generative AI'), 3);
  assert.strictEqual(TC.countWords('up - to - date'), 5, 'tanda hubung berspasi dihitung');
  assert.strictEqual(TC.countWords('343-44 self-efficacy Originality/Value'), 3);
  assert.strictEqual(TC.countWords('The court … concluded'), 4);
  assert.strictEqual(TC.countWords('The court . . . concluded'), 6);
  assert.strictEqual(TC.countWords('cats & dogs'), 3);
  assert.strictEqual(TC.countWords('​From policy'), 2, 'zero-width space di awal tidak menambah kata');
  assert.strictEqual(TC.countWords('a\tb\u000bc  \n d'), 4);
});
test('kode field & teks terhapus tidak dihitung; hasil field dihitung', function () {
  var xml = '<w:document><w:body><w:p><w:r><w:t>Satu dua</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ADDIN ZOTERO_ITEM {"a":"b c d"} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve"> (Smith, 2020)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:del w:id="1"><w:r><w:delText>hapus ini</w:delText></w:r></w:del><w:r><w:rPr><w:vanish/></w:rPr><w:t> tersembunyi</w:t></w:r></w:p></w:body></w:document>';
  var m = DR.read({ document: xml });
  assert.strictEqual(m.body[0].text, 'Satu dua (Smith, 2020) tersembunyi');
  assert.strictEqual(TC.countWords(m.body[0].countText), 4);
});
test('kotak teks & catatan kaki ikut total, header/footer tidak', function () {
  var doc = '<w:document><w:body><w:p><w:r><w:t>Badan naskah tiga</w:t></w:r><w:r><mc:AlternateContent><mc:Choice><w:drawing><w:txbxContent><w:p><w:r><w:t>kotak teks</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>kotak teks</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>';
  var fn = '<w:footnotes><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Catatan kaki empat kata.</w:t></w:r></w:p></w:footnote></w:footnotes>';
  var m = DR.read({ document: doc, footnotes: fn });
  assert.strictEqual(m.body.length, 1);
  assert.strictEqual(m.textboxes.length, 1, 'cadangan mc:Fallback tidak dihitung dua kali');
  var r = TC.analyze(m, {});
  var total = r.counts.find(function (c) { return c.key === 'total'; });
  assert.deepStrictEqual([total.main, total.textboxes, total.footnotes, total.words], [3, 2, 4, 9]);
  r = TC.analyze(m, { rules: { totalIncludesNotes: false } });
  assert.strictEqual(r.counts.find(function (c) { return c.key === 'total'; }).words, 3);
});

// ---------- naskah sintetis (tanpa tabel) ----------
function synth(lines) {
  var body = lines.map(function (t) {
    var caps = /^[A-Z ]+$/.test(t) && t.length > 3;
    return '<w:p>' + (caps ? '<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>' : '') + '<w:r>' + (caps ? '<w:rPr><w:b/></w:rPr>' : '') + '<w:t xml:space="preserve">' + t.replace(/&/g, '&amp;') + '</w:t></w:r></w:p>';
  }).join('');
  return DR.read({ document: '<w:document><w:body>' + body + '<w:sectPr><w:pgNumType w:start="5"/></w:sectPr></w:body></w:document>', app: '<Properties><Pages>3</Pages></Properties>' });
}
var SYN = [
  'A study of things in places',
  'Ani Budi, Citra Dewi',
  'Universitas Contoh, Indonesia',
  'Abstract: Kalimat abstrak satu dua tiga empat lima enam tujuh.',
  'Keywords: zebra; apel; mangga',
  'How to cite: Budi, A., & Dewi, C. (2025). A study of things. Jurnal Contoh, 3(1), 5-8.',
  '1. Introduction',
  'Isi pendahuluan dengan enam kata.',
  '2. Method',
  'Metode tiga kata.',
  '3. Results',
  'Hasil empat kata saja.',
  '3.1 Subbagian',
  'Masih hasil.',
  '4. Discussion',
  'Pembahasan dua.',
  '5. Conclusions',
  'Simpulan.',
  'References',
  'Budi, A. (2020). Buku. Penerbit.'
];
test('naskah bernomor: Results + Discussion dijumlahkan, subjudul 3.1 tidak memotong bagian', function () {
  var r = TC.analyze(synth(SYN), {});
  var rd = r.counts.find(function (c) { return c.key === 'results_discussion'; });
  // Results: "Hasil empat kata saja." (4) + "3.1 Subbagian" (2) + "Masih hasil." (2) = 8; Discussion 2
  assert.strictEqual(rd.words, 10);
  assert.ok(/Results 8 \+ Discussion 2/.test(rd.combined));
  assert.strictEqual(r.counts.find(function (c) { return c.key === 'introduction'; }).words, 5);
  assert.strictEqual(r.counts.find(function (c) { return c.key === 'method'; }).words, 3);
  var tr = r.counts.find(function (c) { return c.key === 'title'; });
  assert.deepStrictEqual([tr.words, tr.limit, tr.state], [6, '≤ 20', 'ok']);
  var r2 = TC.analyze(synth(SYN), { rules: { title: { min: 8, max: 15 } } });
  assert.strictEqual(check(r2, 'words:title').status, 'bad');
  assert.ok(/Kurang 2 kata/.test(check(r2, 'words:title').detail));
  var tc = TC.reportComments(r2, {}).find(function (c) { return c.id === 'words:title'; });
  assert.deepStrictEqual(tc.loc, { story: 'body', index: 0, find: null });
  assert.ok(/Jumlah kata Judul: 6 kata\. Batas 8–15 — kurang 2 kata\./.test(tc.text), tc.text);
  assert.strictEqual(r.counts.find(function (c) { return c.key === 'conclusion'; }).words, 1);
  assert.strictEqual(r.counts.find(function (c) { return c.key === 'abstract'; }).words, 9);
});
test('kata kunci tidak alfabetis → urutan benar + perbaikan', function () {
  var r = TC.analyze(synth(SYN), {});
  var kw = check(r, 'kw');
  assert.strictEqual(kw.status, 'bad');
  assert.deepStrictEqual(kw.sorted, ['apel', 'mangga', 'zebra']);
  assert.ok(kw.fix && kw.fix.items.length === 3 && kw.fix.items[0].offset === 'Keywords: '.length);
  assert.ok(check(r, 'kw:count') === undefined, '3 kata kunci masih dalam batas 3–6');
});
test('rentang halaman: start 5 + 3 halaman = 5–7, tertulis 5-8 → salah & tanda hubung', function () {
  var r = TC.analyze(synth(SYN), {});
  var pg = check(r, 'cite:pages');
  assert.strictEqual(pg.status, 'bad');
  assert.ok(/berakhir di halaman 7/.test(pg.detail) && /en dash/.test(pg.detail));
  assert.deepStrictEqual([pg.fix.find, pg.fix.replace], ['5-8', '5–7']);
  var t = check(r, 'cite:title');
  assert.strictEqual(t.status, 'bad');
  assert.ok(t.diff.some(function (d) { return d.op === 'ins' && /in places/.test(d.text); }));
});

// ---------- dua naskah contoh ----------
var FIX = path.join(__dirname, 'fixtures');
function sample(name) {
  var f = path.join(FIX, name);
  if (!fs.existsSync(f)) return null;
  return JSZip.loadAsync(fs.readFileSync(f)).then(DR.partsFromZip).then(DR.read);
}
test('contoh NEP (tabel halaman 1, How to cite di tabel): semua sesuai', function () {
  return sample('nep-74-96.docx').then(function (m) {
    if (!m) { console.log('        (lewati: fixture tidak ada)'); return; }
    var r = TC.analyze(m, {});
    assert.strictEqual(m.section.pgStart, 74);
    assert.strictEqual(r.pages.count, 23); assert.strictEqual(r.pages.end, 96);
    assert.strictEqual(r.howToCite.story, 'body');
    ['cite:pages', 'cite:title', 'cite:meta', 'kw'].forEach(function (id) { assert.strictEqual(check(r, id).status, 'ok', id + ': ' + check(r, id).detail); });
    assert.deepStrictEqual(r.keywords.items.map(function (k) { return k.text; }), ['Girl student retention', 'NEP 2020', 'Policy implementation', 'Rural government schools']);
    assert.ok(!r.abstract.parts.some(function (p) { return /open access/i.test(p.text); }), 'lisensi CC BY-SA tidak ikut abstrak');
    assert.strictEqual(r.abstract.words, 226);
    assert.deepStrictEqual(r.sections.slice(0, 5).map(function (s) { return s.key; }), ['introduction', 'literature', 'method', 'results_discussion', 'conclusion']);
    assert.deepStrictEqual(r.counts.map(function (c) { return c.key; }), ['title', 'abstract', 'introduction', 'method', 'results_discussion', 'conclusion', 'total']);
    assert.strictEqual(r.counts[0].words, TC.countWords('From policy to practice: NEP 2020, state-level interventions, and girl student retention in rural government schools of West Bengal'));
    assert.strictEqual(check(r, 'words:title').status, 'ok');
    assert.strictEqual(r.counts.find(function (c) { return c.key === 'method'; }).words, 1782);
    assert.strictEqual(r.assets.tables, 5);
  });
});
test('contoh NEP disunting: halaman akhir, judul, penulis, nomor, kata kunci', function () {
  return sample('nep-74-96.docx').then(function (m) {
    if (!m) return;
    var cite = 'To cite this article';
    var h = m.body.find(function (p) { return p.text.indexOf(cite) >= 0; }).text;
    edit(m, 'body', h, h.replace('74–96', '74–95').replace('Roy, B.', 'Ray, B.').replace('2(2)', '2(3)').replace('girl student retention in rural', 'girl retention in rural'));
    edit(m, 'body', 'Girl student retention; ', 'Teacher support; ');
    var r = TC.analyze(m, {});
    var pg = check(r, 'cite:pages');
    assert.strictEqual(pg.status, 'bad'); assert.deepStrictEqual([pg.fix.find, pg.fix.replace], ['74–95', '74–96']);
    var t = check(r, 'cite:title');
    assert.strictEqual(t.status, 'bad'); assert.ok(t.diff.some(function (d) { return d.op === 'ins' && d.text === 'student'; }), JSON.stringify(t.diff));
    var meta = check(r, 'cite:meta');
    assert.strictEqual(meta.status, 'bad');
    assert.ok(/Ray tidak ada di byline/.test(meta.detail) && /DOI memuat v2i2/.test(meta.detail), meta.detail);
    assert.strictEqual(check(r, 'kw').status, 'bad');
  });
});
test('contoh VOICE (How to cite di footer halaman 1): Title Case, tanda hubung, "Doi."', function () {
  return sample('voice-211-226.docx').then(function (m) {
    if (!m) { console.log('        (lewati: fixture tidak ada)'); return; }
    var r = TC.analyze(m, {});
    assert.strictEqual(r.howToCite.story, 'footer:first');
    assert.strictEqual(r.pages.start + '-' + r.pages.end, '211-226');
    var pg = check(r, 'cite:pages');
    assert.strictEqual(pg.status, 'warn'); assert.deepStrictEqual([pg.fix.find, pg.fix.replace], ['211-226', '211–226']);
    var t = check(r, 'cite:title');
    assert.strictEqual(t.status, 'warn'); assert.ok(/Title Case/.test(t.detail));
    assert.ok(/^Generative AI-assisted instruction and college students’/.test(t.fix.replace));
    assert.ok(/Doi\./.test(check(r, 'cite:meta').detail));
    assert.strictEqual(check(r, 'cite:meta').checked.authors, true);
    assert.strictEqual(check(r, 'kw').status, 'ok');
    assert.strictEqual(r.keywords.items.length, 5);
    assert.strictEqual(r.abstract.parts.length, 5);
    // halaman dari tata letak Word mengalahkan statistik file
    var r2 = TC.analyze(m, { pages: { count: 17, source: 'layout' } });
    assert.ok(/227/.test(check(r2, 'cite:pages').detail));
    assert.ok(check(r2, 'cite:pageof') === undefined || /227/.test(check(r2, 'cite:pageof').detail));
  });
});

test('laporan komentar: lokasi tepat & teks; How to cite di footer → judul naskah', function () {
  var RC = require('../report-core.js');
  assert.ok(RC.isTemplateComment(RC.TEMPLATE_PREFIX + ' x') && RC.commentMatches('[Cek Sitasi] a', 'validator') && !RC.commentMatches('[Cek Sitasi] a', 'template'));
  assert.ok(RC.commentMatches(RC.TEMPLATE_PREFIX + ' x', 'all') && !RC.commentMatches(RC.TEMPLATE_PREFIX + ' x', 'validator') && !RC.commentMatches('komentar reviewer', 'all'));
  var m = synth(SYN), r = TC.analyze(m, {});
  var bad = TC.reportComments(r, { includeOk: false });
  assert.ok(bad.every(function (c) { return c.status === 'bad' || c.status === 'warn'; }));
  var pages = bad.find(function (c) { return c.id === 'cite:pages'; });
  assert.deepStrictEqual(pages.loc, { story: 'body', index: 5, find: '5-8' });
  assert.ok(/Ubah menjadi 5–7/.test(pages.text));
  var title = bad.find(function (c) { return c.id === 'cite:title'; });
  assert.strictEqual(title.loc.find, 'A study of things');
  var all = TC.reportComments(r, { includeOk: true });
  var method = all.find(function (c) { return c.id === 'words:method'; });
  assert.deepStrictEqual([method.loc.index, method.text], [8, 'Jumlah kata Method: 3 kata.']);
  assert.ok(all.length > bad.length);
});
test('laporan komentar contoh VOICE: judul & halaman ditempel di judul naskah', function () {
  return sample('voice-211-226.docx').then(function (m) {
    if (!m) return;
    var r = TC.analyze(m, {}), cs = TC.reportComments(r, { includeOk: true });
    var t = cs.find(function (c) { return c.id === 'cite:title'; });
    assert.deepStrictEqual(t.loc, { story: 'body', index: r.title.index });
    assert.ok(/footer halaman pertama/.test(t.text) && /Judul naskah: "Generative AI-assisted instruction/.test(t.text), t.text);
    assert.ok(/Ubah menjadi 211–226/.test(cs.find(function (c) { return c.id === 'cite:pages'; }).text));
    var meth = cs.find(function (c) { return c.id === 'words:method'; });
    assert.ok(/^Jumlah kata Method: 1\.510 kata \(termasuk 75 kata di tabel\)\.$/.test(meth.text), meth.text);
    assert.strictEqual(m.body[meth.loc.index].text.trim(), 'METHOD');
  });
});

// ---------- halaman judul, DOI terpotong, isian templat ----------
function frontDoc(opts) {
  var o = opts || {};
  var lines = ['A study of things in places', o.byline || 'Ani Budi1*, Citra Dewi2, Eko Fajar3'].concat(o.affils || ['1Universitas Contoh, Indonesia', '2Universitas Lain, Indonesia', '3Institut Ketiga, Indonesia'])
    .concat(['Abstract: Kalimat abstrak satu dua tiga empat lima enam tujuh.', 'Keywords: apel; mangga; zebra', 'How to cite: ' + (o.cite || 'Budi, A., Dewi, C., & Fajar, E. (2025). A study of things in places. Jurnal Contoh, 3(1), 5–7. https://doi.org/10.1234/jc.v3i1.42'),
      '1. Introduction', 'Isi.', 'References', 'Budi, A. (2020). Buku. Penerbit.']);
  var body = lines.map(function (t) { return '<w:p><w:r><w:t xml:space="preserve">' + t.replace(/&/g, '&amp;') + '</w:t></w:r></w:p>'; }).join('');
  var hf = { 'word/header1.xml': '<w:hdr><w:p><w:r><w:t>' + (o.header || 'Jurnal Contoh Vol. 3 No. 1 (2025)') + '</w:t></w:r></w:p></w:hdr>' };
  if (o.footer) hf['word/footer1.xml'] = '<w:ftr><w:p>' + o.footer + '</w:p></w:ftr>';
  return DR.read({ document: '<w:document><w:body>' + body + '<w:sectPr><w:headerReference w:type="default" r:id="h1"/>' + (o.footer ? '<w:footerReference w:type="default" r:id="f1"/>' : '') + '<w:pgNumType w:start="5"/></w:sectPr></w:body></w:document>',
    rels: '<Relationships><Relationship Id="h1" Target="header1.xml"/><Relationship Id="f1" Target="footer1.xml"/></Relationships>', hf: hf, app: '<Properties><Pages>3</Pages></Properties>' });
}
test('nomor afiliasi: dobel berisi beda & nomor hilang terdeteksi; afiliasi sama yang diulang tidak', function () {
  var ok = TC.analyze(frontDoc(), {});
  assert.strictEqual(check(ok, 'front:affil').status, 'ok');
  var bad = TC.analyze(frontDoc({ affils: ['1Universitas Contoh, Indonesia', '2Universitas Lain, Indonesia', '2Institut Ketiga, Indonesia'] }), {});
  var a = check(bad, 'front:affil');
  assert.strictEqual(a.status, 'bad');
  assert.ok(/Nomor afiliasi 2 dipakai/.test(a.detail) && /bernomor 3 tidak punya/.test(a.detail), a.detail);
  assert.strictEqual(a.location.index, 4, 'komentar di baris afiliasi yang salah');
  var table = TC.analyze(frontDoc({ byline: 'Ani Budi1□, Citra Dewi2, Eko Fajar2', affils: ['1Universitas Contoh, Indonesia', '2Universitas Lain, Indonesia', '2Universitas lain, indonesia'] }), {});
  assert.strictEqual(check(table, 'front:affil').status, 'ok', 'afiliasi sama yang diulang per penulis bukan masalah');
});
test('header menyebut jurnal lain; DOI terpotong; isian templat', function () {
  var r = TC.analyze(frontDoc({ header: 'Jurnal Contoh Vol. 3 No. 1 (2025)\u000bFundamental and Applied Management Journal' }), {});
  assert.strictEqual(check(r, 'front:journal').status, 'bad');
  assert.ok(/Fundamental and Applied Management Journal/.test(check(r, 'front:journal').detail));
  assert.strictEqual(check(TC.analyze(frontDoc(), {}), 'front:journal').status, 'ok');
  var d = TC.analyze(frontDoc({ cite: 'Budi, A., Dewi, C., & Fajar, E. (2025). A study of things in places. Jurnal Contoh, 3(1), 5–7. https://doi.org/10.1234/jc.v3i1.' }), {});
  assert.ok(/DOI tampak belum lengkap/.test(check(d, 'cite:meta').detail));
  var p = TC.analyze(frontDoc({ cite: 'Budi, A., Dewi, C., & Fajar, E. (2025). A study of things in places. Jurnal Contoh, X(X), XX-XX. Doi. xxxx' }), {});
  assert.strictEqual(check(p, 'cite:placeholder').status, 'bad');
  assert.ok(!check(p, 'cite:meta').detail || !/berbeda dengan header/.test(check(p, 'cite:meta').detail), 'nama jurnal tidak dituduh beda karena isian templat');
});
test('"Page X of N": hanya angka yang diketik yang dibandingkan, field NUMPAGES dilewati', function () {
  function fld(code, val) { return '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ' + code + ' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>' + val + '</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'; }
  var typed = TC.analyze(frontDoc({ footer: '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' + fld('PAGE', '1') + '<w:r><w:t xml:space="preserve"> of 9</w:t></w:r>' }), {});
  assert.ok(/tidak sesuai/.test(check(typed, 'cite:pageof').title));
  var field = TC.analyze(frontDoc({ footer: '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' + fld('PAGE', '2') + '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' + fld('NUMPAGES', '2') }), {});
  assert.strictEqual(check(field, 'cite:pageof'), undefined);
});
test('contoh Nasreen: semua masalah halaman judul & How to cite tertangkap', function () {
  return sample('nasreen-98-112.docx').then(function (m) {
    if (!m) { console.log('        (lewati: fixture tidak ada)'); return; }
    var r = TC.analyze(m, {});
    assert.strictEqual(check(r, 'words:abstract').status, 'bad');
    assert.deepStrictEqual([check(r, 'cite:pages').fix.find, check(r, 'cite:pages').fix.replace], ['98–116', '98–112']);
    assert.strictEqual(check(r, 'cite:title').status, 'bad');
    assert.ok(/DOI tampak belum lengkap/.test(check(r, 'cite:meta').detail));
    assert.strictEqual(check(r, 'front:affil').status, 'bad');
    assert.strictEqual(check(r, 'front:journal').status, 'bad');
    assert.strictEqual(check(r, 'kw').status, 'bad');
  });
});

(function next(i) {
  if (i >= queue.length) { console.log('\n' + passed + ' passed, ' + failed + ' failed.'); if (failed) process.exit(1); return; }
  var t = queue[i];
  Promise.resolve().then(t[1]).then(function () { passed++; console.log('  PASS  ' + t[0]); }, function (e) { failed++; console.log('  FAIL  ' + t[0] + '\n        ' + (e && e.stack || e)); }).then(function () { next(i + 1); });
})(0);
