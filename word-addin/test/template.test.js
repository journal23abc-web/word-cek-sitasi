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

(function next(i) {
  if (i >= queue.length) { console.log('\n' + passed + ' passed, ' + failed + ' failed.'); if (failed) process.exit(1); return; }
  var t = queue[i];
  Promise.resolve().then(t[1]).then(function () { passed++; console.log('  PASS  ' + t[0]); }, function (e) { failed++; console.log('  FAIL  ' + t[0] + '\n        ' + (e && e.stack || e)); }).then(function () { next(i + 1); });
})(0);
