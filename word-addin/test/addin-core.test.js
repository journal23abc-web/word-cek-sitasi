// Tes Node untuk logika add-in Word (tanpa Office.js, tanpa jaringan).
//   node word-addin/test/addin-core.test.js
var assert = require('assert');
var AC = require('../addin-core.js');

var passed = 0, failed = 0, pending = [];
function test(name, fn) {
  try {
    var r = fn();
    if (r && r.then) { pending.push(r.then(function () { passed++; console.log('  PASS  ' + name); }, function (e) { failed++; console.log('  FAIL  ' + name + '\n        ' + (e && e.stack || e)); })); return; }
    passed++; console.log('  PASS  ' + name);
  } catch (e) { failed++; console.log('  FAIL  ' + name + '\n        ' + (e && e.stack || e)); }
}
function paras(lines) { return lines.map(function (t) { return { text: t }; }); }

var APA = [
  'Pendahuluan',
  'Pembelajaran mesin berkembang pesat (Smith & Jones, 2020). Menurut Brown et al. (2019), model besar membutuhkan data. Penelitian lain (Lee and Park, 2021; Adams, 2018) menunjukkan hal serupa. Garcia, Lopez, dan Chen (2022) membahas etika. Studi terbaru (Wang, 2023) juga relevan.',
  '',
  'Daftar Pustaka',
  'Smith, J., & Jones, K. (2020). Deep learning for text classification. Journal of Machine Learning Research, 21(4), 1-20. https://doi.org/10.1234/jmlr.2020.004',
  'Brown, T., Mann, B., Ryder, N., & Subbiah, M. (2019). Language models are few-shot learners. Advances in Neural Information Processing Systems, 33, 1877-1901.',
  'Lee, H., and Park, S. (2021). Attention mechanisms revisited. Neural Networks, 140, 12–25.',
  'Garcia, M., Lopez, R., & Chen, L. (2022). Ethics in artificial intelligence. AI and Society, 37(2), 400–415. doi:10.1007/s00146-021-01234-5',
  'Adams, P. (2018). Statistical learning theory. Springer.',
  'Lampiran A',
  'Kuesioner penelitian (Smith & Jones, 2020).'
];
var IEEE = [
  'I. Introduction',
  'Deep learning has advanced rapidly [1]. Brown et al. [3] showed large models need data. Other studies [2, 4] agree, and attention was revisited in [5]-[6]. See also [9].',
  'References',
  '[1] J. Smith and K. Jones, "Deep learning for text classification," J. Mach. Learn. Res., vol. 21, no. 4, pp. 1-20, 2020, doi: 10.1234/jmlr.2020.004.',
  '[2] H. Lee and S. Park, "Attention mechanisms revisited," Neural Netw., vol. 140, pp. 12–25, 2021.',
  '[3] T. Brown, B. Mann, N. Ryder, and M. Subbiah, "Language models are few-shot learners," in Proc. Adv. Neural Inf. Process. Syst., 2019, pp. 1877–1901.',
  '[4] P. Adams, Statistical Learning Theory. New York, NY, USA: Springer, 2018.',
  '[5] M. Garcia, R. Lopez, and L. Chen, "Ethics in artificial intelligence," AI Soc., vol. 37, no. 2, pp. 400–415, 2022',
  '[6] Q. Zhou, "Unused reference example," Data Sci. J., vol. 14, pp. 1–9, 2015.'
];

function at(st, s) { return st.paragraphs[s.paraIndex].text.substr(s.offset, s.length); }

console.log('\n=== Struktur dokumen ===');
test('heading daftar pustaka & batas lampiran terdeteksi; paragraf kosong dilewati', function () {
  var st = AC.structure(paras(APA));
  assert.strictEqual(st.headingIndex, 3);
  assert.strictEqual(st.refs.length, 5);
  assert.strictEqual(st.endIndex, 9);
  assert.ok(st.articleText.indexOf('Lampiran') < 0);
});
test('penomoran otomatis Word ([1] dari listString) ikut dibaca untuk IEEE', function () {
  var p = paras(IEEE.slice(0, 3)).concat(IEEE.slice(3).map(function (t) { var m = t.match(/^\[(\d+)\]\s*(.*)$/); return { text: m[2], isListItem: true, listString: '[' + m[1] + ']' }; }));
  var st = AC.structure(p);
  assert.ok(st.refs.every(function (r) { return r.labelFromList && /^\[\d+\] /.test(r.line); }));
  assert.strictEqual(AC.detectStyle(st).styleId, 'ieee');
  var v = AC.validate(st, 'ieee');
  v.suggestions.filter(function (s) { return s.kind === 'paragraph' && s.replacement; }).forEach(function (s) {
    assert.ok(!/^\[\d+\]/.test(s.replacement), 'label otomatis tidak boleh ditulis ulang sebagai teks: ' + s.replacement);
  });
});
test('karakter kontrol Word diganti satu-banding-satu (offset tetap)', function () {
  var st = AC.structure([{ text: 'A\u000bB (Smith, 2020)' }]);
  assert.strictEqual(st.paragraphs[0].text.length, 'A\u000bB (Smith, 2020)'.length);
});

console.log('\n=== Validator APA 7 ===');
var stA = AC.structure(paras(APA));
var vA = AC.validate(stA, 'apa7');
test('deteksi gaya APA 7', function () { assert.strictEqual(AC.detectStyle(stA).styleId, 'apa7'); });
test('setiap saran "replace" menunjuk teks yang persis ada di paragraf', function () {
  var reps = vA.suggestions.filter(function (s) { return s.kind === 'replace'; });
  assert.ok(reps.length >= 4);
  reps.forEach(function (s) { assert.strictEqual(at(stA, s), s.original, s.title); assert.ok(s.occurrence >= 0); });
});
test('saran "&" untuk dua penulis dalam grup sitasi', function () {
  var s = vA.suggestions.find(function (x) { return x.original === 'Lee and Park, 2021'; });
  assert.ok(s && s.canApply); assert.strictEqual(s.replacement, 'Lee & Park, 2021');
});
test('saran et al. untuk sitasi naratif 3 penulis', function () {
  var s = vA.suggestions.find(function (x) { return /^Garcia, Lopez/.test(x.original || ''); });
  assert.strictEqual(s.replacement, 'Garcia et al. (2022)');
});
test('sitasi tanpa referensi bisa dicari online', function () {
  var s = vA.suggestions.find(function (x) { return x.original === 'Wang, 2023'; });
  assert.ok(s && s.searchable && !s.canApply);
});
test('referensi: en dash, "&", dan DOI APA disarankan sebagai teks lengkap bercetak miring', function () {
  var lee = vA.suggestions.find(function (x) { return x.kind === 'paragraph' && /^Lee/.test(x.original); });
  assert.ok(/Lee, H\., & Park, S\./.test(lee.replacement));
  assert.ok(lee.replacementRuns.some(function (r) { return r.italic && /Neural Networks/.test(r.text); }));
  var garcia = vA.suggestions.find(function (x) { return x.kind === 'paragraph' && /^Garcia/.test(x.original); });
  assert.ok(/https:\/\/doi\.org\/10\.1007\/s00146-021-01234-5$/.test(garcia.replacement), garcia.replacement);
  assert.ok(garcia.replacement.indexOf('doi:') < 0);
  var smith = vA.suggestions.find(function (x) { return x.kind === 'paragraph' && /^Smith/.test(x.original); });
  assert.ok(/1–20/.test(smith.replacement));
});
test('urutan alfabetis → aksi "Urutkan" dengan permutasi lengkap', function () {
  var s = vA.suggestions.find(function (x) { return x.kind === 'sort'; });
  assert.ok(s && s.canApply);
  assert.deepStrictEqual(s.order.map(function (i) { return stA.refs[i].line.split(',')[0]; }), ['Adams', 'Brown', 'Garcia', 'Lee', 'Smith']);
});
test('rencana cek cetak miring untuk entri yang teksnya sudah benar', function () {
  var adams = vA.italicPlan.find(function (p) { return /^Adams/.test(stA.refs[p.refIndex].line); });
  assert.ok(adams && adams.segments.some(function (g) { return g.text === 'Statistical learning theory'; }), JSON.stringify(vA.italicPlan));
});

console.log('\n=== Validator IEEE ===');
var stI = AC.structure(paras(IEEE));
var vI = AC.validate(stI, 'ieee');
test('kelompok sitasi IEEE diberi koreksi pasti', function () {
  assert.strictEqual(vI.suggestions.find(function (s) { return s.original === '[2, 4]'; }).replacement, '[2], [4]');
  assert.strictEqual(vI.suggestions.find(function (s) { return s.original === '[5]-[6]'; }).replacement, '[5], [6]');
});
test('nama IEEE "T. Brown, B. Mann" tidak dituduh terbalik', function () {
  assert.ok(!vI.suggestions.some(function (s) { return /^\[3\]/.test(s.original || '') && /Urutan nama/.test(s.title); }));
});
test('entri IEEE tanpa titik akhir diperbaiki & nomor asli dipertahankan', function () {
  var s = vI.suggestions.find(function (x) { return x.kind === 'paragraph' && /^\[5\]/.test(x.original); });
  assert.ok(/^\[5\] /.test(s.replacement) && /2022\.$/.test(s.replacement));
});
test('urutan kemunculan salah → aksi penomoran ulang', function () {
  assert.ok(vI.suggestions.some(function (s) { return s.kind === 'restyle'; }));
});
test('saran yang membuang nama penulis diblokir', function () {
  var bad = AC.structure(paras(IEEE.slice(0, 3).concat(['[1] Lee, H., and Park, S., "Attention mechanisms revisited," Neural Netw., vol. 140, pp. 12–25, 2021.'])));
  var v = AC.validate(bad, 'ieee');
  var s = v.suggestions.find(function (x) { return x.kind === 'paragraph'; });
  if (s && s.replacement) { assert.ok(!s.canApply); assert.ok(/park/i.test(s.blockedReason)); }
});
test('coreksi IEEE tidak menghilangkan titik singkatan prosiding', function () {
  var conv = AC.planConversion(stI, 'ieee', 'ieee');
  assert.ok(conv.references.lines.some(function (l) { return /Process\. Syst\., 2019/.test(l.text); }));
});

console.log('\n=== Tautan ===');
test('APA: setiap bagian grup sitasi ditautkan ke referensinya sendiri', function () {
  var plan = AC.planLinks(stA, 'apa7');
  var by = {}; plan.citationLinks.forEach(function (l) { by[l.text] = stA.refs[l.refIndex].line.split(',')[0]; assert.strictEqual(at(stA, l), l.text); });
  assert.strictEqual(by['Lee and Park, 2021'], 'Lee');
  assert.strictEqual(by['Adams, 2018'], 'Adams');
  assert.strictEqual(by['Brown et al. (2019)'], 'Brown');
  assert.ok(plan.unmatched.some(function (u) { return /Wang/.test(u.text); }));
  assert.ok(plan.externalLinks.some(function (e) { return e.href === 'https://doi.org/10.1007/s00146-021-01234-5' && e.text === '10.1007/s00146-021-01234-5'; }));
  plan.citationLinks.forEach(function (l) { assert.ok(/^_CS_REF_\d+$/.test(l.bookmark) && l.bookmark.length <= 40); });
});
test('IEEE: nomor dalam grup [2, 4] ditautkan satu per satu (whole word)', function () {
  var plan = AC.planLinks(stI, 'ieee');
  var nums = plan.citationLinks.filter(function (l) { return l.wholeWord; }).map(function (l) { return l.text + '→' + stI.refs[l.refIndex].line.slice(0, 3); });
  assert.ok(nums.indexOf('2→[2]') >= 0 && nums.indexOf('4→[4]') >= 0, nums.join(' '));
  plan.citationLinks.forEach(function (l) { assert.strictEqual(at(stI, l), l.text); });
  assert.ok(plan.unmatched.some(function (u) { return u.text === '[9]'; }));
});

console.log('\n=== Konversi ===');
test('APA 7 → IEEE: sitasi bernomor sesuai urutan kemunculan, daftar berlabel [n]', function () {
  var plan = AC.planConversion(stA, 'apa7', 'ieee');
  var map = {}; plan.edits.forEach(function (e) { map[e.original] = e.replacement; assert.strictEqual(stA.paragraphs[e.paraIndex].text.substr(e.offset, e.length), e.original); });
  assert.strictEqual(map['(Smith & Jones, 2020)'], '[1]');
  assert.strictEqual(map['(Lee and Park, 2021; Adams, 2018)'], '[3], [4]');
  assert.ok(/^\[1\] J\. Smith and K\. Jones/.test(plan.references.lines[0].text));
  assert.ok(plan.references.lines.every(function (l, i) { return l.text.indexOf('[' + (i + 1) + '] ') === 0; }));
  assert.ok(plan.unmatched.some(function (u) { return /Wang/.test(u.raw); }));
});
test('IEEE → APA 7: daftar alfabetis, tanpa label, cetak miring nama jurnal', function () {
  var plan = AC.planConversion(stI, 'ieee', 'apa7');
  var firsts = plan.references.lines.map(function (l) { return l.text.split(',')[0]; });
  assert.deepStrictEqual(firsts, firsts.slice().sort());
  assert.ok(plan.references.lines.every(function (l) { return !/^\[\d+\]/.test(l.text); }));
  assert.ok(plan.references.detachList);
  assert.ok(plan.edits.some(function (e) { return e.original === '[2, 4]' && /Adams, 2018; Lee & Park, 2021/.test(e.replacement); }), JSON.stringify(plan.edits));
});
test('konversi menandai entri yang gagal dibaca (nama hilang) sebagai berisiko', function () {
  var st = AC.structure(paras(['Intro', 'Text [1].', 'References', '[1] Lee, H., and Park, S., "Attention mechanisms revisited," Neural Netw., vol. 140, pp. 12–25, 2021.']));
  var plan = AC.planConversion(st, 'ieee', 'apa7');
  assert.ok(plan.references.lines[0].risky);
});

test('referensi tak terbaca: dikecualikan, dilaporkan per entri, sisanya tetap dikonversi', function () {
  var st = AC.structure(paras(['Intro', 'Teks (Smith & Jones, 2020) dan (Adams, 2018).', 'Daftar Pustaka',
    'Smith, J., & Jones, K. (2020). Deep learning for text classification. Journal of Machine Learning Research, 21(4), 1–20.',
    'Adams P Statistical learning theory Springer']));
  var plan = AC.planConversion(st, 'apa7', 'ieee');
  assert.ok(!plan.error, plan.error);
  assert.strictEqual(plan.failedRefs.length, 1);
  assert.strictEqual(plan.failedRefs[0].paraIndex, 4);
  assert.ok(plan.failedRefs[0].reason);
  assert.deepStrictEqual(plan.references.paraIndexes, [3]);
  assert.deepStrictEqual(plan.references.flagParaIndexes, [4]);
  assert.ok(plan.unmatched.some(function (u) { return u.raw === '(Adams, 2018)' && u.paraIndex === 1 && st.paragraphs[1].text.substr(u.offset, u.length) === u.raw; }));
  var v = AC.validate(st, 'apa7');
  assert.ok(v.suggestions.some(function (s) { return s.title === 'Referensi tidak terbaca' && s.paraIndex === 4; }));
});
test('tiap entri hasil konversi membawa status & lokasinya', function () {
  var plan = AC.planConversion(stI, 'ieee', 'apa7');
  plan.references.lines.forEach(function (l) { assert.ok(/^(?:ok|review|risky)$/.test(l.status)); assert.ok(l.paraIndex != null); });
  assert.ok(plan.counts.total === stI.refs.length);
});
test('nama jurnal kapital semua dari CrossRef dirapikan', function () {
  assert.strictEqual(AC.tidyCaps('AI & SOCIETY'), 'AI & Society');
  assert.strictEqual(AC.tidyCaps('JOURNAL OF THE IEEE'), 'Journal of the IEEE');
  assert.strictEqual(AC.tidyCaps('Neural Networks'), 'Neural Networks');
});

console.log('\n=== CrossRef (fetch tiruan) ===');
function fakeFetch(routes) {
  return function (url) {
    var hit = routes.find(function (r) { return r[0].test(url); });
    if (!hit) return Promise.resolve({ ok: false, status: 500, json: function () { return Promise.resolve({}); } });
    return Promise.resolve({ ok: hit[1] !== 404, status: hit[1] === 404 ? 404 : 200, json: function () { return Promise.resolve(hit[2]); } });
  };
}
var GARCIA_CR = { status: 'ok', message: { DOI: '10.1007/s00146-021-01234-5', type: 'journal-article', title: ['Ethics in artificial intelligence'], author: [{ family: 'Garcia', given: 'Maria' }, { family: 'Lopez', given: 'Rosa' }, { family: 'Chen', given: 'Li' }], 'container-title': ['AI & SOCIETY'], 'short-container-title': ['AI & Soc.'], volume: '37', issue: '2', page: '400-415', issued: { 'date-parts': [[2022, 3]] } } };
var BROWN_SEARCH = { status: 'ok', message: { items: [
  { DOI: '10.5555/3495724.3495883', type: 'proceedings-article', title: ['Language models are few-shot learners'], author: [{ family: 'Brown', given: 'Tom B.' }], 'container-title': ['Advances in Neural Information Processing Systems'], issued: { 'date-parts': [[2019]] }, page: '1877-1901' },
  { DOI: '10.9999/other', title: ['Something else entirely'], author: [{ family: 'Doe', given: 'J' }], issued: { 'date-parts': [[2010]] } }
] } };
var refsA = vA.references;
function refOf(re) { return refsA.find(function (r) { return re.test(r.raw); }); }

test('DOI valid & metadata cocok → terverifikasi', function () {
  return AC.verifyReference(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, GARCIA_CR]])).then(function (r) { assert.strictEqual(r.kind, 'verified'); });
});
test('DOI tidak terdaftar → peringatan (bisa DataCite), tanpa auto-fix', function () {
  return AC.verifyReference(stA, 'apa7', refOf(/^Smith/), fakeFetch([[/works\/10\.1234/, 404, {}]])).then(function (r) { assert.strictEqual(r.severity, 'suggestion'); assert.ok(!r.canApply); assert.ok(/DataCite/.test(r.description)); });
});
test('DOI milik karya lain → error, TIDAK menawarkan penggantian dari CrossRef', function () {
  var other = JSON.parse(JSON.stringify(GARCIA_CR)); other.message.title = ['Quantum chromodynamics on the lattice']; other.message.author = [{ family: 'Kowalski', given: 'Jan' }];
  return AC.verifyReference(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, other]])).then(function (r) {
    assert.strictEqual(r.severity, 'error'); assert.ok(!r.canApply); assert.ok(!r.replacement); assert.ok(/Quantum/.test(r.description));
  });
});
test('perbedaan kecil dengan CrossRef (kapital, subjudul, tahun online/cetak) tidak dijadikan saran', function () {
  var minor = JSON.parse(JSON.stringify(GARCIA_CR));
  minor.message.title = ['ETHICS IN ARTIFICIAL INTELLIGENCE: A critical review'];
  minor.message.issued = { 'date-parts': [[2021, 11]] }; minor.message['published-print'] = { 'date-parts': [[2022, 6]] };
  return AC.verifyReference(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, minor]])).then(function (r) { assert.strictEqual(r.kind, 'verified'); });
});
test('tahun beda jauh tapi judul & penulis cocok → catatan tenang, tanpa aksi', function () {
  var y = JSON.parse(JSON.stringify(GARCIA_CR)); y.message.issued = { 'date-parts': [[2015]] };
  return AC.verifyReference(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, y]])).then(function (r) {
    assert.ok(r.quiet && !r.canApply && r.severity !== 'error');
  });
});
test('tanpa DOI → DOI ditemukan dan ditambahkan di akhir entri (APA: URL)', function () {
  return AC.verifyReference(stA, 'apa7', refOf(/^Brown/), fakeFetch([[/query\.bibliographic/, 200, BROWN_SEARCH]])).then(function (r) {
    assert.strictEqual(r.kind, 'append'); assert.strictEqual(r.appendText, ' https://doi.org/10.5555/3495724.3495883');
  });
});
test('tanpa DOI (IEEE) → ", doi: …." menggantikan titik akhir', function () {
  var ref = vI.references.find(function (r) { return /^\[3\]/.test(r.raw); });
  return AC.verifyReference(stI, 'ieee', ref, fakeFetch([[/query\.bibliographic/, 200, BROWN_SEARCH]])).then(function (r) {
    assert.ok(r.stripTrailingPeriod); assert.strictEqual(r.appendText, ', doi: 10.5555/3495724.3495883.');
    assert.ok(/1877–1901, doi: 10\.5555\/3495724\.3495883\.$/.test(r.replacement));
  });
});
test('buku tidak dituntut DOI', function () {
  return AC.verifyReference(stA, 'apa7', refOf(/^Adams/), fakeFetch([])).then(function (r) { assert.strictEqual(r, null); });
});
test('item CrossRef diformat ke APA 7 dan IEEE (cetak miring)', function () {
  var apa = AC.formatCrossrefReference(GARCIA_CR.message, 'apa7');
  assert.strictEqual(apa.text, 'Garcia, M., Lopez, R., & Chen, L. (2022). Ethics in artificial intelligence. AI & Society, 37(2), 400–415. https://doi.org/10.1007/s00146-021-01234-5');
  assert.ok(apa.runs.some(function (r) { return r.italic && /AI & Society/.test(r.text); }));
  var ieee = AC.formatCrossrefReference(GARCIA_CR.message, 'ieee', 7);
  assert.ok(/^\[7\] M\. Garcia, R\. Lopez, and L\. Chen, "Ethics in artificial intelligence,"/.test(ieee.text), ieee.text);
});
test('pencarian kandidat penulis + tahun', function () {
  var seen;
  var f = function (url) { seen = url; return fakeFetch([[/./, 200, BROWN_SEARCH]])(url); };
  return AC.searchCandidates({ author: 'Brown', year: '2019', keywords: 'few-shot' }, f).then(function (c) {
    assert.ok(/query\.author=Brown/.test(seen) && /from-pub-date%3A2019/.test(seen));
    assert.strictEqual(c[0].title, 'Language models are few-shot learners');
  });
});
test('posisi sisip: APA alfabetis, IEEE di akhir dengan nomor berikutnya', function () {
  var sorted = AC.structure(paras(['Intro', 'x', 'Daftar Pustaka', 'Adams, P. (2018). A. B.', 'Garcia, M. (2022). C. D.', 'Lee, H. (2021). E. F.', 'Smith, J. (2020). G. H.']));
  var p = AC.insertionPoint(sorted, 'apa7', 'Kim, A. (2020). X.');
  assert.ok(/^Lee/.test(sorted.paragraphs[p.paraIndex].text) && p.where === 'Before');
  var last = AC.insertionPoint(sorted, 'apa7', 'Zhou, Q. (2015). X.');
  assert.ok(/^Smith/.test(sorted.paragraphs[last.paraIndex].text) && last.where === 'After');
  var q = AC.insertionPoint(stI, 'ieee', '');
  assert.strictEqual(q.number, 7); assert.strictEqual(q.where, 'After');
});
test('pengayaan konversi: nama jurnal lengkap & singkatan dari CrossRef', function () {
  var st = AC.structure(paras(['Intro', 'Ethics matter [1].', 'References', '[1] M. Garcia, R. Lopez, and L. Chen, "Ethics in artificial intelligence," AI Soc., vol. 37, no. 2, pp. 400–415, 2022, doi: 10.1007/s00146-021-01234-5.']));
  return AC.buildEnrichmentMetadata(st, 'ieee', fakeFetch([[/works\/10\.1007/, 200, GARCIA_CR]])).then(function (meta) {
    assert.ok(meta && meta.entries.length === 1);
    var plan = AC.planConversion(st, 'ieee', 'apa7', { metadata: meta });
    assert.ok(/AI & Society, 37\(2\), 400–415/.test(plan.references.lines[0].text), plan.references.lines[0].text);
    assert.ok(plan.references.lines[0].enriched.indexOf('nama lengkap jurnal/prosiding') >= 0);
  });
});

console.log('\n=== Metadata (setara versi web) ===');
test('status DOI: valid / minor / mismatch / notfound / nodoi / notexpected', function () {
  var ok = AC.checkMetadata(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, GARCIA_CR]]));
  var minor = (function () { var m = JSON.parse(JSON.stringify(GARCIA_CR)); m.message.issued = { 'date-parts': [[2015]] }; return AC.checkMetadata(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, m]])); })();
  var other = (function () { var m = JSON.parse(JSON.stringify(GARCIA_CR)); m.message.title = ['Quantum chromodynamics']; m.message.author = [{ family: 'Kowalski' }]; return AC.checkMetadata(stA, 'apa7', refOf(/^Garcia/), fakeFetch([[/works\/10\.1007/, 200, m]])); })();
  var nf = AC.checkMetadata(stA, 'apa7', refOf(/^Smith/), fakeFetch([[/works\/10\.1234/, 404, {}]]));
  var nodoi = AC.checkMetadata(stA, 'apa7', refOf(/^Lee/), fakeFetch([]));
  var book = AC.checkMetadata(stA, 'apa7', refOf(/^Adams/), fakeFetch([]));
  return Promise.all([ok, minor, other, nf, nodoi, book]).then(function (r) {
    assert.deepStrictEqual(r.map(function (x) { return x.status; }), ['valid', 'minor', 'mismatch', 'notfound', 'nodoi', 'notexpected']);
    assert.ok(r[2].fields.some(function (f) { return f.field === 'Judul' && !f.ok && /Quantum/.test(f.crossref); }));
    assert.ok(r.every(function (x) { return x.paraIndex != null; }));
  });
});
test('jaringan gagal → status error, bukan dituduh salah', function () {
  var bad = function () { return Promise.reject(new Error('offline')); };
  return AC.checkMetadata(stA, 'apa7', refOf(/^Garcia/), bad).then(function (r) { assert.strictEqual(r.status, 'error'); });
});
test('Cari DOI: kandidat diberi skor & tanda pasti, DOI lama tidak ditawarkan ulang', function () {
  var ref = refOf(/^Brown/);
  return AC.searchDoiCandidates(ref, 'apa7', fakeFetch([[/query\.bibliographic/, 200, BROWN_SEARCH]])).then(function (c) {
    assert.strictEqual(c[0].doi, '10.5555/3495724.3495883'); assert.ok(c[0].sure); assert.ok(c[0].score >= 90);
    assert.ok(c[c.length - 1].score < 50);
  });
});
test('Pakai DOI: mengganti DOI lama di tempatnya, atau menambah di akhir entri', function () {
  var rep = AC.doiEdit(stA, 'apa7', { refIndex: 0, doi: '10.1234/jmlr.2020.004' }, '10.9999/new');
  assert.strictEqual(rep.kind, 'replace'); assert.strictEqual(rep.original, '10.1234/jmlr.2020.004');
  assert.strictEqual(stA.paragraphs[rep.paraIndex].text.substr(rep.offset, rep.length), rep.original);
  var app = AC.doiEdit(stA, 'apa7', { refIndex: 2, doi: null }, '10.9999/new');
  assert.strictEqual(app.kind, 'append'); assert.strictEqual(app.appendText, ' https://doi.org/10.9999/new');
  var ieee = AC.doiEdit(stI, 'ieee', { refIndex: 1, doi: null }, '10.9999/new');
  assert.ok(ieee.stripTrailingPeriod && ieee.appendText === ', doi: 10.9999/new.');
});
test('rentang tahun: persentase referensi terbaru', function () {
  var yr = AC.yearRange(vA.references, 5, 2024);
  assert.strictEqual(yr.from, 2020); assert.strictEqual(yr.to, 2024);
  assert.strictEqual(yr.inRange.length, 3); assert.strictEqual(yr.outRange.length, 2); assert.strictEqual(yr.percent, 60);
});
test('Scopus: nama jurnal "The …" / "&" / nama lama tetap dikenali', function () {
  var SM = require('../../scopus-engine.js');
  var db = new SM.ScopusDatabase();
  db.loadSourceListCompact([
    ['0033-5533', '1531-4650', 'Quarterly Journal of Economics', 1, '1886-2026', 'Journal', 0, '', ''],
    ['0044-118X', null, 'Youth and Society', 1, '1969-2026', 'Journal', 0, '', ''],
    ['0894-587X', null, 'Administration and Policy in Mental Health', 1, '1988-2026', 'Journal', 0, '', '']]);
  var st = AC.structure(paras(['Intro', 'A (Baird et al., 2011) B (Stanton-Salazar, 2011) C (Palinkas et al., 2015).', 'References',
    'Baird, S., McIntosh, C., & Özler, B. (2011). Cash or condition? Evidence from a cash transfer experiment. The Quarterly Journal of Economics, 126(4), 1709–1753.',
    'Palinkas, L. A., Horwitz, S. M., & Green, C. A. (2015). Purposeful sampling. Administration and Policy in Mental Health and Mental Health Services Research, 42(5), 533–544.',
    'Stanton-Salazar, R. D. (2011). A social capital framework. Youth & Society, 43(3), 1066–1109.']));
  var v = AC.validate(st, 'apa7');
  var res = AC.scopusStatuses(v.references, db, SM);
  assert.deepStrictEqual(res.map(function (x) { return x.status; }), ['SCOPUS_SOURCE_ONLY', 'SCOPUS_SOURCE_ONLY', 'SCOPUS_SOURCE_ONLY']);
});

// ---------- Harvard (umum / Cite Them Right) & Harvard Emerald ----------
var HARV = [
  'Introduction',
  'Machine learning has grown quickly (Smith and Jones, 2020). According to Brown et al. (2019), large models need data. Other studies (Lee & Park, 2021; Adams, 2018) agree. Garcia, Lopez and Chen (2022) discuss ethics. A recent study (Wang, 2023) is relevant.',
  'References',
  'Adams, P. (2018) Statistical learning theory. New York: Springer.',
  "Brown, T., Mann, B., Ryder, N. and Subbiah, M. (2019) 'Language models are few-shot learners', Advances in Neural Information Processing Systems, 33, pp. 1877-1901.",
  "Garcia, M., Lopez, R. and Chen, L. (2022) 'Ethics in artificial intelligence', AI and Society, 37(2), pp. 400–415. doi:10.1007/s00146-021-01234-5",
  "Lee, H. & Park, S. (2021) 'Attention mechanisms revisited', Neural Networks, 140, pp. 12–25.",
  "Smith, J. and Jones, K. (2020) 'Deep learning for text classification', Journal of Machine Learning Research, 21(4), pp. 1–20. Available at: https://doi.org/10.1234/jmlr.2020.004"
];
var CR_ITEM = { DOI: '10.5555/3295222.3295349', type: 'journal-article', title: ['Attention is all you need'],
  author: [{ family: 'Vaswani', given: 'Ashish' }, { family: 'Shazeer', given: 'Noam' }], issued: { 'date-parts': [[2017]] },
  'container-title': ['Advances in Neural Information Processing Systems'], volume: '30', page: '5998-6008' };
var EMER = [
  'Introduction',
  'Tax administrations increasingly use analytics (OECD, 2020). Compliance theory (Allingham and Sandmo, 1972; Alm, 2019) frames the debate, and Belahouaoui and Attak (2024) review the field. Fraud detection is studied by Behrends (2022) and in group work (Williams and Calabrese, 2016; Williams & Kavanagh, 2016).',
  'References',
  'Allingham, M.G. and Sandmo, A. (1972), "Income tax evasion: a theoretical analysis", Journal of Public Economics, Vol. 1 No. 3-4, pp. 323-338, doi: 10.1016/0047-2727(72)90010-2.',
  'Alm, J. (2019), "What motivates tax compliance?", Journal of Economic Surveys, Vol. 33 No. 2, pp. 353-388, doi: 10.1111/joes.12272.',
  'Behrends, M.P. (2022), "Under which conditions would the ATAD\'s GAAR benefit from automatic decision-making?", International Tax Studies, pp. 2-19, doi: 10.59403/2gyn7pt.',
  'Belahouaoui, R. and Attak, E.H. (2024), "Digital taxation, artificial intelligence and Tax Administration 3.0", Accounting Research Journal, Vol. 37 No. 2, pp. 172-191, doi: 10.1108/arj-12-2023-0372.',
  'OECD (2020), Tax Administration 3.0: The Digital Transformation of Tax Administration, OECD Publishing, Paris, doi: 10.1787/ca274cc5-en.',
  'Williams, D.W. and Calabrese, T.D. (2016), "The status of budget forecasting", Journal of Public and Nonprofit Affairs, Vol. 2 No. 2, pp. 127-160, doi: 10.20899/jpna.2.2.127-160.',
  'Williams, D.W. and Kavanagh, S.C. (2016), "Local government revenue forecasting methods: competition and comparison", Journal of Public Budgeting, Accounting & Financial Management, Vol. 28 No. 4, pp. 488-526.'
];

console.log('\n=== Harvard (umum) ===');
var stH = AC.structure(paras(HARV));
var vH = AC.validate(stH, 'harvard');
test('deteksi gaya Harvard (bukan lagi dipaksa APA 7)', function () {
  var d = AC.detectStyle(stH);
  assert.strictEqual(d.styleId, 'harvard'); assert.strictEqual(d.note, null);
  assert.strictEqual(AC.detectStyle(AC.structure(paras(APA))).styleId, 'apa7', 'naskah APA tetap APA');
  assert.strictEqual(AC.detectStyle(AC.structure(paras(IEEE))).styleId, 'ieee', 'naskah IEEE tetap IEEE');
});
test('sitasi: "&" dalam grup → "and"; et al. naratif; urutan alfabetis', function () {
  var amp = vH.suggestions.find(function (s) { return s.original === 'Lee & Park, 2021'; });
  assert.ok(amp && amp.canApply); assert.strictEqual(amp.replacement, 'Lee and Park, 2021');
  assert.strictEqual(at(stH, amp), amp.original);
  var etal = vH.suggestions.find(function (s) { return /^Garcia, Lopez and Chen/.test(s.original || ''); });
  assert.strictEqual(etal.replacement, 'Garcia et al. (2022)');
  var missing = vH.suggestions.find(function (s) { return s.original === 'Wang, 2023'; });
  assert.ok(missing.searchable, 'sitasi tanpa referensi bisa dicari online');
});
test('"&" tunggal → "and" tanpa menghapus nomor halaman', function () {
  var st = AC.structure(paras(['Intro', 'As shown (Lee & Park, 2021, p. 14).', 'References', HARV[6]]));
  var s = AC.validate(st, 'harvard').suggestions.find(function (x) { return /Pemisah dua penulis/.test(x.title); });
  assert.strictEqual(s.replacement, '(Lee and Park, 2021, p. 14)');
});
test('referensi: tanpa koma Oxford, "and" bukan "&", en dash, "Available at:" untuk DOI', function () {
  function card(re) { return vH.suggestions.find(function (s) { return s.kind === 'paragraph' && re.test(s.original); }); }
  assert.ok(!/,\s+and\s+Subbiah/.test(card(/^Brown/).replacement), 'tanpa koma sebelum "and"');
  assert.ok(/pp\. 1877–1901/.test(card(/^Brown/).replacement));
  assert.ok(/^Lee, H\. and Park, S\./.test(card(/^Lee/).replacement));
  assert.ok(/Available at: https:\/\/doi\.org\/10\.1007/.test(card(/^Garcia/).replacement));
  assert.ok(!card(/^Smith/) || !card(/^Smith/).replacement, 'entri yang sudah benar (kutip lurus) tidak disarankan ulang');
  vH.suggestions.filter(function (s) { return s.kind === 'paragraph'; }).forEach(function (s) { assert.ok(s.canApply, s.original); });
});
test('catatan "varian Harvard" tampil sekali per dokumen, bukan per entri', function () {
  assert.ok(vH.styleNote && /Cite Them Right/.test(vH.styleNote));
  assert.ok(!vH.suggestions.some(function (s) { return (s.reasons || []).some(function (r) { return /Varian Harvard|varian institusi/i.test(r.title + r.description); }); }));
});
test('tautan Harvard: setiap bagian sitasi ke referensinya', function () {
  var lp = AC.planLinks(stH, 'harvard');
  assert.strictEqual(lp.citationLinks.length, 5); assert.strictEqual(lp.unmatched.length, 1);
  lp.citationLinks.forEach(function (l) { assert.strictEqual(stH.paragraphs[l.paraIndex].text.substr(l.offset, l.length), l.text); });
});
test('konversi Harvard → APA 7 / IEEE, dan APA 7 / IEEE → Harvard', function () {
  var toApa = AC.planConversion(stH, 'harvard', 'apa7');
  assert.strictEqual(toApa.counts.converted, 5);
  assert.ok(toApa.edits.some(function (e) { return e.replacement === '(Smith & Jones, 2020)'; }));
  var toIeee = AC.planConversion(stH, 'harvard', 'ieee');
  assert.ok(toIeee.references.lines.every(function (l) { return /^\[\d+\] /.test(l.text); }));
  var fromApa = AC.planConversion(AC.structure(paras(APA)), 'apa7', 'harvard');
  assert.strictEqual(fromApa.counts.risky, 0);
  assert.ok(fromApa.references.detachList, 'penomoran Word dilepas untuk gaya tanpa nomor');
  fromApa.references.lines.forEach(function (l) {
    assert.ok(!l.notes.some(function (n) { return /varian institusi/i.test(n); }), 'catatan varian bukan alasan "perlu dicek"');
    // Satu-satunya yang perlu dicek: buku APA tanpa kota terbit (Harvard menulis "Kota: Penerbit").
    assert.strictEqual(l.status, /^Adams/.test(l.text) ? 'review' : 'ok', l.text);
  });
  assert.ok(fromApa.references.lines.some(function (l) { return /^Brown, T\., Mann, B\., Ryder, N\. and Subbiah, M\. \(2019\) ‘/.test(l.text); }));
  var fromIeee = AC.planConversion(AC.structure(paras(IEEE)), 'ieee', 'harvard');
  assert.ok(fromIeee.edits.some(function (e) { return e.original === '[2, 4]' && e.replacement === '(Adams, 2018; Lee and Park, 2021)'; }));
  assert.ok(fromIeee.references.lines.some(function (l) { return l.status === 'review' && l.notes.some(function (n) { return /singkatan/.test(n); }); }), 'nama jurnal singkatan IEEE ditandai');
});
test('DOI & sisip referensi CrossRef dalam format Harvard', function () {
  var app = AC.doiEdit(stH, 'harvard', { refIndex: 3, doi: null }, '10.9999/new');
  assert.strictEqual(app.appendText, ' Available at: https://doi.org/10.9999/new'); assert.ok(!app.stripTrailingPeriod);
  var f = AC.formatCrossrefReference(CR_ITEM, 'harvard');
  assert.ok(/^Vaswani, A\. and Shazeer, N\. \(2017\) ‘Attention is all you need’, /.test(f.text), f.text);
  assert.strictEqual(AC.insertionPoint(stH, 'harvard', f.apaLine).where, 'After');
});

console.log('\n=== Harvard (Emerald) ===');
var stE = AC.structure(paras(EMER));
var vE = AC.validate(stE, 'harvard-emerald');
test('varian Emerald terdeteksi dari "(Tahun)," di daftar pustaka', function () {
  assert.strictEqual(AC.detectStyle(stE).styleId, 'harvard-emerald');
  assert.ok(/Emerald/.test(vE.styleNote));
});
test('entri Emerald yang benar tidak diubah (inisial rapat, Vol./No., tanda hubung, doi:)', function () {
  var cards = vE.suggestions.filter(function (s) { return s.kind === 'paragraph' && s.replacement; });
  assert.deepStrictEqual(cards.map(function (c) { return c.original.slice(0, 12); }), [], 'tidak ada saran tulis-ulang untuk entri yang sudah sesuai');
});
test('inisial rapat penulis terakhir ("Attak, E.H.") dan "M.P." tidak hilang', function () {
  var p = AC.planConversion(stE, 'harvard-emerald', 'apa7');
  assert.ok(p.references.lines.some(function (l) { return /Attak, E\. H\./.test(l.text); }));
  assert.ok(p.references.lines.some(function (l) { return /^Behrends, M\. P\./.test(l.text); }));
  assert.ok(p.references.lines.some(function (l) { return /^OECD\. \(2020\)\. Tax Administration 3\.0: The Digital/.test(l.text); }), 'judul buku bertitik utuh');
});
test('pengaman: saran yang membuang inisial rapat tidak bisa diterapkan', function () {
  assert.ok(AC.lostTokens('x', 'x').length === 0);
  var st = AC.structure(paras(['Intro', 'See (Behrends, 2022).', 'References', 'Behrends, M.P. (2022) Title. Publisher.']));
  var v = AC.validate(st, 'harvard');
  v.suggestions.filter(function (s) { return s.kind === 'paragraph' && s.replacement && !/M\.P\./.test(s.replacement); }).forEach(function (s) { assert.ok(!s.canApply); });
});
test('rekan penulis berbeda di tahun sama bukan "penulis sama"; "&" dalam grup terdeteksi', function () {
  assert.ok(!vE.suggestions.some(function (s) { return /penulis sama/.test(s.title); }));
  var amp = vE.suggestions.find(function (s) { return s.original === 'Williams & Kavanagh, 2016'; });
  assert.strictEqual(amp.replacement, 'Williams and Kavanagh, 2016');
});
test('konversi APA 7 → Harvard Emerald dan Emerald → IEEE', function () {
  var p = AC.planConversion(AC.structure(paras(APA)), 'apa7', 'harvard-emerald');
  var smith = p.references.lines.find(function (l) { return /^Smith/.test(l.text); });
  assert.strictEqual(smith.text, 'Smith, J. and Jones, K. (2020), "Deep learning for text classification", Journal of Machine Learning Research, Vol. 21 No. 4, pp. 1-20, doi: 10.1234/jmlr.2020.004.');
  assert.deepStrictEqual(smith.runs.filter(function (r) { return r.italic; }).map(function (r) { return r.text; }), ['Journal of Machine Learning Research']);
  var ie = AC.planConversion(stE, 'harvard-emerald', 'ieee');
  assert.strictEqual(ie.counts.risky + ie.counts.failed, 0);
  assert.ok(ie.edits.some(function (e) { return e.original === '(OECD, 2020)' && e.replacement === '[1]'; }));
});
test('DOI Emerald: ", doi: …." menggantikan titik akhir; CrossRef → Emerald', function () {
  var app = AC.doiEdit(stE, 'harvard-emerald', { refIndex: 6, doi: null }, '10.9999/new');
  assert.strictEqual(app.appendText, ', doi: 10.9999/new.'); assert.ok(app.stripTrailingPeriod);
  var f = AC.formatCrossrefReference(CR_ITEM, 'harvard-emerald');
  assert.ok(/^Vaswani, A\. and Shazeer, N\. \(2017\), "Attention is all you need", /.test(f.text), f.text);
  assert.ok(f.runs.some(function (r) { return r.italic; }));
});
test('metadata & rentang tahun memakai referensi Harvard', function () {
  var yr = AC.yearRange(vE.references, 10, 2026);
  assert.strictEqual(yr.inRange.length + yr.outRange.length, 7);
  return AC.checkMetadata(stE, 'harvard-emerald', vE.references.find(function (r) { return /^Behrends/.test(r.raw); }), fakeFetch([])).then(function (r) {
    assert.ok(r && r.doi === '10.59403/2gyn7pt');
  });
});

console.log('\n=== Laporan (Metadata & keseluruhan) ===');
var RC = require('../report-core.js');
test('laporan metadata: tahun, tanpa DOI, Scopus, ringkasan — berprefiks Metadata', function () {
  var refs = vE.references;
  var online = {}; refs.forEach(function (r) { online['harvard-emerald|' + r.raw] = /^Williams, D\.W\. and Kavanagh/.test(r.raw) ? { status: 'nodoi' } : { status: 'valid' }; });
  var scopus = refs.map(function (r) { return { ref: r, status: /Nonprofit/.test(r.raw) ? 'UNKNOWN' : 'SCOPUS_SOURCE_ONLY', confidence: 0.9, matchedSource: { title: r.journal } }; });
  var rep = RC.build(stE, [], { includeYears: true, includeDoi: true, includeNoDoi: true, includeScopus: true, scopus: scopus, years: 10, references: refs,
    online: online, styleId: 'harvard-emerald', prefix: RC.METADATA_PREFIX, summary: { style: 'Harvard (Emerald)', doi: { ok: 6, bad: 0, nodoi: 1 }, scopus: { indexed: 6, unknown: 1 } } });
  assert.strictEqual(rep.counts.nodoi, 1); assert.strictEqual(rep.counts.scopus, 1); assert.ok(rep.counts.year >= 2); assert.strictEqual(rep.counts.summary, 1);
  assert.ok(rep.items.every(function (it) { return it.comment.indexOf(RC.METADATA_PREFIX) === 0 && RC.isMetadataComment(it.comment); }));
  var title = rep.items.find(function (it) { return it.paraIndex === 0; });
  assert.ok(/^\[Cek Sitasi · Metadata\] Ringkasan laporan — Gaya sitasi: Harvard \(Emerald\)/.test(title.comment), 'ringkasan di judul');
  var sc = rep.items.find(function (it) { return /Scopus — Tidak ditemukan/.test(it.comment); });
  assert.strictEqual(sc.text, 'Journal of Public and Nonprofit Affairs', 'komentar Scopus ditempel di nama jurnal');
  rep.items.forEach(function (it) { if (!it.whole && it.length <= 240) assert.strictEqual(stE.paragraphs[it.paraIndex].text.substr(it.offset, it.length), it.text); });
});
test('laporan keseluruhan: Validator + Template + Metadata dalam satu set komentar; ringkasan jadi butir pertama', function () {
  var tpl = [{ id: 't1', countAs: 'template', severity: 'error', comment: 'Template — Abstrak melebihi batas.', color: 'Red', paraIndex: 0, offset: 0, length: 5 }];
  var rep = RC.build(stE, vE.suggestions.concat(tpl), { includeErrors: false, includeSuggestions: true, includeYears: true, years: 10, references: vE.references,
    styleId: 'harvard-emerald', summary: { style: 'Harvard (Emerald)', validator: { errors: 3, suggestions: 1 }, template: { bad: 1, warn: 0, ok: 3 } } });
  assert.strictEqual(rep.counts.template, 1, 'template tetap ikut walau "Perlu diperbaiki" Validator tidak dicentang');
  assert.strictEqual(rep.counts.error, 0);
  var first = rep.items.find(function (it) { return it.paraIndex === 0; });
  assert.ok(/^\[Cek Sitasi\] 1\) Ringkasan laporan/.test(first.comment) && /2\) Template — Abstrak/.test(first.comment));
  assert.ok(rep.items.every(function (it) { return RC.commentMatches(it.comment, 'all') && RC.commentMatches(it.comment, 'validator'); }));
});
test('jenis komentar: laporan per tab tidak saling menghapus; "all" menghapus semuanya', function () {
  var M = RC.METADATA_PREFIX + ' x', T = RC.TEMPLATE_PREFIX + ' y', V = RC.PREFIX + ' z';
  assert.ok(RC.commentMatches(M, 'metadata') && !RC.commentMatches(M, 'validator') && !RC.commentMatches(M, 'template'));
  assert.ok(!RC.commentMatches(V, 'metadata') && RC.commentMatches(V, 'validator'));
  assert.ok([M, T, V].every(function (c) { return RC.commentMatches(c, 'all'); }) && !RC.commentMatches('komentar reviewer', 'all'));
});

console.log('\n=== Utilitas ===');
test('occurrenceIndex meniru urutan hasil pencarian Word', function () {
  assert.strictEqual(AC.occurrenceIndex('[1] a [1] b [12]', '[1]', 6, false), 1);
  assert.strictEqual(AC.occurrenceIndex('1, 12, 1', '1', 7, true), 1);
});
test('wordDiff menandai kata yang berubah', function () {
  var d = AC.wordDiff('Lee, H., and Park', 'Lee, H., & Park');
  assert.ok(d.some(function (o) { return o.op === 'del' && o.text === 'and'; }) && d.some(function (o) { return o.op === 'ins' && o.text === '&'; }));
});

console.log('\n=== Kinerja ===');
test('naskah IEEE dengan judul buku tanpa tanda kutip tidak membuat validasi macet', function () {
  // Regresi: pola nama-terbalik IEEE dengan lookbehind bertingkat backtracking eksponensial
  // pada entri penulis institusi + judul laporan tanpa tanda kutip ([6]) — validasi tidak pernah
  // selesai (bila regresi terjadi lagi, tes ini menggantung, bukan gagal).
  var refs = [
    '[1] A. Bick, A. Blandin, and D. J. Deming, The rapid adoption of generative AI, NBER Working Paper No. 32966. National Bureau of Economic Research, 2024, doi: 10.3386/w32966.',
    '[2] J. Berryhill, K. K. Heang, R. Clogher, and K. McBride, Hello, World: Artificial intelligence and its use in the public sector, OECD Working Papers on Public Governance No. 36. OECD Publishing, 2019, doi: 10.1787/726fd39d-en.',
    '[3] Department of Communications and Digital Technologies, South Africa national artificial intelligence policy framework. Republic of South Africa, 2024. [Online]. Available: https://example.org/policy.pdf',
    '[6] Amnesty International, Xenophobic machines: Discrimination through unregulated use of algorithms in the Dutch childcare benefits scandal, EUR 35/4686/2021, 2021. [Online]. Available: https://www.amnesty.org/en/documents/eur35/4686/2021/en/',
    '[4] H. Margetts and C. Dorobantu, “Rethink government with AI,” Nature, vol. 568, no. 7751, pp. 163–165, 2019, doi: 10.1038/d41586-019-01099-5.',
    '[5] Lee, H., and Park, S., “Attention mechanisms revisited,” Neural Netw., vol. 140, pp. 12–25, 2021.'
  ];
  var lines = ['I. Introduction', 'Public sector adoption is uneven [1], [2]. Policy frameworks differ [3], and evidence is mixed [4], [5]. Harms are documented [6].', 'References'].concat(refs);
  var st = AC.structure(paras(lines));
  var t0 = Date.now(), v = AC.validate(st, 'ieee');
  assert.ok(Date.now() - t0 < 3000, 'validasi ' + (Date.now() - t0) + ' ms');
  var lee = v.suggestions.find(function (s) { return s.kind === 'paragraph' && /^\[5\]/.test(s.original); });
  assert.ok(lee && lee.reasons.some(function (r) { return /Urutan nama penulis IEEE/.test(r.title); }), 'nama terbalik tetap terdeteksi');
  assert.ok(!v.suggestions.some(function (s) { return /^\[4\]/.test(s.original || '') && (s.reasons || []).some(function (r) { return /Urutan nama/.test(r.title); }); }), '"H. Margetts and C. Dorobantu" tidak dituduh terbalik');
});

Promise.all(pending).then(function () {
  console.log('\n' + passed + ' passed, ' + failed + ' failed.');
  process.exit(failed ? 1 : 0);
});
