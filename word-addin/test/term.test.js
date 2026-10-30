// Tes Node: Cek Konsistensi Istilah — naskah sintetis (EN & ID) dengan pola inkonsistensi
// khas tulisan berbantuan AI, plus naskah contoh di test/fixtures bila ada.
//   node word-addin/test/term.test.js
var assert = require('assert');
var fs = require('fs');
var path = require('path');
var JSZip = require('../vendor/jszip.min.js');
var DR = require('../docx-reader.js');
var TM = require('../term-core.js');

var FIX = path.join(__dirname, 'fixtures');
var passed = 0, failed = 0, queue = [];
function test(name, fn) { queue.push([name, fn]); }

function esc(t) { return t.replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
// lines: string (paragraf biasa) atau { h: 'Heading' } (heading tebal)
function doc(lines) {
  var body = lines.map(function (l) {
    if (typeof l === 'object') return '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + esc(l.h) + '</w:t></w:r></w:p>';
    return '<w:p><w:r><w:t xml:space="preserve">' + esc(l) + '</w:t></w:r></w:p>';
  }).join('');
  return DR.read({ document: '<w:document><w:body>' + body + '<w:sectPr/></w:body></w:document>', rels: '<Relationships/>', hf: {} });
}
function find(r, pred) { return r.findings.filter(pred); }
function byId(r, re) { return r.findings.filter(function (f) { return re.test(f.id); }); }
function sample(name) {
  var f = path.join(FIX, name);
  if (!fs.existsSync(f)) return null;
  return JSZip.loadAsync(fs.readFileSync(f)).then(DR.partsFromZip).then(DR.read);
}

var EN = [
  'Perceived Usefulness and Behavioural Intention to Use Mobile Learning among University Students',
  'Rina Sari1, Budi Santoso2',
  '1Universitas Contoh, Indonesia',
  'Abstract: This study examines how perceived usefulness shapes behavioural intention to use mobile learning. Data from 312 university students were analysed using PLS-SEM. Perceived usefulness had a strong effect on behavioral intention.',
  'Keywords: behavioural intention; mobile learning; perceived usefulness; technology acceptance',
  { h: '1. Introduction' },
  'Mobile learning (m-learning) has expanded rapidly in higher education. The Technology Acceptance Model (TAM) explains adoption through perceived usefulness (PU) and perceived ease of use (PEOU). TAM has been widely used to study mobile learning.',
  'Prior work shows that PU predicts behavioural intention (BI). However, the role of Perceived Usefulness in mobile-learning contexts remains debated. Students rarely discuss perceived usefulness explicitly.',
  'In this study, perceived ease of use is expected to influence perceived usefulness, and both are expected to shape behavioural intention. Self-efficacy may also matter.',
  { h: '2. Method' },
  'A questionnaire measured perceived usefulness, perceived ease of use and behavioural intention. The Mobile Learning Readiness Scale (MLRS) was adapted for this study. Items of the perceived usefulness scale were translated.',
  'Behavioural intension was measured with four items. The students completed the online questionnaire in class. Mobile self efficacy was measured with a short scale.',
  'The Mobile Learning Readiness Inventory (MLRI) was used to screen the participants before the main survey.',
  { h: '3. Results' },
  'Perceived usefulness significantly predicted behavioural intention. Perceived ease of use predicted perceived usefulness. The m-learning readiness of students was moderate. The pre-test and pretest results were similar.',
  'Students with higher mobile learning experience reported stronger behavioural intention. The pretest scores were used as covariates.',
  'The effect of perceived usefulness on behavioural intentions was larger for students in mobile-learning courses. Data were analyzed with SmartPLS. The perceived usefulness scale showed good reliability.',
  { h: '4. Conclusion' },
  'Perceived usefulness is the key driver of behavioural intention to use mobile learning. Future research should examine actual mobile learning use and the role of perceived ease of use.',
  { h: 'References' },
  'Davis, F. D. (1989). Perceived usefulness, perceived ease of use, and user acceptance of information technology. MIS Quarterly, 13(3), 319–340.'
];

test('nama variabel: huruf kapital di tengah kalimat, tanda hubung, salah ketik', function () {
  var r = TM.analyze(doc(EN), {});
  var pu = byId(r, /^term:perceivedusefulness$/)[0];
  assert.ok(pu, 'perceived usefulness dilaporkan: ' + r.findings.map(function (f) { return f.id; }).join(', '));
  assert.ok(pu.kinds.indexOf('case') >= 0);
  assert.strictEqual(pu.status, 'warn');
  assert.ok(pu.fix.items.some(function (i) { return i.text === 'Perceived Usefulness' && i.next === 'perceived usefulness'; }), 'kapital di tengah kalimat diseragamkan');
  assert.ok(!pu.fix.items.some(function (i) { return i.zone === 'title' || i.zone === 'heading'; }), 'judul & heading tidak diubah');
  assert.ok(!pu.fix.items.some(function (i) { return /^Perceived usefulness/.test(i.text); }), 'awal kalimat tidak dianggap salah');

  assert.ok(!byId(r, /^term:mobilelearning$/).length, '"mobile-learning contexts" (kata sifat majemuk) vs "mobile learning" bukan masalah');
  var se = byId(r, /^term:selfefficacy$/)[0];
  assert.ok(se && se.kinds.indexOf('compound') >= 0, 'self-efficacy vs self efficacy (awalan terikat)');
  assert.deepStrictEqual(se.fix.items.map(function (i) { return i.next; }), ['self-efficacy']);

  var typo = byId(r, /^typo:/)[0];
  assert.ok(typo && /intension/.test(typo.title), 'behavioural intension = salah ketik');
  assert.strictEqual(typo.fix.items[0].next, 'Behavioural intention');

  var pre = byId(r, /^term:pretest$/)[0];
  assert.ok(pre && pre.fix.items[0].text === 'pre-test' && pre.fix.items[0].next === 'pretest');
});

test('ejaan UK/US bercampur & singkatan', function () {
  var r = TM.analyze(doc(EN), {});
  var sp = byId(r, /^spell:behavior/)[0];
  assert.ok(sp, 'behavioural/behavioral');
  assert.strictEqual(sp.recommended, 'behavioural');
  assert.deepStrictEqual(sp.fix.items.map(function (i) { return i.text + '→' + i.next; }), ['behavioral→behavioural']);
  var an = byId(r, /^spell:analyze/)[0];
  assert.ok(an && an.fix.items.length === 1, 'analysed/analyzed (1 lawan 1, ikut ragam dominan naskah)');
  assert.strictEqual(an.fix.items[0].next, 'analysed', 'naskah lebih banyak memakai ejaan UK');

  // TAM didefinisikan; PU dipakai sebelum didefinisikan? (tidak — didefinisikan di kalimat yang sama sebelum dipakai)
  assert.ok(!byId(r, /^acr:early:body:PU$/).length);
  // MLRS vs MLRI: dua nama untuk instrumen yang hampir sama, tak pernah disebut bersama
  var two = byId(r, /^acr:twonames:/)[0];
  assert.ok(two && /MLRS/.test(two.title) && /MLRI/.test(two.title), 'MLRS vs MLRI: ' + r.findings.map(function (f) { return f.id; }).join(', '));
  // PU didefinisikan tapi bentuk panjang masih ditulis; PEOU didefinisikan tapi tidak pernah dipakai
  assert.ok(byId(r, /^acr:full:PU$/).length, 'perceived usefulness masih ditulis lengkap');
  assert.ok(byId(r, /^acr:unused:PEOU$/).length, 'PEOU tidak dipakai');
  assert.ok(!byId(r, /acr:undef:SmartPLS/).length, 'nama perangkat lunak umum tidak diminta didefinisikan');
});

test('sinonim dalam nama variabel; dua istilah yang disebut bersama tidak dianggap sinonim', function () {
  var lines = [
    'Academic Stress among First-Year Students',
    'Abstract: We examine academic stress in first-year students. Academic stress was measured at two points.',
    'Keywords: academic stress; first-year students',
    { h: 'Introduction' },
    'Academic stress is common among first-year students. Many first-year students report academic stress during examinations.',
    'High educational stress lowers wellbeing. Interventions that reduce academic stress are needed.',
    'Job satisfaction and work engagement are distinct constructs; job satisfaction and work satisfaction were both measured.',
    { h: 'Method' },
    'Academic stress was measured with a validated scale. Job satisfaction of staff was also recorded.',
    { h: 'References' }
  ];
  var r = TM.analyze(doc(lines), {});
  var syn = byId(r, /^syn:/);
  assert.ok(syn.some(function (f) { return /academic stress/.test(f.title) && /educational stress/.test(f.title); }), 'academic/educational stress');
  assert.ok(!syn.some(function (f) { return /work satisfaction/.test(f.title); }), 'disebut bersama dalam satu kalimat → konsep berbeda');
  var s = syn.filter(function (f) { return /educational stress/.test(f.title); })[0];
  assert.strictEqual(s.fix.items[0].next, 'academic stress');
});

test('tunggal/jamak: kata penjelas dilaporkan, kata inti & pola subjek–predikat tidak', function () {
  var lines = [
    'Social Skill Training for Shy Adolescents',
    'Abstract: Social skill training reduced shyness. The findings suggest that social skills training works.',
    'Keywords: shyness; social skill training',
    { h: 'Introduction' },
    'Social skill training is widely used. We delivered social skill training to adolescents. Several social skill trainings exist.',
    'Our social skills training lasted six weeks. The finding suggests that social skill training is effective. The findings suggest the same.',
    'Levels of shyness fell; the level of shyness was measured weekly. Social skill training groups met weekly.',
    { h: 'References' }
  ];
  var r = TM.analyze(doc(lines), {});
  var sst = byId(r, /^term:socialskilltraining$/)[0];
  assert.ok(sst && sst.kinds.indexOf('number') >= 0, r.findings.map(function (f) { return f.id; }).join(', '));
  assert.strictEqual(sst.status, 'bad', 'istilah di judul/kata kunci → merah');
  assert.deepStrictEqual(sst.fix.items.map(function (i) { return i.text + '→' + i.next; }), ['social skills training→social skill training', 'social skills training→social skill training']);
  assert.ok(!byId(r, /^term:findingsuggest/).length, '"findings suggest"/"finding suggests" = tata bahasa');
  assert.ok(!byId(r, /^term:levelofshyness/).length, '"levels of"/"level of" = kata inti');
});

test('pemeriksaan bahasa Indonesia: sinonim & kapital', function () {
  var lines = [
    'Pengaruh Kualitas Pelayanan terhadap Kepuasan Pelanggan',
    'Abstrak: Penelitian ini menguji pengaruh kualitas pelayanan terhadap kepuasan pelanggan. Kualitas pelayanan diukur dengan kuesioner.',
    'Kata kunci: kepuasan pelanggan; kualitas pelayanan',
    { h: 'Pendahuluan' },
    'Kualitas pelayanan menentukan kepuasan pelanggan. Banyak perusahaan meningkatkan kualitas layanan agar kepuasan konsumen naik.',
    'Menurut teori, Kualitas Pelayanan yang baik meningkatkan loyalitas. Kepuasan pelanggan menjadi tujuan utama.',
    { h: 'Metode' },
    'Data kualitas pelayanan dan kepuasan pelanggan dikumpulkan dari 150 responden.',
    { h: 'Daftar Pustaka' }
  ];
  var r = TM.analyze(doc(lines), {});
  assert.strictEqual(r.lang, 'id');
  var syn = byId(r, /^syn:/).map(function (f) { return f.title; }).join(' | ');
  assert.ok(/kualitas layanan/.test(syn), 'kualitas pelayanan / kualitas layanan: ' + syn);
  assert.ok(/kepuasan konsumen/.test(syn), 'kepuasan pelanggan / kepuasan konsumen: ' + syn);
  var kp = byId(r, /^term:kualitaspelayanan$/)[0];
  assert.ok(kp && kp.fix.items.some(function (i) { return i.text === 'Kualitas Pelayanan' && i.next === 'kualitas pelayanan'; }));
});

test('kata kunci salah ketik diberi usulan', function () {
  var lines = [
    'Behavioural Dimension of Stress',
    'Abstract: The behavioural dimension of stress was examined in the behavioural dimension subscale.',
    'Keywords: Behavioural Dimantion; stress',
    { h: 'Introduction' },
    'The behavioural dimension matters. Stress has a behavioural dimension and a physiological dimension.',
    { h: 'References' }
  ];
  var r = TM.analyze(doc(lines), {});
  var kw = byId(r, /^kw:absent:/)[0];
  assert.ok(kw && kw.status === 'bad');
  assert.strictEqual(kw.fix.items[0].next, 'Behavioural Dimension');
});

test('laporan komentar: satu komentar per kemunculan yang perlu diubah', function () {
  var r = TM.analyze(doc(EN), {});
  var cs = TM.reportComments(r, {});
  assert.ok(cs.length > 3);
  assert.ok(cs.every(function (c) { return c.loc && c.loc.find && /^Istilah — /.test(c.text); }));
  assert.ok(!cs.some(function (c) { return c.status === 'info'; }), 'info tidak ikut kecuali diminta');
  var skip = TM.reportComments(r, { skip: function (f) { return /^spell:/.test(f.id); } });
  assert.ok(skip.length < cs.length && !skip.some(function (c) { return /^spell:/.test(c.id); }), 'temuan yang diabaikan tidak ikut');
});

test('prefiks komentar istilah terpisah dari laporan lain', function () {
  var RC = require('../report-core.js');
  var t = RC.TERM_PREFIX + ' Istilah — x';
  assert.ok(RC.commentMatches(t, 'term') && RC.commentMatches(t, 'all'));
  assert.ok(!RC.commentMatches(t, 'validator') && !RC.commentMatches(t, 'template'));
  assert.ok(!RC.commentMatches(RC.TEMPLATE_PREFIX + ' x', 'term'));
  assert.ok(/Konsistensi istilah: 2 perlu diperbaiki/.test(RC.summaryComment({ term: { bad: 2, warn: 1 } })));
});

test('contoh Nasreen: nama variabel & instrumen', function () {
  return sample('nasreen-98-112.docx').then(function (m) {
    if (!m) { console.log('        (lewati: fixture tidak ada)'); return; }
    var r = TM.analyze(m, {});
    var ids = r.findings.map(function (f) { return f.id; });
    var sst = byId(r, /^term:socialskilltraining$/)[0];
    assert.ok(sst && sst.kinds.indexOf('number') >= 0 && sst.kinds.indexOf('case') >= 0, ids.join(', '));
    assert.ok(byId(r, /^term:academicstress$/).length, 'academic Stress');
    assert.ok(byId(r, /^term:academicstresschecklistscale$/).length, 'check list vs checklist');
    assert.ok(byId(r, /^spell:behavior/).length && byId(r, /^term:pretest$/).length && byId(r, /^term:posttest$/).length);
    assert.ok(byId(r, /^acr:early:body:ASRS$/).length, 'ASRS dipakai sebelum didefinisikan');
    assert.ok(!byId(r, /^acr:twonames:/).length, 'ASCS & ASRS memang dua instrumen (disebut bersama)');
    var kw = byId(r, /^kw:absent:/)[0];
    assert.ok(kw && /Dimension/.test(kw.fix.items[0].next));
    assert.ok(!ids.some(function (i) { return /scoredecrease|provideevidence|findingsuggest/.test(i); }), 'frasa kata kerja tidak dilaporkan');
  });
});

test('naskah contoh lain: tidak banyak alarm palsu', function () {
  return Promise.all(['nep-74-96.docx', 'voice-211-226.docx'].map(sample)).then(function (ms) {
    ms.forEach(function (m) {
      if (!m) return;
      var r = TM.analyze(m, {});
      assert.ok(r.summary.bad <= 2 && r.summary.warn <= 6, JSON.stringify(r.summary) + ' ' + r.findings.map(function (f) { return f.id; }).join(', '));
    });
  });
});

(function next(i) {
  if (i >= queue.length) { console.log('\n' + passed + ' passed, ' + failed + ' failed.'); if (failed) process.exit(1); return; }
  var t = queue[i];
  Promise.resolve().then(t[1]).then(function () { passed++; console.log('  PASS  ' + t[0]); }, function (e) { failed++; console.log('  FAIL  ' + t[0] + '\n        ' + (e && e.stack || e)); }).then(function () { next(i + 1); });
})(0);
