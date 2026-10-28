// Tes Node untuk WordBridge.insertReport dengan tiruan Office.js. Tiruan bisa meniru dua perilaku
// Word desktop: (A) format yang diantrikan sebelum insertComment hilang, (B) nama warna
// sorotan ditolak diam-diam (hanya hex yang diterima).
//   node word-addin/test/word-bridge.test.js
var assert = require('assert');
var passed = 0, failed = 0, queue = [];
function test(name, fn) { queue.push([name, fn]); }

function fakeWord(texts, mode) {
  var doc = { hls: [], comments: [], track: 'Off' };
  function Range(p, start, len) {
    var r = { p: p, start: start, len: len, load: function () { return r; } };
    r.font = {};
    Object.defineProperty(r.font, 'highlightColor', {
      get: function () {
        var h = doc.hls.filter(function (x) { return x.p === p && x.start <= start && x.start + x.len >= start + len; }).pop();
        return h ? h.color : null;
      },
      set: function (v) {
        doc.hls = doc.hls.filter(function (x) { return !(x.p === p && x.start >= start && x.start + x.len <= start + len); });
        if (v == null) return;
        if (mode.names === false && !/^#/.test(v)) return; // (B)
        doc.hls.push({ p: p, start: start, len: len, color: v });
      }
    });
    r.insertComment = function (text) {
      if (mode.dropBeforeComment) doc.hls = doc.hls.filter(function (x) { return !(x.p === p && x.start < start + len && x.start + x.len > start); }); // (A)
      var c = { content: text, getRange: function () { return Range(p, start, len); }, load: function () { return c; }, delete: function () { doc.comments.splice(doc.comments.indexOf(c), 1); } };
      doc.comments.push(c); return c;
    };
    r.getComments = function () { var items = doc.comments.filter(function (c) { var g = c.getRange(); return g.p === p && g.start < start + len && g.start + g.len > start; }); return { items: items, load: function () {} }; };
    return r;
  }
  var paras = texts.map(function (t, i) {
    return { text: t, getRange: function () { return Range(i, 0, t.length); },
      search: function (q) { var items = [], k = -1; while ((k = t.indexOf(q, k + 1)) >= 0) items.push(Range(i, k, q.length)); return { items: items, load: function () {} }; } };
  });
  var ctx = { document: { load: function () {}, get changeTrackingMode() { return doc.track; }, set changeTrackingMode(v) { doc.track = v; },
    body: { paragraphs: { items: paras, load: function () {} }, getComments: function () { return { items: doc.comments.slice(), load: function () {} }; } } },
    sync: function () { return Promise.resolve(); } };
  global.window = global;
  global.Office = { context: { requirements: { isSetSupported: function () { return true; } } } };
  global.Word = { run: function (fn) { return fn(ctx); } };
  return doc;
}
function loadBridge() {
  delete require.cache[require.resolve('../word-bridge.js')];
  global.ReportCore = require('../report-core.js');
  require('../word-bridge.js');
  return global.WordBridge;
}
var TEXTS = ['Judul naskah', 'Isi (Lee & Park, 2021) dan lainnya.', 'References', 'Lee, H. (2009), "Judul", Jurnal, Vol. 1, pp. 1-2.'];
var ITEMS = [
  { paraIndex: 0, offset: 0, length: 12, text: 'Judul naskah', whole: true, color: 'Turquoise', comment: '[Cek Sitasi] Ringkasan' },
  { paraIndex: 1, offset: 5, length: 16, text: 'Lee & Park, 2021', occurrence: 0, color: 'Red', comment: '[Cek Sitasi] Perlu diperbaiki' },
  { paraIndex: 3, offset: 9, length: 4, text: '2009', occurrence: 0, color: 'Yellow', comment: '[Cek Sitasi · Metadata] Tahun' }
];
[['Word normal', {}], ['(A) sorotan hilang saat insertComment', { dropBeforeComment: true }], ['(B) nama warna ditolak, hex diterima', { names: false }], ['(A)+(B) sekaligus', { dropBeforeComment: true, names: false }]].forEach(function (x) {
  test(x[0] + ': setiap komentar tetap tersorot', function () {
    var doc = fakeWord(TEXTS, x[1]), WB = loadBridge();
    return WB.insertReport(ITEMS, { highlight: true, replace: 'all' }).then(function (r) {
      assert.strictEqual(r.added, 3); assert.strictEqual(r.wantHighlight, 3); assert.strictEqual(r.highlighted, 3, JSON.stringify(r));
      assert.strictEqual(doc.comments.length, 3);
      doc.comments.forEach(function (c) { assert.ok(c.getRange().font.highlightColor, 'tersorot: ' + c.content); });
      var red = doc.hls.find(function (h) { return h.p === 1; });
      assert.ok(/^(?:Red|#FF0000)$/.test(red.color));
    });
  });
});
test('laporan ulang menghapus sorotan & komentar lama (tidak dobel)', function () {
  var doc = fakeWord(TEXTS, { dropBeforeComment: true }), WB = loadBridge();
  return WB.insertReport(ITEMS, { highlight: true, replace: 'all' }).then(function () { return WB.insertReport(ITEMS.slice(1), { highlight: true, replace: 'all' }); }).then(function (r) {
    assert.strictEqual(r.removed, 3); assert.strictEqual(doc.comments.length, 2);
    assert.strictEqual(doc.hls.filter(function (h) { return h.p === 0; }).length, 0, 'sorotan judul lama dihapus');
  });
});
test('sorotan dimatikan → tidak ada warna', function () {
  var doc = fakeWord(TEXTS, {}), WB = loadBridge();
  return WB.insertReport(ITEMS, { highlight: false, replace: 'all' }).then(function (r) { assert.strictEqual(r.wantHighlight, 0); assert.strictEqual(doc.hls.length, 0); });
});

(function run(i) {
  if (i >= queue.length) { console.log('\n' + passed + ' passed, ' + failed + ' failed.'); process.exit(failed ? 1 : 0); }
  var t = queue[i];
  Promise.resolve().then(t[1]).then(function () { passed++; console.log('  PASS  ' + t[0]); }, function (e) { failed++; console.log('  FAIL  ' + t[0] + '\n        ' + (e && e.stack || e)); }).then(function () { run(i + 1); });
})(0);
