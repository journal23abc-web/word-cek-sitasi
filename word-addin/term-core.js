/*
 * Cek Konsistensi Istilah — mesin murni (tanpa Office.js, bisa diuji di Node).
 *
 * Masukan: model DocxReader (paragraf badan + info sel tabel) dan, bila ada, hasil
 * TemplateCore.analyze (judul, abstrak, kata kunci, heading). Keluaran: daftar temuan, masing-
 * masing dengan bentuk-bentuk yang dipakai, jumlah, lokasi tiap kemunculan, saran bentuk baku,
 * dan (bila aman) daftar penggantian untuk "Seragamkan".
 *
 * Pendekatan (berbasis aturan, deterministik — tidak menebak makna bebas):
 *  1. Teks dipecah per zona: judul, abstrak, kata kunci, isi, heading, sel tabel, keterangan
 *     gambar/tabel. Daftar pustaka & halaman depan (penulis, afiliasi, info artikel) diabaikan.
 *  2. Setiap frasa 1–5 kata diberi KUNCI KONSEP: huruf kecil, ejaan UK→US, tiap kata dijadikan
 *     tunggal, pemisah (spasi / tanda hubung / digabung) dihapus. "Social Skills Training",
 *     "social skill training", "pre-test", "pretest" → kunci yang sama dengan bentuk dasarnya.
 *  3. Frasa yang hampir selalu muncul di dalam frasa yang lebih panjang ("skill training" di dalam
 *     "social skill training") dilebur ke frasa panjangnya, sehingga yang dilaporkan adalah nama
 *     istilah yang utuh, bukan potongannya.
 *  4. Dalam satu kunci konsep, perbedaan yang DILAPORKAN hanya yang memang tidak konsisten:
 *       - tunggal/jamak pada kata penjelas ("social skill training" vs "social skills training");
 *         jamak pada kata inti (kata terakhir) dianggap tata bahasa, tidak dilaporkan;
 *       - digabung vs tanda hubung vs spasi ("pretest"/"pre-test"; spasi vs tanda hubung hanya
 *         untuk awalan terikat seperti pre-, post-, self-, non-, co-, multi-);
 *       - huruf kapital di tengah kalimat ("Social Skill Training" vs "social skill training",
 *         "academic Stress"). Awal kalimat, heading, judul, sel tabel, dan definisi singkatan
 *         tidak dihitung.
 *  5. Ejaan Inggris UK vs US dalam kata yang sama (behavioural/behavioral, organise/organize) —
 *     hanya bila KEDUA ejaan sama-sama dipakai.
 *  6. Singkatan: kepanjangan "Full Term (FT)" dicocokkan dengan huruf awal; dilaporkan bila satu
 *     singkatan punya beberapa kepanjangan, satu kepanjangan punya beberapa singkatan, dua
 *     singkatan berbeda untuk nama yang hampir sama (mis. ASRS vs ASCS), dipakai sebelum
 *     didefinisikan, atau tidak pernah dipakai.
 *  7. Sinonim dalam nama variabel: dari kamus kecil sinonim akademik (EN & ID) — mis. "academic
 *     stress"/"academic pressure", "job satisfaction"/"work satisfaction", "kinerja"/"performa".
 *     Dua frasa yang pernah muncul bersama dalam SATU kalimat dianggap konsep berbeda (penulis
 *     sengaja membedakan), jadi tidak dilaporkan.
 *  8. Kata kunci yang tidak muncul sama sekali di abstrak/isi (sering salah ketik atau istilahnya
 *     berganti).
 */
(function (root) {
  'use strict';

  var STOP = new Set((
    'a an the and or nor but if then than so as of in on at by to for from with without within into onto upon over under ' +
    'about above below between among amongst across through throughout during before after since until while whilst ' +
    'is are was were be been being am do does did done has have had having can could may might must shall should will would ' +
    'this that these those there here it its they them their theirs we us our ours you your he she his her him i me my ' +
    'which who whom whose what when where why how whether not no yes also only both each every either neither any some ' +
    'all most more less many much few several such other another same own very too just even still yet thus hence ' +
    'however therefore moreover furthermore although though because via per et al etc ie eg vs ' +
    'one two three four five first second third et al ' +
    // Bahasa Indonesia
    'yang dan di ke dari pada untuk dengan dalam terhadap ini itu atau juga sebagai oleh adalah akan tidak ada para serta ' +
    'bahwa karena secara lebih sangat telah sudah dapat bisa antara tersebut setiap sebuah suatu satu hal maka namun tetapi ' +
    'agar supaya jika bila apabila ketika saat sehingga hingga sampai bagi tentang yaitu yakni ialah merupakan menjadi ' +
    'kami kita mereka ia dia saya anda nya pun lah kah belum masih hanya saja semua seluruh beberapa banyak berbagai'
  ).split(/\s+/));
  // Awalan terikat: "self efficacy" vs "self-efficacy" memang tidak konsisten.
  var BOUND = new Set('pre post self non co multi inter intra sub socio psycho anti semi quasi meta pseudo cross counter micro macro mid cyber e well re un bi tri'.split(' '));
  // Singkatan umum yang tidak perlu didefinisikan.
  var KNOWN_ABBR = new Set((
    'USA US UK EU UN UNESCO UNICEF WHO OECD ASEAN NATO IMF GDP GNP COVID SARS HIV AIDS DNA RNA PCR IQ ICT IT AI PC TV ' +
    'PhD MSc MA BA BSc MBA MD CEO CFO HR ID PDF URL HTML DOI ISSN ISBN ORCID APA IEEE MLA ' +
    'SPSS AMOS SEM PLS ANOVA ANCOVA MANOVA SD SE CI OR RR df M N n p SMS GPS USB ' +
    'NGO NGOs SME SMEs MSME MSMEs UMKM PT CV TK SD SMP SMA SMK MTs MA MI PAUD DKI RI NKRI TNI POLRI DPR MPR BUMN ' +
    'AD BC CE BCE AM PM GMT UTC ISO SmartPLS WarpPLS NVivo LISREL STATA JASP MPlus Mplus SAS ChatGPT GPT LLM LLMs OpenAI'
  ).split(/\s+/));

  // ---------- KAMUS SINONIM ----------
  // Lema: huruf kecil, ejaan US, bentuk tunggal. Satu kelompok = kata yang sering dipertukarkan
  // penulis (atau AI) untuk konsep yang sama di dalam nama variabel.
  var SYN_GROUPS = [
    ['training', 'intervention', 'program', 'treatment'],
    ['scale', 'inventory', 'questionnaire', 'checklist', 'instrument', 'survey'],
    ['performance', 'achievement', 'attainment'],
    ['student', 'learner', 'pupil'],
    ['teacher', 'educator', 'instructor'],
    ['employee', 'worker', 'staff', 'personnel'],
    ['job', 'work', 'occupational'],
    ['satisfaction', 'contentment'],
    ['intention', 'intent'],
    ['usage', 'utilization', 'use'],
    ['customer', 'consumer', 'client'],
    ['skill', 'ability', 'competence', 'competency', 'capability', 'proficiency'],
    ['online', 'digital', 'virtual'],
    ['organizational', 'corporate'],
    ['academic', 'educational', 'scholastic'],
    ['stress', 'pressure', 'strain'],
    ['anxiety', 'apprehension'],
    ['thinking', 'reasoning'],
    ['disposition', 'tendency', 'inclination'],
    ['instruction', 'teaching'],
    ['model', 'framework'],
    ['participant', 'respondent'],
    ['engagement', 'involvement'],
    ['trust', 'confidence'],
    ['purchase', 'buying'],
    ['wellbeing', 'welfare'],
    // Bahasa Indonesia
    ['pengaruh', 'dampak', 'efek'],
    ['kinerja', 'performa'],
    ['prestasi', 'capaian'],
    ['karyawan', 'pegawai', 'pekerja', 'staf'],
    ['siswa', 'murid', 'pelajar'],
    ['guru', 'pendidik', 'pengajar'],
    ['pelatihan', 'training'],
    ['minat', 'niat', 'intensi'],
    ['kepercayaan', 'trust'],
    ['pelanggan', 'konsumen'],
    ['kualitas', 'mutu'],
    ['layanan', 'pelayanan'],
    ['kemampuan', 'keterampilan', 'kompetensi', 'kecakapan', 'keahlian'],
    ['motivasi', 'dorongan'],
    ['stres', 'tekanan'],
    ['akademik', 'akademis'],
    ['daring', 'online'],
    ['pembelajaran', 'pengajaran'],
    ['instrumen', 'angket', 'kuesioner'],
    ['responden', 'partisipan']
  ];
  var SYN = new Map();
  SYN_GROUPS.forEach(function (g) {
    g.forEach(function (w) { var s = SYN.get(w) || new Set(); g.forEach(function (x) { if (x !== w) s.add(x); }); SYN.set(w, s); });
  });
  // Pasangan lawan kata: dua nama instrumen yang hanya berbeda di kata ini memang berbeda.
  var CONTRAST = [['positive', 'negative'], ['pre', 'post'], ['pretest', 'posttest'], ['internal', 'external'], ['high', 'low'], ['upper', 'lower'],
    ['short', 'long'], ['intrinsic', 'extrinsic'], ['direct', 'indirect'], ['male', 'female'], ['experimental', 'control'], ['primary', 'secondary'],
    ['independent', 'dependent'], ['first', 'second'], ['total', 'partial'], ['physical', 'mental'], ['verbal', 'nonverbal'], ['online', 'offline'],
    ['formal', 'informal'], ['public', 'private'], ['state', 'trait'], ['explicit', 'implicit'], ['individual', 'collective'], ['personal', 'social'],
    ['self', 'other'], ['mother', 'father'], ['before', 'after'], ['awal', 'akhir'], ['positif', 'negatif'], ['internal', 'eksternal']];

  // ---------- EJAAN UK → US ----------
  var US_MAP = {
    programme: 'program', programmes: 'programs', judgement: 'judgment', judgements: 'judgments', acknowledgement: 'acknowledgment', acknowledgements: 'acknowledgments',
    ageing: 'aging', grey: 'gray', enrolment: 'enrollment', enrolments: 'enrollments', enrol: 'enroll', enrols: 'enrolls', fulfil: 'fulfill', fulfils: 'fulfills',
    fulfilment: 'fulfillment', skilful: 'skillful', skilfully: 'skillfully', wilful: 'willful', instalment: 'installment', instalments: 'installments',
    catalogue: 'catalog', catalogues: 'catalogs', catalogued: 'cataloged', analogue: 'analog', analogues: 'analogs', defence: 'defense', defences: 'defenses',
    offence: 'offense', offences: 'offenses', pretence: 'pretense', sceptical: 'skeptical', scepticism: 'skepticism', sceptic: 'skeptic', sceptics: 'skeptics',
    mould: 'mold', moulds: 'molds', aluminium: 'aluminum', manoeuvre: 'maneuver', manoeuvres: 'maneuvers', artefact: 'artifact', artefacts: 'artifacts',
    jewellery: 'jewelry', draught: 'draft', plough: 'plow', cosy: 'cozy', moustache: 'mustache', pyjamas: 'pajamas', sulphur: 'sulfur', tonne: 'ton', tonnes: 'tons',
    diarrhoea: 'diarrhea', oedema: 'edema', practise: 'practise' // "practise" (kata kerja UK) ≠ "practice" (kata benda) — sengaja tidak dipetakan
  };
  var OUR_RE = /^(behavi|misbehavi|col|discol|fav|unfav|lab|hon|dishon|harb|neighb|rum|hum|vap|endeav|od|vig|rig|ard|clam|savi|flav|tum|parl|arm|splend|cand)our(s|ed|ing|ite|ites|able|ably|al|ally|less|ful|er|ers|ism|ist|ists|hood|hoods|y)?$/;
  var ISE_EXC = ['advise', 'arise', 'chastise', 'circumcise', 'comprise', 'compromise', 'concise', 'demise', 'despise', 'devise', 'disguise', 'enterprise', 'excise',
    'exercise', 'expertise', 'franchise', 'improvise', 'incise', 'merchandise', 'precise', 'premise', 'promise', 'raise', 'revise', 'rise', 'supervise', 'surmise',
    'surprise', 'televise', 'treatise', 'wise', 'paradise', 'noise', 'poise', 'praise', 'cruise', 'bruise', 'anise', 'porpoise', 'tortoise', 'turquoise', 'mortise',
    'apprise', 'reprise', 'valise', 'guise', 'vise', 'chemise', 'cerise', 'practise', 'licence', 'reise', 'expertise'];
  var LL_RE = /^(counsel|model|label|travel|cancel|signal|level|fuel|total|channel|tunnel|dial|marvel|pedal|quarrel|rival|shovel|funnel|jewel|equal|initial|duel|yodel|unravel|tranquil|libel|panel|pencil|spiral|snorkel|swivel|bevel|grovel|kennel|tassel|towel|chisel|enamel|gravel|label|parcel|program)l(ing|ed|er|ers|or|ors|ist|ists|ity)$/;
  var TRE_RE = /^(cent|epicent|fib|theat|lit|spect|somb|calib|lust|meag|sab)re(s|d)?$/;
  var AE_RULES = [[/^paed/, 'ped'], [/^orthopaed/, 'orthoped'], [/^anaem/, 'anem'], [/^haem/, 'hem'], [/^leukaem/, 'leukem'], [/^anaesth/, 'anesth'],
    [/^oestr/, 'estr'], [/^foet/, 'fet'], [/^oesoph/, 'esoph'], [/^gynaec/, 'gynec'], [/^encyclopaed/, 'encycloped'], [/^caesar/, 'cesar']];

  function usKey(w) {
    if (Object.prototype.hasOwnProperty.call(US_MAP, w)) return US_MAP[w];
    var m;
    if ((m = OUR_RE.exec(w))) return m[1] + 'or' + (m[2] || '');
    if ((m = LL_RE.exec(w))) return m[1] + m[2];
    if ((m = TRE_RE.exec(w))) return m[1] + 'er' + (m[2] === 'd' ? 'ed' : m[2] || '');
    // -ise / -isation / -yse (kata kerja yang di AS ditulis -ize / -yze)
    if ((m = /^(\p{L}{3,}?)is(e|es|ed|ing|ation|ations|ational|er|ers|able)$/u.exec(w))) {
      var base = m[1] + 'ise';
      if (!ISE_EXC.some(function (x) { return base === x || base.slice(-x.length) === x; }) && !/[aeiou]$/.test(m[1])) return m[1] + 'iz' + m[2];
    }
    if ((m = /^(ana|para|cata|dia|electro|hydro)lys(e|ed|ing|er|ers)$/.exec(w))) return m[1] + 'lyz' + m[2];
    for (var i = 0; i < AE_RULES.length; i++) if (AE_RULES[i][0].test(w)) return w.replace(AE_RULES[i][0], AE_RULES[i][1]);
    return w;
  }

  // ---------- TUNGGAL / JAMAK (Inggris) ----------
  var IRREG = { children: 'child', women: 'woman', men: 'man', criteria: 'criterion', phenomena: 'phenomenon', indices: 'index', matrices: 'matrix', appendices: 'appendix',
    analyses: 'analysis', hypotheses: 'hypothesis', theses: 'thesis', crises: 'crisis', diagnoses: 'diagnosis', syntheses: 'synthesis', emphases: 'emphasis',
    parentheses: 'parenthesis', bases: 'basis', feet: 'foot', teeth: 'tooth', mice: 'mouse', curricula: 'curriculum', stimuli: 'stimulus', loci: 'locus' };
  var INVAR = new Set('bias atlas canvas gas pancreas overseas whereas alias lens news series species means mathematics physics economics politics ethics statistics ' +
    'logistics linguistics semantics dynamics genetics electronics athletics gymnastics aesthetics ergonomics demographics analytics thanks always perhaps across ' +
    'status campus corpus virus focus bonus census consensus apparatus stimulus nexus thesis analysis basis crisis axis emphasis synthesis diagnosis hypothesis ' +
    'process access success stress class mass glass less unless progress congress address business witness illness awareness wellness fitness kindness ' +
    'data media criteria phenomena chaos ethos pathos kudos cosmos diabetes herpes rabies measles mumps scissors trousers'.split(' '));
  function singular(w) {
    if (w.length <= 3 || INVAR.has(w)) return w;
    if (IRREG[w]) return IRREG[w];
    if (/(ss|us|is|ics|ous|[^s]ness)$/.test(w)) return w;
    if (/[^aeiou]ies$/.test(w) && w.length > 4) return w.slice(0, -3) + 'y';
    if (/(sses|xes|zes|ches|shes)$/.test(w)) return w.slice(0, -2);
    if (/s$/.test(w) && !/ss$/.test(w)) return w.slice(0, -1);
    return w;
  }
  function pluralize(w) {
    if (/(s|x|z|ch|sh)$/i.test(w)) return w + 'es';
    if (/[^aeiou]y$/i.test(w)) return w.slice(0, -1) + 'ies';
    return w + 's';
  }

  // ---------- TOKENISASI ----------
  var WORD_RE = /[\p{L}\p{M}\p{N}]+(?:['’][\p{L}]+)*/gu;
  // Singkatan: ≥2 huruf kapital, diakhiri huruf kapital/angka (boleh + "s"): SST, ASRS, GenAI, PhD, COVID-19 → "COVID".
  function isAcr(w) {
    if (w.length < 2 || w.length > 10 || !/^\p{Lu}/u.test(w)) return false;
    var caps = (w.match(/\p{Lu}/gu) || []).length;
    if (caps < 2) return false;
    return /^\p{Lu}[\p{Lu}\p{N}]*s?$/u.test(w) || /^\p{Lu}\p{Ll}{1,3}\p{Lu}[\p{Lu}\p{N}]*s?$/u.test(w);
  }
  function caseOf(w) { return /^\p{Lu}+$/u.test(w) ? 'U' : /^\p{Lu}/u.test(w) ? 'T' : /\p{Lu}/u.test(w) ? 'M' : 'L'; }

  function tokenize(text, base, lang) {
    var out = [], m, prevEnd = 0, sent = 0;
    WORD_RE.lastIndex = 0;
    while ((m = WORD_RE.exec(text))) {
      var w = m[0], sep = text.slice(prevEnd, m.index), first = !out.length;
      var joint = first ? 'brk' : /^[-‑]$/.test(sep) || /^ [-‑] $/.test(sep) ? 'hy' : /^[  ]{1,2}$/.test(sep) ? 'sp' : 'brk';
      var start = first ? /^[\s"“‘(\[•\-–—*\d.)]*$/.test(sep) : /[.!?:]["”’)\]]*\s+["“‘(\[]*$/.test(sep) || /[\n\t•]/.test(sep);
      if (start && !first) sent++;
      var l = w.toLocaleLowerCase('en'), poss = /['’]s$/i.test(w);
      var core = poss ? l.replace(/['’]s$/, '') : l;
      var us = usKey(core);
      out.push({ w: w, l: l, us: us, k: lang === 'id' ? us : singular(us), s: base + m.index, e: base + m.index + w.length, joint: joint, start: start, sent: sent,
        poss: poss || /^['’]/.test(text.slice(m.index + w.length, m.index + w.length + 1)) && /s$/i.test(w),
        alpha: /^[\p{L}\p{M}]+(?:['’]\p{L}+)?$/u.test(w), acr: isAcr(w), cs: caseOf(w) });
      prevEnd = m.index + w.length;
    }
    return out;
  }
  // Kata kerja umum di tulisan ilmiah: frasa yang diawali/diakhiri kata kerja bukan nama istilah
  // ("scores decreased", "provides evidence", "findings suggest").
  var VERBS = new Set(('suggest show indicate reveal demonstrate provide increase decrease improve reduce affect influence predict explain support find found ' +
    'report present include contribute lead led result remain appear seem tend help enable allow require make made take took give gave become became ' +
    'confirm conclude examine investigate assess measure determine identify obtain achieve experience exhibit display highlight emphasize recommend ' +
    'mean imply raise lower decline drop rise grow differ vary exceed reach yield produce cause enhance promote facilitate hinder prevent ensure ' +
    'is are was were has have had do does did can could may might will would should must').split(' '));
  var NOT_PAST = new Set('need speed seed feed breed creed greed bleed weed deed reed bed red shed hundred sacred kindred embed bred wed'.split(' '));
  function verbish(t, atEnd) {
    var l = t.l, b = l.replace(/(es|s|ed|d)$/, '');
    if (VERBS.has(l) || VERBS.has(b) || VERBS.has(l.replace(/s$/, '')) || VERBS.has(l.replace(/ed$/, '')) || VERBS.has(l.replace(/d$/, '')) || VERBS.has(l.replace(/ied$/, 'y')) || VERBS.has(l.replace(/ies$/, 'y'))) return true;
    return !!atEnd && /ed$/.test(l) && l.length > 4 && !NOT_PAST.has(l);
  }
  function edgeOk(t) { return content(t) && !verbish(t); }
  function content(t) { return t.alpha && !t.poss && !STOP.has(t.l) && (t.l.length >= 2) && !/^[ivxlcdm]+$/.test(t.l) || t.alpha && t.l.length === 1 && false; }

  // ---------- ZONA DOKUMEN ----------
  var REF_RE = /^\s*(references?|reference\s+list|bibliography|daftar\s+pustaka|referensi|rujukan|works\s+cited|literature\s+cited|sumber\s+rujukan)\s*:?\s*$/i;
  var CAPTION_RE = /^\s*(table|tabel|figure|fig\.|gambar|chart|grafik|diagram|bagan|lampiran|appendix)\s*[\dIVX]+/i;
  var END_RE = /^\s*(acknowledge?ments?|ucapan\s+terima\s+kasih|author\s+contributions?|conflicts?\s+of\s+interest|funding|declarations?|data\s+availability|competing\s+interests?|ethical\s+(approval|statement))\b/i;

  function layout(model, tpl) {
    var body = model.body || [];
    var TC = root.TemplateCore || (typeof require === 'function' ? safeRequire('./template-core.js') : null);
    if (!tpl && TC) tpl = TC.analyze(model, {});
    tpl = tpl || {};
    var secs = tpl.sections || [];
    var heads = new Set(secs.map(function (s) { return s.index; }));
    var refIdx = body.length;
    var ref = secs.filter(function (s) { return s.key === 'references'; })[0];
    if (ref) refIdx = ref.index;
    else for (var i = body.length - 1; i > body.length * 0.3; i--) if (REF_RE.test(body[i].text)) { refIdx = i; break; }
    var segs = [];
    var absIdx = new Set();
    if (tpl.title) segs.push({ index: tpl.title.index, base: 0, text: body[tpl.title.index].text, zone: 'title' });
    if (tpl.abstract) tpl.abstract.parts.forEach(function (p) { absIdx.add(p.index); segs.push({ index: p.index, base: p.offset || 0, text: body[p.index].text.slice(p.offset || 0), zone: 'abstract' }); });
    var kwIdx = new Set();
    if (tpl.keywords) {
      kwIdx.add(tpl.keywords.heading.index);
      tpl.keywords.items.forEach(function (it) { kwIdx.add(it.index); segs.push({ index: it.index, base: it.offset, text: it.text, zone: 'keywords', keyword: it.text }); });
    }
    var start = secs.length ? secs[0].index : Math.max.apply(null, [0].concat(Array.from(absIdx), Array.from(kwIdx)).map(function (x) { return x + 1; }));
    var endMatter = false;
    for (var j = start; j < refIdx; j++) {
      var p = body[j], t = p.text;
      if (!t || !t.trim() || absIdx.has(j) || kwIdx.has(j) || (tpl.title && j === tpl.title.index)) continue;
      if (heads.has(j) && END_RE.test(t)) endMatter = true; else if (heads.has(j)) endMatter = false;
      if (endMatter) continue;
      var zone = heads.has(j) || looksHeading(p) ? 'heading' : p.cell ? 'cell' : CAPTION_RE.test(t) ? 'caption' : 'body';
      segs.push({ index: j, base: 0, text: t, zone: zone });
    }
    return { segs: segs, refIndex: refIdx, tpl: tpl };
  }
  function looksHeading(p) {
    var t = (p.text || '').trim();
    if (!t || p.cell || t.length > 120 || /[.:;,]$/.test(t)) return false;
    var words = t.split(/\s+/).length;
    if (words > 12) return false;
    if (p.bold || /^heading|^judul/i.test(p.style || '') || /^(\d+(\.\d+)*\.?|[IVX]+\.|[A-H]\.)\s+\p{Lu}/u.test(t)) return true;
    // Subjudul bergaya "Title Case" tanpa titik di akhir
    var cw = t.split(/\s+/).filter(function (w) { return /^\p{L}/u.test(w) && !STOP.has(w.toLocaleLowerCase('en')); });
    return words >= 2 && cw.length >= 2 && cw.filter(function (w) { return /^\p{Lu}/u.test(w); }).length / cw.length >= 0.75;
  }
  function safeRequire(p) { try { return require(p); } catch (e) { return null; } }

  function detectLang(segs) {
    var en = 0, id = 0;
    segs.slice(0, 400).forEach(function (s) {
      (s.text.toLowerCase().match(/\b(the|and|of|is|are|with|this|that)\b/g) || []).forEach(function () { en++; });
      (s.text.toLowerCase().match(/\b(yang|dan|dengan|pada|untuk|dalam|adalah|ini)\b/g) || []).forEach(function () { id++; });
    });
    return id > en ? 'id' : 'en';
  }

  // ---------- ANALISIS UTAMA ----------
  function analyze(model, opts) {
    var o = opts || {};
    var lay = layout(model, o.template);
    var lang = o.lang || detectLang(lay.segs);
    var body = model.body || [];
    var segs = lay.segs.map(function (s, si) { var toks = tokenize(s.text, s.base, lang); toks.forEach(function (t) { t.seg = si; }); return Object.assign({}, s, { toks: toks }); });

    // Definisi singkatan dulu: kemunculan bentuk panjang di dalam definisi tidak ikut cek huruf kapital.
    var acr = findAcronyms(segs, body);
    var defSpans = acr.defs.map(function (d) { return { seg: d.seg, s: d.fullStart, e: d.fullEnd }; });
    function inDef(seg, s, e) { return defSpans.some(function (d) { return d.seg === seg && s >= d.s && e <= d.e; }); }

    // ---- indeks frasa ----
    var fam = new Map(); // key → { key, occ: [] }
    segs.forEach(function (sg, si) {
      var T = sg.toks;
      for (var i = 0; i < T.length; i++) {
        if (!content(T[i])) continue;
        var startVerb = verbish(T[i]);
        for (var n = 1; n <= 5 && i + n <= T.length; n++) {
          var j = i + n - 1;
          if (n > 1) {
            if (T[j].joint === 'brk' || T[j].poss && j < i + n - 1) break;
            if (!T[j].alpha || T[j - 1].poss) break;
            if (STOP.has(T[j].l) && T[j].l !== 'of' && T[j].l !== 'and') break;
          }
          if (!content(T[j])) continue;
          if (n > 1 && (startVerb || verbish(T[j], true))) continue;
          var toks = T.slice(i, j + 1);
          if (toks.some(function (t) { return /\d/.test(t.w); })) break;
          if (toks.filter(function (t) { return t.l === 'and'; }).length) continue; // "and" hanya boleh sebagai jembatan, frasa tidak disimpan
          var key = toks.map(function (t) { return t.k; }).join('');
          var f = fam.get(key);
          if (!f) { f = { key: key, occ: [] }; fam.set(key, f); }
          f.occ.push({ seg: si, ti: i, tn: n, index: sg.index, zone: sg.zone, s: T[i].s, e: T[j].e, mid: !T[i].start, sent: si + ':' + T[i].sent,
            def: inDef(si, T[i].s, T[j].e), toks: toks });
        }
      }
    });

    // ---- frasa yang terserap ke frasa lebih panjang ----
    var bySeg = new Map();
    fam.forEach(function (f) {
      if (f.occ.length < 2) return;
      f.occ.forEach(function (oc) { var l = bySeg.get(oc.seg); if (!l) bySeg.set(oc.seg, l = []); l.push({ f: f, oc: oc }); });
    });
    fam.forEach(function (f) {
      if (f.occ.length < 2) return;
      var tally = new Map();
      f.occ.forEach(function (oc) {
        var seen = new Set();
        (bySeg.get(oc.seg) || []).forEach(function (x) {
          if (x.f === f || seen.has(x.f)) return;
          if (x.oc.ti <= oc.ti && x.oc.ti + x.oc.tn >= oc.ti + oc.tn && x.oc.tn > oc.tn) { seen.add(x.f); tally.set(x.f, (tally.get(x.f) || 0) + 1); }
        });
      });
      tally.forEach(function (n, g) { if (n >= 0.75 * f.occ.length && g.occ.length >= 2) f.nested = g; });
    });

    var anchorKeys = new Set();
    acr.defs.forEach(function (d) { anchorKeys.add(d.fullKey); });
    fam.forEach(function (f) { if (f.occ.some(function (oc) { return oc.zone === 'keywords' && wholeKeyword(segs[oc.seg], oc); })) anchorKeys.add(f.key); });
    fam.forEach(function (f) { if (f.occ.some(function (oc) { return oc.zone === 'title'; }) && f.occ[0].tn > 1 && !f.nested) f.inTitle = true; });

    var findings = [];
    var claimed = []; // rentang kemunculan yang sudah dilaporkan (hindari laporan ganda)

    // ---- 1. variasi penulisan istilah ----
    var cands = [];
    fam.forEach(function (f) {
      if (f.occ.length < 2 || f.nested) return;
      var maxN = Math.max.apply(null, f.occ.map(function (x) { return x.tn; }));
      var anchor = anchorKeys.has(f.key) || !!f.inTitle;
      var r = variation(f, lang, anchor, segs);
      if (!r) return;
      r.anchor = anchor;
      r.maxN = maxN;
      cands.push(r);
    });
    // Frasa pendek dulu: bila perbedaan pada frasa panjang ("social skills training intervention")
    // sudah dijelaskan oleh frasa pendek yang dilaporkan ("social skills training"), frasa panjang
    // tidak dilaporkan lagi. Sebaliknya frasa pendek yang hanya potongan sudah dilebur di atas.
    cands.sort(function (a, b) { return (a.maxN - b.maxN) || (b.anchor - a.anchor) || (b.f.occ.length - a.f.occ.length); });
    cands.forEach(function (r) {
      var explained = r.changes.filter(function (oc) {
        return claimed.some(function (c) { return c.seg === oc.seg && c.s < oc.e && oc.s < c.e && (c.s >= oc.s && c.e <= oc.e || oc.s >= c.s && oc.e <= c.e) && c.kinds.some(function (k) { return r.kinds.has(k); }); });
      }).length;
      if (explained >= r.changes.length * 0.6) return;
      var kinds = Array.from(r.kinds);
      r.f.occ.forEach(function (oc) { claimed.push({ seg: oc.seg, s: oc.s, e: oc.e, kinds: kinds }); });
      findings.push(termFinding(r, segs, body, lang));
    });

    // ---- 2. ejaan UK/US ----
    spellingFindings(segs, body).forEach(function (x) { findings.push(x); });

    // ---- 3. singkatan ----
    acronymFindings(acr, segs, body, fam, lang).forEach(function (x) { findings.push(x); });

    // ---- 4. sinonim dalam nama variabel ----
    synonymFindings(fam, anchorKeys, segs, body, lang).forEach(function (x) { findings.push(x); });

    // ---- 4b. salah ketik di dalam nama istilah ----
    typoFindings(fam, anchorKeys, segs, body, lang).forEach(function (x) { findings.push(x); });

    // ---- 5. kata kunci yang tidak muncul di naskah ----
    keywordFindings(segs, fam, lang).forEach(function (x) { findings.push(x); });

    var ORDER = { bad: 0, warn: 1, info: 2 }, GORDER = { term: 0, typo: 1, acronym: 2, synonym: 3, spelling: 4, compound: 5, keyword: 6 };
    findings.sort(function (a, b) { return (ORDER[a.status] - ORDER[b.status]) || (GORDER[a.group] - GORDER[b.group]) || (b.total - a.total); });
    var seenId = new Set();
    findings = findings.filter(function (f) { if (seenId.has(f.id)) return false; seenId.add(f.id); return true; });
    return {
      lang: lang, findings: findings,
      acronyms: acr.defs.map(function (d) { return { abbr: d.abbr, full: d.full, index: d.index }; }),
      stats: { segments: segs.length, terms: Array.from(fam.values()).filter(function (f) { return f.occ.length >= 3 && !f.nested && f.occ[0].tn > 1; }).length },
      summary: { bad: findings.filter(function (f) { return f.status === 'bad'; }).length, warn: findings.filter(function (f) { return f.status === 'warn'; }).length, info: findings.filter(function (f) { return f.status === 'info'; }).length }
    };
  }

  function wholeKeyword(seg, oc) {
    if (seg.zone !== 'keywords') return true;
    return oc.ti === 0 && oc.ti + oc.tn === seg.toks.length;
  }

  // Tanda tangan bentuk: tunggal/jamak kata penjelas + pemisah (digabung/hubung/spasi).
  function signature(oc, lang) {
    var parts = [];
    oc.toks.forEach(function (t, i) {
      var last = i === oc.toks.length - 1 || oc.toks[i + 1].l === 'of';
      var word = last || lang === 'id' ? t.k : t.us;
      if (i > 0) {
        var prev = oc.toks[i - 1], j = t.joint;
        if (j === 'hy' && !BOUND.has(prev.k) && !(prev.l.length === 1)) j = 'sp';
        parts.push(j === 'hy' ? '-' : ' ');
      }
      parts.push(word);
    });
    return parts.join('');
  }
  // Kelas huruf satu kemunculan: 'lower' | 'title' | 'mixed' (huruf kecil lalu kapital, mis.
  // "academic Stress") | null (tidak dinilai: pola "Kapital kecil" biasanya awal kalimat yang
  // tidak terdeteksi, atau bagian dari nama kapital yang lebih panjang).
  function caseClass(oc, segs) {
    var ws = oc.toks.filter(function (t) { return !STOP.has(t.l) && !t.acr; });
    if (!ws.length) return null;
    var cs = ws.map(function (t) { return t.cs === 'U' && t.w.length > 1 ? 'T' : t.cs === 'M' ? 'T' : t.cs; });
    if (cs.every(function (c) { return c === 'L'; })) return 'lower';
    if (cs.every(function (c) { return c === 'T'; })) {
      if (segs) {
        var T = segs[oc.seg].toks, a = T[oc.ti - 1], b = T[oc.ti + oc.tn];
        var cap = function (t, j) { return t && j !== 'brk' && t.alpha && !STOP.has(t.l) && (t.cs === 'T' || t.cs === 'U' && !t.acr); };
        if (b && cap(b, b.joint) || a && oc.mid && cap(a, T[oc.ti].joint)) return null; // bagian dari nama yang lebih panjang
      }
      return 'title';
    }
    for (var i = 1; i < cs.length; i++) if (cs[i - 1] === 'L' && cs[i] === 'T') return 'mixed';
    return null;
  }
  function caseCounted(oc) { return oc.mid && !oc.def && (oc.zone === 'body' || oc.zone === 'abstract'); }

  function variation(f, lang, anchor, segs) {
    var bySig = new Map();
    f.occ.forEach(function (oc) { var s = signature(oc, lang); oc.sig = s; bySig.set(s, (bySig.get(s) || 0) + 1); });
    var sigs = Array.from(bySig.keys());
    var multi = f.occ.some(function (oc) { return oc.tn > 1; });
    var kinds = new Set(), ignored = new Set();
    var domSig = null;
    if (sigs.length > 1) {
      domSig = pickDominant(sigs, bySig, f);
      var domOcc = f.occ.filter(function (oc) { return oc.sig === domSig; })[0];
      sigs.forEach(function (s) {
        if (s === domSig) return;
        var a = s.replace(/[- ]/g, ''), b = domSig.replace(/[- ]/g, '');
        if (a === b) { kinds.add('compound'); return; }
        // Pola subjek–predikat ("findings suggest" / "finding suggests"): jamak berpindah dari kata
        // pertama ke kata terakhir → itu kesesuaian tata bahasa, bukan nama istilah.
        var oOcc = f.occ.filter(function (oc) { return oc.sig === s; })[0];
        if (lang !== 'id' && oOcc.tn === domOcc.tn && agreement(domOcc, oOcc)) { ignored.add(s); return; }
        kinds.add('number');
      });
    }
    // huruf kapital
    var ccount = { lower: 0, title: 0, mixed: 0 };
    var counted = f.occ.filter(caseCounted);
    counted.forEach(function (oc) { var c = caseClass(oc, segs); oc.cc = c; if (c) ccount[c]++; });
    var domCase = null;
    var caseIssue = false;
    if (multi || anchor) {
      caseIssue = ccount.lower > 0 && ccount.title > 0 || ccount.mixed > 0 && ccount.lower + ccount.title > 0;
      if (caseIssue) { domCase = ccount.lower >= ccount.title ? 'lower' : 'title'; kinds.add('case'); }
    }
    if (!kinds.size) return null;
    if (!domSig) domSig = sigs[0];
    // Beda tunggal/jamak saja pada istilah biasa yang jarang muncul (≤2×) — terlalu lemah untuk dilaporkan.
    if (kinds.size === 1 && kinds.has('number') && !anchor && f.occ.filter(function (oc) { return !ignored.has(oc.sig); }).length < 3) return null;
    // daftar kemunculan yang perlu diubah
    var changes = f.occ.filter(function (oc) {
      return oc.sig !== domSig && !ignored.has(oc.sig) || caseIssue && caseCounted(oc) && oc.cc && oc.cc !== domCase;
    });
    if (!changes.length) return null;
    return { f: f, kinds: kinds, bySig: bySig, domSig: domSig, domCase: domCase, ccount: ccount, changes: changes, caseIssue: caseIssue };
  }
  function plural(t) { return singular(t.us) !== t.us; }
  function agreement(a, b) {
    var n = a.toks.length, last = n - 1;
    var headFlip = plural(a.toks[last]) !== plural(b.toks[last]);
    if (!headFlip) return false;
    for (var i = 0; i < last; i++) if (plural(a.toks[i]) !== plural(b.toks[i]) && plural(a.toks[i]) !== plural(a.toks[last])) return true;
    return false;
  }
  function pickDominant(sigs, bySig, f) {
    var best = sigs[0];
    sigs.forEach(function (s) {
      var a = bySig.get(s), b = bySig.get(best);
      if (a > b) best = s;
      else if (a === b) {
        var inA = f.occ.some(function (oc) { return oc.sig === s && (oc.zone === 'title' || oc.zone === 'keywords' || oc.def); });
        var inB = f.occ.some(function (oc) { return oc.sig === best && (oc.zone === 'title' || oc.zone === 'keywords' || oc.def); });
        if (inA && !inB) best = s;
      }
    });
    return best;
  }

  function occText(segs, oc) { var sg = segs[oc.seg]; return sg.text.slice(oc.s - sg.base, oc.e - sg.base); }
  function loc(segs, body, oc) {
    var sg = segs[oc.seg];
    return { index: sg.index, offset: oc.s, length: oc.e - oc.s, text: occText(segs, oc), zone: sg.zone };
  }

  // Bentuk pengganti untuk satu kemunculan: ikuti bentuk baku, tetapi pertahankan jamak/tunggal kata
  // terakhir (tata bahasa kalimatnya) dan huruf kapital sesuai posisi (awal kalimat/heading/judul).
  function rebuild(templ, oc, opts) {
    var orig = oc.toks, lang = opts.lang;
    var words = templ.words.slice(), seps = templ.seps.slice();
    // Jamak/tunggal kata inti (kata terakhir, atau kata sebelum "of") mengikuti kalimat aslinya.
    if (lang !== 'id' && words.length) {
      var pairs = words.length === orig.length
        ? words.map(function (w, i) { return i; }).filter(function (i) { return i === words.length - 1 || words[i + 1].toLocaleLowerCase('en') === 'of'; }).map(function (i) { return [i, i]; })
        : [[words.length - 1, orig.length - 1]];
      pairs.forEach(function (pq) {
        var li = pq[0], o = orig[pq[1]], lw = words[li].toLocaleLowerCase('en');
        var origPlural = singular(o.us) !== o.us, tplPlural = singular(usKey(lw)) !== usKey(lw);
        if (origPlural && !tplPlural) words[li] = pluralize(words[li]);
        else if (!origPlural && tplPlural) { var sg = singular(lw); words[li] = /^\p{Lu}/u.test(words[li]) ? cap(sg) : sg; }
      });
    }
    // huruf kapital
    var keepCase = !!opts.keepCase || !oc.mid || oc.def || oc.zone === 'heading' || oc.zone === 'title' || oc.zone === 'cell' || oc.zone === 'caption' || oc.zone === 'keywords';
    var cc0 = oc.cc !== undefined ? oc.cc : caseClass(oc);
    if (cc0 === null) keepCase = true;
    var cls = keepCase ? rawClass(oc) : (opts.domCase || rawClass(oc));
    var allUpper = orig.every(function (t) { return t.cs === 'U' || !/\p{L}/u.test(t.w); }) && orig.some(function (t) { return t.w.length > 2 && t.cs === 'U'; });
    words = words.map(function (w, i) {
      if (isAcr(w)) return w;
      if (allUpper) return w.toLocaleUpperCase('en');
      var lw = w.toLocaleLowerCase('en');
      if (cls === 'title' && !(i > 0 && STOP.has(lw))) return cap(lw);
      return lw;
    });
    if (!allUpper && !oc.mid && words.length) words[0] = cap(words[0]);
    var out = '';
    words.forEach(function (w, i) { out += (i ? seps[i] : '') + w; });
    return out;
  }
  function rawClass(oc) {
    var ws = oc.toks.filter(function (t) { return !STOP.has(t.l) && !t.acr; });
    return ws.length && ws.every(function (t) { return t.cs !== 'L'; }) ? 'title' : 'lower';
  }
  function cap(w) { return w ? w.charAt(0).toLocaleUpperCase('en') + w.slice(1) : w; }
  // Templat bentuk baku dari satu kemunculan; ejaan tiap kata diambil yang paling sering di
  // antara kemunculan bertanda tangan sama (supaya tidak ikut ejaan minoritas).
  function templateFrom(oc, pool) {
    var words = oc.toks.map(function (t, i) {
      if (!pool || !pool.length) return t.w;
      var tally = new Map();
      pool.forEach(function (p) { if (p.toks.length === oc.toks.length) { var l = p.toks[i].l; tally.set(l, (tally.get(l) || 0) + 1); } });
      var best = Array.from(tally.entries()).sort(function (a, b) { return b[1] - a[1]; })[0];
      if (!best || best[0] === t.l) return t.w;
      return t.cs === 'T' ? cap(best[0]) : t.cs === 'U' ? best[0].toLocaleUpperCase('en') : best[0];
    });
    return { words: words, seps: oc.toks.map(function (t) { return t.joint === 'hy' ? '-' : ' '; }) };
  }

  var KIND_LABEL = { number: 'tunggal/jamak', compound: 'spasi/tanda hubung/digabung', case: 'huruf kapital' };

  function termFinding(r, segs, body, lang) {
    var f = r.f;
    // bentuk tampilan per tanda tangan (+ kelas huruf untuk isu kapital)
    // Kemunculan "netral" (heading, awal kalimat, judul, tabel, definisi) tidak dinilai huruf
    // kapitalnya; dimasukkan ke baris bentuk yang sama dengan kelas huruf dominannya.
    var forms = new Map(), sigCase = new Map();
    f.occ.forEach(function (oc) {
      if (!r.caseIssue || !caseCounted(oc) || !oc.cc) return;
      var m = sigCase.get(oc.sig) || {}; m[oc.cc] = (m[oc.cc] || 0) + 1; sigCase.set(oc.sig, m);
    });
    function sigDefaultCase(sig) {
      if (!r.caseIssue) return null;
      if (sig === r.domSig) return r.domCase;
      var m = sigCase.get(sig); if (!m) return r.domCase;
      return Object.keys(m).sort(function (a, b) { return m[b] - m[a]; })[0];
    }
    f.occ.forEach(function (oc) {
      var counted = r.caseIssue && caseCounted(oc) && oc.cc;
      var cc = counted ? oc.cc : sigDefaultCase(oc.sig);
      var id = oc.sig + (cc ? '|' + cc : '');
      var fm = forms.get(id);
      if (!fm) { fm = { id: id, sig: oc.sig, cc: cc, occ: [], texts: new Map(), neutral: 0 }; forms.set(id, fm); }
      fm.occ.push(oc);
      if (!counted && r.caseIssue) fm.neutral++;
      var t = occText(segs, oc);
      if (counted || !r.caseIssue) fm.texts.set(t, (fm.texts.get(t) || 0) + 1);
      else if (!fm.texts.size) fm.fallback = t;
    });
    // representasi bentuk baku: kemunculan dengan tanda tangan dominan + kelas huruf dominan di tengah kalimat
    var pool = f.occ.filter(function (oc) { return oc.sig === r.domSig; });
    var rep = pool.filter(function (oc) { return caseCounted(oc) && (!r.domCase || oc.cc === r.domCase); })[0] ||
      pool.filter(function (oc) { return oc.mid; })[0] || pool[0];
    var templ = templateFrom(rep, pool);
    var domText = rebuild(templ, Object.assign({}, rep, { mid: true, def: false, zone: 'body' }), { lang: lang, domCase: r.domCase || caseClass(rep) });
    var formList = Array.from(forms.values()).map(function (fm) {
      var tx = Array.from(fm.texts.entries()).sort(function (a, b) { return b[1] - a[1]; });
      var label = tx.length ? tx[0][0] : fm.fallback;
      var isDom = fm.sig === r.domSig && (!r.caseIssue || !fm.cc || fm.cc === r.domCase);
      var note = fm.cc && r.caseIssue && fm.occ.length > fm.neutral ? (fm.cc === 'title' ? 'kapital di tengah kalimat' : fm.cc === 'mixed' ? 'kapital tidak beraturan' : 'huruf kecil') : '';
      if (fm.neutral) note += (note ? '; ' : '') + 'termasuk ' + fm.neutral + '× di heading/awal kalimat/tabel';
      return { label: label, count: fm.occ.length, dominant: isDom, occ: fm.occ.map(function (oc) { return loc(segs, body, oc); }), note: note };
    }).sort(function (a, b) { return (b.dominant - a.dominant) || (b.count - a.count); });
    // Bila hanya beda kapital, gabungkan bentuk "heading/awal kalimat" ke bentuk dominan agar daftar ringkas.
    var items = [];
    r.changes.forEach(function (oc) {
      var next = rebuild(templ, oc, { lang: lang, domCase: r.domCase });
      var cur = occText(segs, oc);
      if (next !== cur) items.push(Object.assign(loc(segs, body, oc), { next: next }));
    });
    var kinds = Array.from(r.kinds).map(function (k) { return KIND_LABEL[k]; });
    var mixedCase = r.caseIssue && r.ccount.mixed > 0;
    var status = r.kinds.has('number') || r.kinds.has('compound') ? (r.anchor ? 'bad' : 'warn') : mixedCase ? 'bad' : 'warn';
    if (!r.anchor && f.occ.length < 3 && !mixedCase && !r.kinds.has('compound')) status = 'info';
    var onlyCompound = r.kinds.size === 1 && r.kinds.has('compound');
    var group = onlyCompound && !r.anchor ? 'compound' : 'term';
    var total = f.occ.length;
    var detail = (r.anchor ? 'Istilah kunci (ada di judul, kata kunci, atau definisi singkatan). ' : '') + 'Beda ' + kinds.join(', ') + '. ' +
      'Seragamkan menjadi "' + domText + '"' + (r.caseIssue ? ' — ' + (r.domCase === 'title' ? 'kapital tiap kata' : 'huruf kecil') + ' di tengah kalimat' : '') + '.' +
      (r.kinds.has('number') ? ' Jamak/tunggal kata inti tetap mengikuti kalimat.' : '') +
      (r.caseIssue ? ' Awal kalimat, heading, judul, dan tabel tidak dihitung.' : '');
    return {
      id: 'term:' + f.key, group: group, status: status, kinds: Array.from(r.kinds), anchor: r.anchor,
      title: cap(domText.toLocaleLowerCase('en') === domText ? domText : domText) + ' — ' + formList.length + ' bentuk penulisan',
      detail: detail, recommended: domText, forms: formList, total: total,
      fix: items.length ? { label: 'Seragamkan ke "' + domText + '"', items: items } : null,
      location: items[0] || formList[0].occ[0]
    };
  }

  // ---------- EJAAN UK/US ----------
  function spellingFindings(segs, body) {
    var by = new Map();
    segs.forEach(function (sg, si) {
      sg.toks.forEach(function (t) {
        if (!t.alpha || t.acr || t.l.length < 4) return;
        var core = t.l.replace(/['’]s$/, '');
        var gk = singular(t.us);
        var g = by.get(gk);
        if (!g) by.set(gk, g = { key: gk, forms: new Map() });
        var fm = g.forms.get(core);
        if (!fm) g.forms.set(core, fm = { form: core, uk: core !== t.us, occ: [] });
        fm.occ.push({ seg: si, t: t });
      });
    });
    var out = [], ukTotal = 0, usTotal = 0;
    var groups = [];
    by.forEach(function (g) {
      if (g.forms.size < 2) return;
      var fs = Array.from(g.forms.values());
      // harus ada pasangan UK–US untuk bentuk (tunggal/jamak) yang sama
      var pairOk = fs.some(function (x) { return x.uk && fs.some(function (y) { return !y.uk && usKey(x.form) === y.form; }); });
      if (!pairOk) return;
      fs.forEach(function (x) { if (x.uk) ukTotal += x.occ.length; else usTotal += x.occ.length; });
      groups.push({ g: g, fs: fs });
    });
    groups.forEach(function (x) {
      var fs = x.fs.sort(function (a, b) { return b.occ.length - a.occ.length; });
      var nUk = 0, nUs = 0;
      fs.forEach(function (f) { if (f.uk) nUk += f.occ.length; else nUs += f.occ.length; });
      var useUk = nUk !== nUs ? nUk > nUs : ukTotal >= usTotal;
      var dom = fs.filter(function (f) { return f.uk === useUk; })[0];
      var ukOf = function (us) { var m = fs.filter(function (f) { return f.uk && usKey(f.form) === us; })[0]; return m ? m.form : null; };
      var items = [];
      fs.forEach(function (f) {
        if (f.uk === useUk) return;
        var target = useUk ? ukOf(f.form) : usKey(f.form);
        if (!target) return;
        f.occ.forEach(function (o) {
          var sg = segs[o.seg], cur = sg.text.slice(o.t.s - sg.base, o.t.e - sg.base);
          var next = target + (/['’]s$/.test(cur) ? cur.slice(-2) : '');
          if (o.t.cs === 'U') next = next.toLocaleUpperCase('en'); else if (o.t.cs === 'T') next = cap(next);
          if (next !== cur) items.push({ index: sg.index, offset: o.t.s, length: o.t.e - o.t.s, text: cur, zone: sg.zone, next: next });
        });
      });
      var total = fs.reduce(function (n, f) { return n + f.occ.length; }, 0);
      out.push({
        id: 'spell:' + x.g.key, group: 'spelling', status: 'warn', total: total,
        title: 'Ejaan UK/US bercampur: ' + fs.filter(function (f) { return f.uk === useUk; }).map(function (f) { return f.form; }).join(', ') + ' / ' + fs.filter(function (f) { return f.uk !== useUk; }).map(function (f) { return f.form; }).join(', '),
        detail: 'Kata yang sama ditulis dengan ejaan Inggris (' + fs.filter(function (f) { return f.uk; }).map(function (f) { return f.form; }).join(', ') + ') dan Amerika (' +
          fs.filter(function (f) { return !f.uk; }).map(function (f) { return f.form; }).join(', ') + '). Pilih satu ragam ejaan untuk seluruh naskah' +
          (ukTotal !== usTotal ? ' — naskah ini lebih banyak memakai ejaan ' + (ukTotal > usTotal ? 'Inggris (UK)' : 'Amerika (US)') : '') + '.',
        recommended: dom.form,
        forms: fs.map(function (f) {
          return { label: f.form, count: f.occ.length, dominant: f.uk === useUk, note: f.uk ? 'UK' : 'US', occ: f.occ.map(function (o) { var sg = segs[o.seg]; return { index: sg.index, offset: o.t.s, length: o.t.e - o.t.s, text: sg.text.slice(o.t.s - sg.base, o.t.e - sg.base), zone: sg.zone }; }) };
        }),
        fix: items.length ? { label: 'Seragamkan ke ejaan ' + (useUk ? 'UK' : 'US'), items: items } : null,
        location: items[0] || null
      });
    });
    return out;
  }

  // ---------- SINGKATAN ----------
  var ACR_SKIP = new Set('of and the for in on to with a an dan di untuk yang dalam terhadap pada atau &'.split(' '));
  function findAcronyms(segs, body) {
    var defs = [], mentions = new Map();
    segs.forEach(function (sg, si) {
      var text = sg.text, T = sg.toks;
      // (1) "Full Term (FT)"  (2) "FT (Full Term)"  (3) "Full Term (hereafter FT)"
      var re = /\(\s*(?:(?:hereafter|hereinafter|henceforth|selanjutnya|disingkat|abbreviated\s+as|referred\s+to\s+as|selanjutnya\s+disebut)\s*(?:referred\s+to\s+as\s+|as\s+|disebut\s+)?[,:]?\s*)?["“]?([\p{Lu}][\p{L}\p{N}&-]{1,11})["”]?\s*(?:[;,][^)]{0,60})?\)/gu, m;
      while ((m = re.exec(text))) {
        var abbr = m[1].replace(/[-]+$/, '');
        if (!isAcr(abbr) && !/^\p{Lu}[\p{Ll}]*\p{Lu}/u.test(abbr)) continue;
        if (/^\p{Lu}[\p{Ll}]+$/u.test(abbr)) continue;
        var letters = abbr.replace(/s$/, '').replace(/[^\p{Lu}\p{N}]/gu, '').replace(/\p{N}+/gu, '');
        if (letters.length < 2) continue;
        var parenStart = sg.base + m.index;
        var before = T.filter(function (t) { return t.e <= parenStart; });
        var match = matchInitials(before, letters);
        if (match) {
          var full = text.slice(match.s - sg.base, match.e - sg.base);
          defs.push({ abbr: abbr.replace(/s$/, '') === abbr ? abbr : abbr.replace(/s$/, ''), full: full, fullKey: match.toks.map(function (t) { return t.k; }).join(''),
            fullToks: match.toks, seg: si, index: sg.index, zone: sg.zone, fullStart: match.s, fullEnd: match.e, at: parenStart, len: m[0].length, sent: si + ':' + match.toks[0].sent });
        }
      }
      // (2) "FT (Full Term)"
      var re2 = /\b([\p{Lu}][\p{Lu}\p{N}&-]{1,9})\s*\(\s*([^()]{3,90}?)\s*\)/gu;
      while ((m = re2.exec(text))) {
        var ab = m[1];
        if (!isAcr(ab)) continue;
        var inner = m[2];
        var innerStart = sg.base + m.index + m[0].indexOf(inner, m[1].length);
        var toks = T.filter(function (t) { return t.s >= innerStart && t.e <= innerStart + inner.length; });
        var lt = ab.replace(/s$/, '').replace(/[^\p{Lu}]/gu, '');
        var mm = matchInitials(toks, lt, true);
        if (mm && mm.toks[0] === toks[0] && mm.toks[mm.toks.length - 1] === toks[toks.length - 1]) {
          defs.push({ abbr: ab, full: text.slice(mm.s - sg.base, mm.e - sg.base), fullKey: mm.toks.map(function (t) { return t.k; }).join(''), fullToks: mm.toks,
            seg: si, index: sg.index, zone: sg.zone, fullStart: mm.s, fullEnd: mm.e, at: sg.base + m.index, len: m[0].length, sent: si + ':' + mm.toks[0].sent, reverse: true });
        }
      }
      T.forEach(function (t, ti) {
        if (!t.acr) return;
        var base = t.w.replace(/s$/, '');
        var l = mentions.get(base); if (!l) mentions.set(base, l = []);
        l.push({ seg: si, index: sg.index, zone: sg.zone, s: t.s, e: t.e, text: t.w, sent: si + ':' + t.sent });
      });
    });
    return { defs: defs, mentions: mentions };
  }
  // Cocokkan huruf singkatan dengan huruf awal kata-kata sebelum tanda kurung (dari belakang).
  function matchInitials(toks, letters, forward) {
    var L = letters.toLocaleLowerCase('en').split('');
    if (!toks.length) return null;
    var end = toks.length - 1;
    var i = end, k = L.length - 1, used = [];
    var skipped = 0;
    while (k >= 0 && i >= 0) {
      var t = toks[i];
      var w = t.l;
      if (!t.alpha) return null;
      if (w.charAt(0) === L[k]) {
        // satu kata boleh menyumbang beberapa huruf ("GenAI" → Generative AI; "Self-Efficacy" → SE)
        used.unshift(t); k--;
        i--; skipped = 0;
        continue;
      }
      if (ACR_SKIP.has(w) && used.length && skipped < 2) { i--; skipped++; continue; }
      // kata dengan beberapa huruf kapital di dalamnya (mis. "eLearning", "PowerPoint")
      var caps = t.w.replace(/[^\p{Lu}]/gu, '').toLocaleLowerCase('en');
      if (caps.length > 1 && caps.slice(-1) === L[k]) {
        var cs = caps.split(''), kk = k, ok = true;
        for (var c = cs.length - 1; c >= 0 && kk >= 0; c--) { if (cs[c] !== L[kk]) { ok = false; break; } kk--; }
        if (ok) { used.unshift(t); k = kk; i--; continue; }
      }
      return null;
    }
    if (k >= 0 || !used.length) return null;
    // jangan mulai dari kata sambung
    while (used.length && ACR_SKIP.has(used[0].l)) used.shift();
    if (used.length < 2 && letters.length > 2) return null;
    // kata-kata harus bersambung (spasi/tanda hubung) — tidak melintasi koma atau titik
    var first = toks.indexOf(used[0]);
    for (var q = first + 1; q <= end; q++) if (toks[q].joint === 'brk') return null;
    var all = toks.slice(first, end + 1);
    return { toks: all, s: all[0].s, e: all[all.length - 1].e };
  }

  function normFull(d) { return d.fullToks.map(function (t) { return t.k; }).join(' '); }

  function acronymFindings(acr, segs, body, fam, lang) {
    var out = [];
    var byAbbr = new Map();
    acr.defs.forEach(function (d) { var l = byAbbr.get(d.abbr); if (!l) byAbbr.set(d.abbr, l = []); l.push(d); });
    function dloc(d) { return { index: d.index, offset: d.fullStart, length: d.at + d.len - d.fullStart, text: segs[d.seg].text.slice(d.fullStart - segs[d.seg].base, d.at + d.len - segs[d.seg].base), zone: d.zone }; }
    function mloc(m) { return { index: m.index, offset: m.s, length: m.e - m.s, text: m.text, zone: m.zone }; }
    var ORDER = function (a, b) { return a.seg - b.seg || a.at - b.at; };

    // (a) satu singkatan, beberapa kepanjangan
    byAbbr.forEach(function (list, abbr) {
      var fulls = new Map();
      list.forEach(function (d) { var k = normFull(d); var l = fulls.get(k); if (!l) fulls.set(k, l = []); l.push(d); });
      if (fulls.size < 2) return;
      var forms = Array.from(fulls.values()).map(function (l) { return { label: l[0].full, count: l.length, occ: l.map(dloc) }; }).sort(function (a, b) { return b.count - a.count; });
      forms[0].dominant = true;
      out.push({ id: 'acr:multi:' + abbr, group: 'acronym', status: 'bad', total: list.length,
        title: abbr + ' punya ' + fulls.size + ' kepanjangan berbeda',
        detail: 'Singkatan ' + abbr + ' didefinisikan sebagai ' + forms.map(function (f) { return '"' + f.label + '"'; }).join(' dan ') + '. Satu singkatan harus selalu merujuk ke satu nama yang sama.',
        forms: forms, location: forms[1].occ[0] });
    });
    // (b) satu kepanjangan, beberapa singkatan
    var byFull = new Map();
    acr.defs.forEach(function (d) { var k = d.fullKey; var l = byFull.get(k); if (!l) byFull.set(k, l = []); l.push(d); });
    byFull.forEach(function (list) {
      var abbrs = Array.from(new Set(list.map(function (d) { return d.abbr; })));
      if (abbrs.length < 2) return;
      var forms = abbrs.map(function (a) { var l = list.filter(function (d) { return d.abbr === a; }); return { label: l[0].full + ' (' + a + ')', count: (acr.mentions.get(a) || []).length, occ: l.map(dloc) }; }).sort(function (a, b) { return b.count - a.count; });
      forms[0].dominant = true;
      out.push({ id: 'acr:samefull:' + list[0].fullKey, group: 'acronym', status: 'bad', total: list.length,
        title: '"' + list[0].full + '" disingkat ' + abbrs.join(' dan '),
        detail: 'Nama yang sama memakai ' + abbrs.length + ' singkatan berbeda (' + abbrs.join(', ') + '). Pilih satu singkatan untuk seluruh naskah.',
        forms: forms, location: forms[1].occ[0] });
    });
    // (c) dua singkatan berbeda untuk nama yang hampir sama (mis. instrumen yang sama diberi dua nama)
    var uniq = [];
    byAbbr.forEach(function (list, abbr) { uniq.push({ abbr: abbr, d: list[0], toks: list[0].fullToks.map(function (t) { return t.k; }) }); });
    for (var i = 0; i < uniq.length; i++) for (var j = i + 1; j < uniq.length; j++) {
      var A = uniq[i], B = uniq[j];
      if (A.d.fullKey === B.d.fullKey) continue;
      var diff = nearSame(A.toks, B.toks);
      if (!diff) continue;
      var nA = (acr.mentions.get(A.abbr) || []).length, nB = (acr.mentions.get(B.abbr) || []).length;
      // Dipakai bersama dalam satu kalimat → memang dua hal berbeda.
      var sentA = new Set((acr.mentions.get(A.abbr) || []).map(function (m) { return m.sent; }));
      var together = (acr.mentions.get(B.abbr) || []).some(function (m) { return sentA.has(m.sent); });
      if (together) continue;
      var dom = nA >= nB ? A : B, other = dom === A ? B : A;
      var occOf = function (x) { return (acr.mentions.get(x.abbr) || []).map(mloc); };
      out.push({ id: 'acr:twonames:' + A.abbr + '|' + B.abbr, group: 'acronym', status: 'warn', total: nA + nB,
        title: 'Dua nama untuk hal yang sama? ' + dom.abbr + ' vs ' + other.abbr,
        detail: '"' + dom.d.full + '" (' + dom.abbr + ') dan "' + other.d.full + '" (' + other.abbr + ') hanya berbeda pada kata "' + diff.a + '"/"' + diff.b + '" dan tidak pernah disebut bersama. ' +
          'Bila keduanya instrumen/variabel yang sama, pakai satu nama & satu singkatan saja.',
        forms: [{ label: dom.d.full + ' (' + dom.abbr + ')', count: Math.max(1, (dom === A ? nA : nB)), dominant: true, occ: [dloc(dom.d)].concat(occOf(dom)) },
          { label: other.d.full + ' (' + other.abbr + ')', count: Math.max(1, (dom === A ? nB : nA)), occ: [dloc(other.d)].concat(occOf(other)) }],
        location: dloc(other.d) });
    }
    // (d) dipakai sebelum didefinisikan, (e) tidak pernah dipakai, (f) bentuk panjang masih dipakai
    byAbbr.forEach(function (list, abbr) {
      var defs = list.slice().sort(ORDER);
      var ms = (acr.mentions.get(abbr) || []).filter(function (m) { return !defs.some(function (d) { return d.seg === m.seg && m.s >= d.fullStart && m.e <= d.at + d.len; }); });
      var useful = ms.filter(function (m) { return m.zone === 'body' || m.zone === 'abstract'; });
      // urutan dokumen: indeks segmen
      var bodyDefs = defs.filter(function (d) { return d.zone === 'body'; });
      var absDefs = defs.filter(function (d) { return d.zone === 'abstract'; });
      [['abstract', absDefs], ['body', bodyDefs]].forEach(function (z) {
        var zoneMs = useful.filter(function (m) { return m.zone === z[0]; });
        if (!zoneMs.length) return;
        var firstDef = z[1][0];
        var firstUse = zoneMs[0];
        if (z[0] === 'body' && !firstDef && absDefs.length) return; // didefinisikan di abstrak saja — dibiarkan
        if (firstDef && (firstUse.seg < firstDef.seg || firstUse.seg === firstDef.seg && firstUse.s < firstDef.fullStart)) {
          out.push({ id: 'acr:early:' + z[0] + ':' + abbr, group: 'acronym', status: 'warn', total: zoneMs.length,
            title: abbr + ' dipakai sebelum didefinisikan' + (z[0] === 'abstract' ? ' (abstrak)' : ''),
            detail: abbr + ' sudah dipakai sebelum definisi "' + firstDef.full + ' (' + abbr + ')". Pindahkan definisinya ke pemakaian pertama.',
            forms: [{ label: 'pemakaian pertama', count: 1, occ: [mloc(firstUse)] }, { label: 'definisi', count: z[1].length, dominant: true, occ: z[1].map(dloc) }],
            location: mloc(firstUse) });
        }
      });
      if (bodyDefs.length > 1) {
        out.push({ id: 'acr:redef:' + abbr, group: 'acronym', status: 'info', total: bodyDefs.length,
          title: abbr + ' didefinisikan ' + bodyDefs.length + ' kali di isi naskah',
          detail: 'Cukup definisikan sekali pada pemakaian pertama di isi naskah, selanjutnya pakai singkatannya.',
          forms: [{ label: 'definisi', count: bodyDefs.length, occ: bodyDefs.map(dloc) }], location: dloc(bodyDefs[1]) });
      }
      if (!ms.length) {
        out.push({ id: 'acr:unused:' + abbr, group: 'acronym', status: 'info', total: defs.length,
          title: abbr + ' didefinisikan tetapi tidak dipakai',
          detail: 'Singkatan ' + abbr + ' ("' + defs[0].full + '") tidak pernah dipakai lagi setelah didefinisikan. Hapus singkatannya atau pakai secara konsisten.',
          forms: [{ label: 'definisi', count: defs.length, occ: defs.map(dloc) }], location: dloc(defs[0]) });
      }
      // bentuk panjang masih ditulis utuh setelah singkatan didefinisikan (di isi naskah)
      var f = fam.get(defs[0].fullKey);
      var anchorDef = bodyDefs[0];
      if (f && anchorDef && ms.length) {
        var after = f.occ.filter(function (oc) { return oc.zone === 'body' && !oc.def && (oc.seg > anchorDef.seg || oc.seg === anchorDef.seg && oc.s > anchorDef.at); });
        if (after.length >= 2) {
          out.push({ id: 'acr:full:' + abbr, group: 'acronym', status: 'info', total: after.length,
            title: '"' + defs[0].full + '" masih ditulis lengkap ' + after.length + '× setelah disingkat ' + abbr,
            detail: 'Setelah singkatan didefinisikan, sebaiknya pakai ' + abbr + ' secara konsisten (kecuali di awal kalimat atau heading bila gaya jurnal meminta).',
            forms: [{ label: abbr, count: ms.length, dominant: true, occ: ms.map(mloc) }, { label: defs[0].full, count: after.length, occ: after.map(function (oc) { return loc(segs, body, oc); }) }],
            location: loc(segs, body, after[0]) });
        }
      }
    });
    // (g) singkatan dipakai berulang tanpa definisi
    acr.mentions.forEach(function (ms, abbr) {
      if (byAbbr.has(abbr) || KNOWN_ABBR.has(abbr) || abbr.length < 2) return;
      var useful = ms.filter(function (m) { return m.zone === 'body' || m.zone === 'abstract'; });
      if (useful.length < 3) return;
      if (/^[IVXLCDM]+$/.test(abbr)) return;
      // didefinisikan dalam pola lain (mis. "ABBR stands for", "ABBR, i.e.")? lewati bila ada "=" di dekatnya
      out.push({ id: 'acr:undef:' + abbr, group: 'acronym', status: 'info', total: useful.length,
        title: abbr + ' dipakai ' + useful.length + '× tanpa definisi',
        detail: 'Tidak ditemukan kepanjangan "Nama Lengkap (' + abbr + ')". Bila bukan singkatan umum, definisikan pada pemakaian pertama.',
        forms: [{ label: abbr, count: useful.length, occ: useful.map(mloc) }], location: mloc(useful[0]) });
    });
    return out;
  }
  // Dua nama sepanjang sama yang hanya berbeda satu kata (bukan pasangan lawan kata) dan kata terakhirnya sama.
  function nearSame(a, b) {
    if (a.length !== b.length || a.length < 2) return null;
    var d = [];
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) d.push(i);
    if (d.length !== 1) return null;
    var i0 = d[0];
    if (i0 === a.length - 1 && !(SYN.get(a[i0]) && SYN.get(a[i0]).has(b[i0]))) return null;
    var x = a[i0], y = b[i0];
    if (CONTRAST.some(function (p) { return p[0] === x && p[1] === y || p[1] === x && p[0] === y; })) return null;
    if (/^\d/.test(x) || /^\d/.test(y)) return null;
    return { a: x, b: y };
  }

  // ---------- SINONIM ----------
  function synonymFindings(fam, anchorKeys, segs, body, lang) {
    var out = [];
    var byToks = new Map(); // key-dengan-spasi → famili (hanya frasa banyak kata)
    fam.forEach(function (f) {
      var oc = f.occ[0];
      if (oc.tn < 2) return;
      byToks.set(oc.toks.map(function (t) { return t.k; }).join(' '), f);
    });
    var pairs = [];
    var done = new Set();
    byToks.forEach(function (f, k) {
      if (f.nested && !anchorKeys.has(f.key)) return;
      if (f.occ.length < 2 && !anchorKeys.has(f.key)) return;
      var toks = k.split(' ');
      toks.forEach(function (w, i) {
        var syns = SYN.get(w); if (!syns) return;
        syns.forEach(function (s) {
          var k2 = toks.slice(0, i).concat([s], toks.slice(i + 1)).join(' ');
          var g = byToks.get(k2);
          if (!g || g === f || g.nested && !anchorKeys.has(g.key)) return;
          var id = [f.key, g.key].sort().join('|');
          if (done.has(id)) return;
          done.add(id);
          pairs.push({ a: f, b: g, w1: w, w2: s });
        });
      });
    });
    var ADJ_END = /(ic|ical|ive|ous|ful|less|able|ible|ary|al)$/;
    pairs.forEach(function (p) {
      var la = p.a.occ[0].toks[p.a.occ[0].toks.length - 1], lb = p.b.occ[0].toks[p.b.occ[0].toks.length - 1];
      if (ADJ_END.test(la.k) && !SYN.has(la.k) || ADJ_END.test(lb.k) && !SYN.has(lb.k)) return;
      if (['academic', 'educational', 'scholastic', 'online', 'digital', 'virtual', 'organizational', 'corporate', 'occupational', 'akademik', 'akademis'].indexOf(la.k) >= 0) return;
      var sa = new Set(p.a.occ.map(function (o) { return o.sent; }));
      if (p.b.occ.some(function (o) { return sa.has(o.sent); })) return; // muncul bersama → konsep berbeda
      // Kedua frasa harus "istilah": tidak terserap ke frasa lain yang berbeda maknanya.
      var A = p.a, B = p.b;
      var nA = A.occ.length, nB = B.occ.length;
      if (nA + nB < 3) return;
      var anchor = anchorKeys.has(A.key) || anchorKeys.has(B.key) || A.inTitle || B.inTitle;
      var dom = nA > nB || nA === nB && (anchorKeys.has(A.key) || A.inTitle) ? A : B, other = dom === A ? B : A;
      var rep = dom.occ.filter(function (o) { return o.mid && (o.zone === 'body' || o.zone === 'abstract'); })[0] || dom.occ[0];
      var templ = templateFrom(rep);
      var domText = rebuild(templ, Object.assign({}, rep, { mid: true, def: false, zone: 'body' }), { lang: lang, domCase: caseClass(rep) === 'title' ? 'title' : 'lower' });
      var items = other.occ.map(function (oc) {
        var cur = occText(segs, oc), next = rebuild(templ, oc, { lang: lang, domCase: caseClass(oc) || 'lower' });
        return next !== cur ? Object.assign(loc(segs, body, oc), { next: next }) : null;
      }).filter(Boolean);
      var label = function (f) {
        var m = new Map(); f.occ.forEach(function (o) { var t = occText(segs, o); m.set(t, (m.get(t) || 0) + 1); });
        return Array.from(m.entries()).sort(function (a, b) { return b[1] - a[1]; })[0][0];
      };
      out.push({
        id: 'syn:' + [A.key, B.key].sort().join('|'), group: 'synonym', status: 'warn', anchor: anchor, total: nA + nB,
        title: 'Sinonim untuk variabel yang sama? "' + label(dom) + '" / "' + label(other) + '"',
        detail: 'Dua istilah ini hanya berbeda pada "' + (dom === A ? p.w1 : p.w2) + '"/"' + (dom === A ? p.w2 : p.w1) + '" dan tidak pernah dipakai bersama dalam satu kalimat — kemungkinan nama berbeda untuk konsep/variabel yang sama. ' +
          'Bila memang sama, pakai satu istilah (' + label(dom) + ', ' + dom.occ.length + '×).',
        recommended: domText,
        forms: [{ label: label(dom), count: dom.occ.length, dominant: true, occ: dom.occ.map(function (o) { return loc(segs, body, o); }) },
          { label: label(other), count: other.occ.length, occ: other.occ.map(function (o) { return loc(segs, body, o); }) }],
        fix: items.length ? { label: 'Ganti ke "' + domText + '"', items: items } : null,
        location: items[0] || loc(segs, body, other.occ[0])
      });
    });
    return out;
  }

  // ---------- SALAH KETIK DALAM ISTILAH ----------
  // "behavioural intention" (12×) vs "behavioural intension" (1×): satu kata berbeda 1–2 huruf,
  // dan kata yang berbeda itu tidak pernah dipakai di tempat lain dalam naskah.
  function typoFindings(fam, anchorKeys, segs, body, lang) {
    var out = [], uni = new Map();
    fam.forEach(function (f) { if (f.occ[0].tn === 1 && f.occ.every(function (o) { return o.tn === 1; })) uni.set(f.key, f.occ.length); });
    var multi = [];
    fam.forEach(function (f) { if (f.occ[0].tn >= 2 && !f.nested && f.occ.every(function (o) { return o.tn === f.occ[0].tn; })) multi.push(f); });
    var byLen = new Map();
    multi.forEach(function (f) { var n = f.occ[0].tn; var l = byLen.get(n); if (!l) byLen.set(n, l = []); l.push(f); });
    var done = new Set();
    multi.forEach(function (dom) {
      if (dom.occ.length < 3) return;
      var dt = dom.occ[0].toks.map(function (t) { return t.k; });
      (byLen.get(dt.length) || []).forEach(function (g) {
        if (g === dom || g.occ.length * 3 > dom.occ.length || done.has(g.key)) return;
        var gt = g.occ[0].toks.map(function (t) { return t.k; }), diff = -1;
        for (var i = 0; i < dt.length; i++) if (dt[i] !== gt[i]) { if (diff >= 0) return; diff = i; }
        if (diff < 0) return;
        var a = dt[diff], b = gt[diff];
        if (a.length < 5 || b.length < 5 || STOP.has(b)) return;
        var d = lev(a, b);
        if (d < 1 || d > Math.max(1, Math.floor(a.length / 5))) return;
        if ((uni.get(b) || 0) > g.occ.length) return; // kata itu dipakai juga di tempat lain → bukan salah ketik
        if (SYN.has(a) && SYN.get(a).has(b)) return;
        done.add(g.key);
        var rep = dom.occ.filter(function (o) { return o.mid && (o.zone === 'body' || o.zone === 'abstract'); })[0] || dom.occ[0];
        var word = rep.toks[diff].l;
        var items = g.occ.map(function (oc) {
          var t = oc.toks[diff], cur = occText(segs, oc);
          var w = t.cs === 'U' ? word.toLocaleUpperCase('en') : t.cs === 'T' ? cap(word) : word;
          if (lang !== 'id' && plural(t) && !plural(rep.toks[diff])) w = pluralize(w);
          var sg = segs[oc.seg];
          var next = cur.slice(0, t.s - oc.s) + w + cur.slice(t.e - oc.s);
          return next !== cur ? Object.assign(loc(segs, body, oc), { next: next }) : null;
        }).filter(Boolean);
        var lab = function (f) { var m = new Map(); f.occ.forEach(function (o) { var t = occText(segs, o).toLocaleLowerCase('en'); m.set(t, (m.get(t) || 0) + 1); }); return Array.from(m.entries()).sort(function (x, y) { return y[1] - x[1]; })[0][0]; };
        out.push({ id: 'typo:' + g.key, group: 'typo', status: 'bad', total: dom.occ.length + g.occ.length, anchor: anchorKeys.has(dom.key),
          title: 'Kemungkinan salah ketik: "' + lab(g) + '"',
          detail: 'Di tempat lain istilah ini ditulis "' + lab(dom) + '" (' + dom.occ.length + '×); "' + g.occ[0].toks[diff].l + '" tidak dipakai di bagian lain naskah.',
          recommended: lab(dom),
          forms: [{ label: lab(dom), count: dom.occ.length, dominant: true, occ: dom.occ.map(function (o) { return loc(segs, body, o); }) },
            { label: lab(g), count: g.occ.length, occ: g.occ.map(function (o) { return loc(segs, body, o); }) }],
          fix: items.length ? { label: 'Perbaiki ke "' + lab(dom) + '"', items: items } : null,
          location: items[0] || loc(segs, body, g.occ[0]) });
      });
    });
    return out;
  }

  // ---------- KATA KUNCI ----------
  function keywordFindings(segs, fam, lang) {
    var out = [];
    segs.forEach(function (sg) {
      if (sg.zone !== 'keywords' || !sg.toks.length) return;
      var key = sg.toks.filter(function (t) { return content(t) || t.alpha && t.l === 'of'; }).map(function (t) { return t.k; }).join('');
      if (!key) return;
      var f = fam.get(key);
      var inText = f && f.occ.some(function (o) { return o.zone === 'body' || o.zone === 'abstract'; });
      if (inText) return;
      // acronym keywords: "NEP 2020", "AI" → cek kemunculan token
      if (sg.toks.some(function (t) { return /\d/.test(t.w); })) {
        var re = new RegExp(sg.keyword.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'), 'i');
        if (segs.some(function (s) { return (s.zone === 'body' || s.zone === 'abstract') && re.test(s.text); })) return;
      }
      // cari kata yang paling mirip di naskah
      // Semua kata kuncinya ada dalam satu kalimat ("retention of girl students") → dianggap dipakai.
      var need = sg.toks.filter(content).map(function (t) { return t.k; });
      var bySent = new Map();
      segs.forEach(function (s2) {
        if (s2.zone !== 'body' && s2.zone !== 'abstract') return;
        s2.toks.forEach(function (t) { var id = s2.index + ':' + t.sent; var st = bySent.get(id); if (!st) bySent.set(id, st = new Set()); st.add(t.k); });
      });
      var loose = false;
      bySent.forEach(function (st) { if (!loose && need.length && need.every(function (k) { return st.has(k); })) loose = true; });
      if (loose) return;
      // cari istilah yang paling mirip di naskah (salah ketik per kata, jumlah kata sama)
      var best = null, kt = sg.toks.filter(function (t) { return t.alpha; }).map(function (t) { return t.k; });
      fam.forEach(function (g) {
        var g0 = g.occ[0];
        if (g0.tn !== kt.length || !g.occ.some(function (o) { return o.zone === 'body' || o.zone === 'abstract'; })) return;
        var gt = g0.toks.map(function (t) { return t.k; }), d = 0;
        for (var i = 0; i < kt.length; i++) {
          var di = lev(kt[i], gt[i]);
          if (di > Math.max(1, Math.floor(kt[i].length / 4))) return;
          d += di;
        }
        if (d > 0 && (!best || d < best.d || d === best.d && g.occ.length > best.f.occ.length)) best = { d: d, f: g };
      });
      var bestOcc = best ? best.f.occ.filter(function (o) { return o.mid && (o.zone === 'body' || o.zone === 'abstract'); })[0] || best.f.occ[0] : null;
      var sugg = best ? occText(segs, bestOcc) : null;
      if (sugg && bestOcc.mid && rawClass(bestOcc) !== 'title') sugg = sugg.toLocaleLowerCase('en');
      out.push({ id: 'kw:absent:' + key, group: 'keyword', status: best ? 'bad' : 'info', total: 1,
        title: 'Kata kunci "' + sg.keyword.trim() + '" tidak muncul di abstrak/isi',
        detail: best ? 'Kemungkinan salah ketik — di naskah tertulis "' + sugg + '" (' + best.f.occ.length + '×).' : 'Kata kunci sebaiknya memakai istilah yang sama persis dengan yang dipakai di abstrak/isi naskah.',
        recommended: sugg,
        forms: [{ label: sg.keyword.trim(), count: 1, occ: [{ index: sg.index, offset: sg.base, length: sg.text.length, text: sg.text, zone: 'keywords' }] }],
        fix: best ? { label: 'Ganti ke "' + sugg + '"', items: [{ index: sg.index, offset: sg.toks[0].s, length: sg.toks[sg.toks.length - 1].e - sg.toks[0].s,
          text: sg.text.slice(sg.toks[0].s - sg.base, sg.toks[sg.toks.length - 1].e - sg.base), zone: 'keywords',
          next: keywordFix(sg, best.f, segs) }] } : null,
        location: { index: sg.index, offset: sg.base, length: sg.text.length, text: sg.text, zone: 'keywords' } });
    });
    return out;
  }
  // Perbaiki hanya kata yang salah ketik; ejaan & kapital kata lain dipertahankan.
  function keywordFix(sg, f, segs) {
    var out = '', toks = sg.toks.filter(function (t) { return t.alpha; });
    toks.forEach(function (t, i) {
      var w = t.w;
      var tally = new Map();
      f.occ.forEach(function (o) { var x = o.toks[i]; if (x && x.k !== t.k) tally.set(x.l, (tally.get(x.l) || 0) + 1); });
      if (tally.size && lev(t.k, f.occ[0].toks[i].k) > 0) {
        var b = Array.from(tally.entries()).sort(function (p, q) { return q[1] - p[1]; })[0][0];
        w = t.cs === 'U' ? b.toLocaleUpperCase('en') : t.cs === 'T' ? cap(b) : b;
      }
      out += (i ? sg.text.slice(toks[i - 1].e - sg.base, t.s - sg.base) : '') + w;
    });
    return out;
  }
  function lev(a, b) {
    if (Math.abs(a.length - b.length) > 3) return 99;
    var dp = []; for (var i = 0; i <= a.length; i++) { dp[i] = [i]; }
    for (var j = 1; j <= b.length; j++) dp[0][j] = j;
    for (i = 1; i <= a.length; i++) for (j = 1; j <= b.length; j++) dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return dp[a.length][b.length];
  }

  // ---------- LAPORAN KOMENTAR ----------
  // Satu komentar per kemunculan yang perlu diubah (maks. o.perFinding per temuan).
  function reportComments(result, opts) {
    var o = opts || {}, per = o.perFinding || 12, out = [];
    result.findings.forEach(function (f) {
      if (o.skip && o.skip(f)) return;
      if (f.status === 'info' && !o.includeInfo) return;
      var targets = f.fix ? f.fix.items : f.location ? [f.location] : [];
      if (!targets.length) return;
      var head = commentText(f);
      targets.slice(0, per).forEach(function (t, i) {
        var txt = f.fix && t.next ? 'Istilah — "' + t.text + '" → "' + t.next + '". ' + (i === 0 ? head : 'Lihat komentar pertama untuk istilah ini.') : 'Istilah — ' + head;
        out.push({ id: f.id, status: f.status, loc: { index: t.index, offset: t.offset, length: t.length, find: t.text }, text: txt });
      });
      if (targets.length > per) out[out.length - 1].text += ' (dan ' + (targets.length - per) + ' kemunculan lain)';
    });
    return out;
  }
  function commentText(f) {
    var forms = (f.forms || []).map(function (x) { return '"' + x.label + '" ' + x.count + '×'; }).join(', ');
    return f.title + '. ' + f.detail + (forms ? ' Bentuk yang dipakai: ' + forms + '.' : '');
  }

  var api = { analyze: analyze, reportComments: reportComments, usKey: usKey, singular: singular, tokenize: tokenize, matchInitials: matchInitials, layout: layout, SYN_GROUPS: SYN_GROUPS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TermCore = api;
})(typeof window !== 'undefined' ? window : this);
