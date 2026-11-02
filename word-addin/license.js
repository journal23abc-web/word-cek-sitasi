/*!
 * license.js — gerbang lisensi bersama untuk add-in Word.
 * File yang SAMA dipakai di word-template dan word-cek-sitasi; yang beda
 * hanya atribut data-produk di tag <script>:
 *
 *   <script src="license.js" data-produk="template"></script>   (word-template)
 *   <script src="license.js" data-produk="sitasi"></script>     (word-cek-sitasi)
 *
 * Pengguna memasukkan kode SEKALI. Token disimpan di perangkat, jadi update
 * kode di GitHub tidak meminta kode lagi (alamat situs tetap sama).
 */
(function (global) {
  'use strict';

  // ====== GANTI dengan URL Web App dari Apps Script (akhiran /exec) ======
  var API = 'https://script.google.com/macros/s/AKfycbwqbjlmnK2Ln7YjYrspTb0fKSp0OuI0qvizy9z_HcDQ4tcXtrsxFYnFFUNEA-6hjkZX/exec';
  // =======================================================================

  var CEK_ULANG_MS = 3 * 864e5;      // cek ulang ke server tiap 3 hari (di latar belakang)
  var OFFLINE_MAKS_MS = 14 * 864e5;  // boleh dipakai offline paling lama 14 hari sejak cek terakhir
  var NAMA = { template: 'Template Paste', sitasi: 'Cek Sitasi' };

  var doc = global.document;
  var cryptoObj = global.crypto || global.msCrypto;
  var sedangMenunggu = {};

  /* ---------- penyimpanan ---------- */
  function simpan(k, v) {
    try { if (v === null) global.localStorage.removeItem(k); else global.localStorage.setItem(k, v); } catch (e) { /* abaikan */ }
  }
  function ambil(k) { try { return global.localStorage.getItem(k); } catch (e) { return null; } }

  function uuid() {
    if (cryptoObj && cryptoObj.randomUUID) return cryptoObj.randomUUID();
    var b = new Uint8Array(16);
    cryptoObj.getRandomValues(b);
    b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
    var h = '';
    for (var i = 0; i < 16; i++) h += (b[i] + 256).toString(16).slice(1);
    return h.slice(0, 8) + '-' + h.slice(8, 12) + '-' + h.slice(12, 16) + '-' + h.slice(16, 20) + '-' + h.slice(20);
  }

  function idPerangkat() {
    var d = ambil('lisensi.device');
    if (!d || !/^[0-9a-f-]{36}$/.test(d)) { d = uuid(); simpan('lisensi.device', d); }
    return d;
  }

  /* ---------- server ---------- */
  function panggil(data) {
    // text/plain = "simple request": tidak ada preflight CORS ke Apps Script
    return fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(data),
      redirect: 'follow'
    }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    });
  }

  /* ---------- tampilan kunci ---------- */
  function pasangGaya() {
    if (doc.getElementById('lisensi-gaya')) return;
    var s = doc.createElement('style');
    s.id = 'lisensi-gaya';
    s.textContent =
      'html.lisensi-kunci body>*:not(#lisensi-layar){visibility:hidden!important}' +
      '#lisensi-layar{position:fixed;top:0;left:0;right:0;bottom:0;z-index:2147483647;background:#fff;color:#1f2328;' +
      'display:flex;align-items:center;justify-content:center;padding:16px;font:14px/1.5 "Segoe UI",system-ui,sans-serif}' +
      '#lisensi-layar .kotak{width:100%;max-width:320px}' +
      '#lisensi-layar h1{font-size:17px;margin:0 0 4px}' +
      '#lisensi-layar p{margin:0 0 14px;color:#57606a}' +
      '#lisensi-layar input{width:100%;box-sizing:border-box;padding:9px 10px;font:inherit;font-family:Consolas,monospace;' +
      'letter-spacing:1px;text-transform:uppercase;border:1px solid #c9d1d9;border-radius:6px}' +
      '#lisensi-layar button{margin-top:10px;width:100%;padding:9px;font:inherit;font-weight:600;color:#fff;' +
      'background:#2b579a;border:0;border-radius:6px;cursor:pointer}' +
      '#lisensi-layar button[disabled]{opacity:.6;cursor:default}' +
      '#lisensi-layar .pesan{min-height:20px;margin-top:10px;color:#cf222e}';
    (doc.head || doc.documentElement).appendChild(s);
  }

  function kunciLayar() { pasangGaya(); doc.documentElement.className += ' lisensi-kunci'; }
  function bukaLayar() {
    doc.documentElement.className = doc.documentElement.className.replace(/\s*lisensi-kunci/g, '');
    var el = doc.getElementById('lisensi-layar');
    if (el) el.parentNode.removeChild(el);
  }

  function saatSiap(fn) {
    if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', fn); else fn();
  }

  function tampilkanForm(produk, pesanAwal) {
    return new Promise(function (resolve) {
      kunciLayar();
      saatSiap(function () {
        var lama = doc.getElementById('lisensi-layar');
        if (lama) lama.parentNode.removeChild(lama);

        var layar = doc.createElement('div');
        layar.id = 'lisensi-layar';
        layar.innerHTML =
          '<form class="kotak" autocomplete="off">' +
          '<h1></h1><p>Masukkan kode lisensi Anda. Kode hanya perlu dimasukkan sekali di perangkat ini.</p>' +
          '<input name="kode" placeholder="XXXX-XXXX-XXXX-XXXX" maxlength="40" spellcheck="false">' +
          '<button type="submit">Aktifkan</button><div class="pesan" role="alert"></div></form>';
        layar.querySelector('h1').textContent = 'Aktivasi ' + (NAMA[produk] || produk);
        var form = layar.querySelector('form'), input = form.kode,
            tombol = form.querySelector('button'), pesan = form.querySelector('.pesan');
        pesan.textContent = pesanAwal || '';
        doc.body.appendChild(layar);
        input.focus();

        form.addEventListener('submit', function (ev) {
          ev.preventDefault();
          var kode = input.value.trim();
          if (!kode) { pesan.textContent = 'Kode belum diisi.'; return; }
          tombol.disabled = true; tombol.textContent = 'Memeriksa…'; pesan.textContent = '';
          panggil({ aksi: 'aktivasi', kode: kode, produk: produk, device: idPerangkat() })
            .then(function (res) {
              if (res && res.ok && res.token) {
                simpan('lisensi.token.' + produk, res.token);
                simpan('lisensi.ok.' + produk, String(Date.now()));
                bukaLayar();
                resolve(true);
              } else {
                pesan.textContent = (res && res.pesan) || 'Aktivasi gagal.';
              }
            })
            .catch(function () { pesan.textContent = 'Tidak bisa terhubung ke server lisensi. Periksa internet.'; })
            .then(function () { tombol.disabled = false; tombol.textContent = 'Aktifkan'; });
        });
      });
    });
  }

  /* ---------- alur utama ---------- */
  function cekServer(produk, token) {
    return panggil({ aksi: 'cek', token: token, device: idPerangkat() }).then(function (res) {
      if (res && res.ok) { simpan('lisensi.ok.' + produk, String(Date.now())); return 'ok'; }
      if (res && res.cabut) { simpan('lisensi.token.' + produk, null); return res.pesan || 'Lisensi tidak berlaku.'; }
      throw new Error('jawaban tidak dikenal');
    });
  }

  /** Promise yang selesai setelah lisensi produk ini sah. */
  function pastikan(produk) {
    if (sedangMenunggu[produk]) return sedangMenunggu[produk];
    var token = ambil('lisensi.token.' + produk);
    var okTerakhir = Number(ambil('lisensi.ok.' + produk) || 0);
    var umur = Date.now() - okTerakhir;
    var janji;

    if (!token) {
      janji = tampilkanForm(produk);
    } else if (umur < CEK_ULANG_MS) {
      // jalur cepat: langsung buka, lalu cek diam-diam di latar belakang
      janji = Promise.resolve(true);
      cekServer(produk, token).then(function (h) {
        if (h !== 'ok') tampilkanForm(produk, h);
      }).catch(function () { /* offline: biarkan */ });
    } else {
      kunciLayar();
      janji = cekServer(produk, token).then(function (h) {
        if (h === 'ok') { bukaLayar(); return true; }
        return tampilkanForm(produk, h);
      }).catch(function () {
        if (umur < OFFLINE_MAKS_MS) { bukaLayar(); return true; }
        return tampilkanForm(produk, 'Perlu internet untuk memeriksa lisensi.');
      });
    }
    sedangMenunggu[produk] = janji;
    return janji;
  }

  global.Lisensi = { pastikan: pastikan, idPerangkat: idPerangkat };

  // Mulai otomatis bila tag <script> punya data-produk
  var tag = doc.currentScript || (function () { var t = doc.getElementsByTagName('script'); return t[t.length - 1]; })();
  var produkOtomatis = tag && tag.getAttribute('data-produk');
  if (produkOtomatis) pastikan(produkOtomatis);
})(this);
