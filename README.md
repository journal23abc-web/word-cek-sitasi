# Cek Sitasi untuk Word

Add-in Microsoft Word untuk memeriksa sitasi dan daftar pustaka **APA 7** dan **IEEE**:

- **Validator** — deteksi otomatis dan saran perbaikan yang bisa diterima satu klik (ala Paperpal).
- **Tautan** — sitasi bisa diklik untuk melompat ke entri daftar pustakanya; DOI/URL jadi tautan aktif.
- **Konversi** — APA 7 ⇄ IEEE dengan pratinjau.
- **Cari online** — verifikasi DOI, cari DOI yang hilang, dan cari referensi untuk sitasi yang belum ada di daftar pustaka (CrossRef).

Situs & panduan pemasangan: https://journal23abc-web.github.io/word-cek-sitasi/
Demo tanpa Word: https://journal23abc-web.github.io/word-cek-sitasi/word-addin/demo.html

## Struktur repo

```
index.html                 halaman depan (unduh manifest, panduan pasang)
engine.js                  mesin sitasi (dari cek-sitasi)
bibliography-engine.js     model & format bibliografi (dari cek-sitasi)
converter-engine.js        konversi antar gaya (dari cek-sitasi)
word-addin/                add-in Word — lihat word-addin/README.md
```

Ketiga file mesin harus tetap berada di akar repo; `word-addin/taskpane.html` memuatnya dari `../`.

## Tes

```
npm test
```

## Asal

Mesin pemeriksa berasal dari [journal23abc-web/cek-sitasi](https://github.com/journal23abc-web/cek-sitasi). Bila mesin di sana diperbarui, salin ulang `engine.js`, `bibliography-engine.js`, dan `converter-engine.js` ke repo ini.
