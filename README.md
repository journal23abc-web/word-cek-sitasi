# Cek Sitasi untuk Word

Add-in Microsoft Word untuk memeriksa sitasi dan daftar pustaka **APA 7**, **IEEE**, dan **Harvard** (umum / Cite Them Right dan gaya Emerald):

- **Validator** — deteksi otomatis dan saran perbaikan yang bisa diterima satu klik (ala Paperpal).
- **Tautan** — sitasi bisa diklik untuk melompat ke entri daftar pustakanya. Di Word 2016/2019 dibuat sebagai salinan tertaut memakai mesin "Tautkan Sitasi" versi web.
- **Metadata** — status DOI (CrossRef), Cari DOI, status Scopus, rentang tahun referensi, dan laporan metadata sebagai komentar Word.
- **Konversi** — antar APA 7, IEEE, Harvard (umum), dan Harvard (Emerald) dengan pratinjau.
- **Template** — jumlah kata Judul, Abstrak, Introduction, Method, Results and Discussion, Conclusion & total dengan aturan hitung Word, How to cite (rentang halaman, judul, penulis/volume/DOI), dan kata kunci alfabetis; hasilnya bisa dijadikan komentar Word.
- **Istilah** — konsistensi istilah & nama variabel: bentuk penulisan berbeda (tunggal/jamak, tanda hubung, huruf kapital), salah ketik dalam istilah, ejaan UK/US bercampur, singkatan (kepanjangan ganda, dua nama untuk instrumen yang sama, dipakai sebelum didefinisikan), kemungkinan sinonim nama variabel, dan kata kunci yang tidak muncul di naskah; bisa diseragamkan sekali klik atau dijadikan komentar Word.
- **Laporan** — hasil pemeriksaan dijadikan komentar Word + sorotan warna, langsung di dokumen atau di salinan: per tab (Template, Metadata) atau **laporan keseluruhan** yang menggabungkan Validator, Metadata (DOI, Scopus, tahun), Template, dan Istilah dalam satu dokumen beserta ringkasannya.
- **Aman untuk Zotero/Mendeley/EndNote** — sitasi & daftar pustaka buatan pengelola sitasi hanya dilaporkan, tidak pernah ditulis ulang.

Situs & panduan pemasangan: https://journal23abc-web.github.io/word-cek-sitasi/
Demo tanpa Word: https://journal23abc-web.github.io/word-cek-sitasi/word-addin/demo.html

## Struktur repo

```
index.html                 halaman depan (unduh manifest, panduan pasang)
engine.js                  mesin sitasi (dari cek-sitasi)
bibliography-engine.js     model & format bibliografi (dari cek-sitasi)
converter-engine.js        konversi antar gaya (dari cek-sitasi)
link-engine.js             penautan .docx (dari cek-sitasi, tanpa perubahan)
scopus-engine.js           pencocokan Scopus (dari cek-sitasi, tanpa perubahan)
word-addin/                add-in Word — lihat word-addin/README.md
```

Semua file mesin harus tetap berada di akar repo; `word-addin/taskpane.html` memuatnya dari `../`.

## Tes

```
npm test
```

## Asal

Mesin pemeriksa berasal dari [journal23abc-web/cek-sitasi](https://github.com/journal23abc-web/cek-sitasi). Mesin di repo ini memuat perbaikan Harvard (varian Emerald, inisial rapat penulis terakhir, tanpa koma Oxford, pemisah "and"/"&" dalam kelompok sitasi, aturan singkatan institusi hanya untuk APA 7) yang perlu ikut dibawa ke cek-sitasi. Bila mesin di sana diperbarui, salin ulang `engine.js`, `bibliography-engine.js`, `converter-engine.js`, `link-engine.js`, dan `scopus-engine.js` ke repo ini. (Opsional: salin juga `scopus_sources.json` dan `scopus_proceedings.json` agar data Scopus dimuat dari repo ini sendiri.)
