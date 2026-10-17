# Cek Sitasi — Add-in Microsoft Word

Panel samping di Word yang memakai mesin **cek-sitasi** yang sama dengan versi web, dengan tiga fitur:

| Fitur | Yang dilakukan |
|---|---|
| **Validator** | Membaca dokumen yang sedang terbuka, langsung mendeteksi masalah sitasi & daftar pustaka **APA 7** / **IEEE**, lalu menampilkan kartu saran ala Paperpal: teks lama dicoret → teks baru, tombol **Terima / Abaikan / Lihat**, dan **Terima semua**. Diperiksa ulang otomatis saat dokumen diedit. |
| **Tautan** | Menautkan setiap sitasi di teks ke entri daftar pustakanya (bookmark + hyperlink internal; Ctrl+klik untuk melompat). DOI/URL di daftar pustaka dijadikan tautan aktif. Warna teks sitasi tetap seperti aslinya. |
| **Konversi** | APA 7 ⇄ IEEE: sitasi di teks diganti, daftar pustaka ditulis ulang (urutan, penomoran, cetak miring, indentasi gantung). Ada pratinjau sebelum diterapkan. |

Plus **pencarian online (CrossRef)** yang berjalan di latar belakang:

- DOI diverifikasi. CrossRef **tidak dijadikan acuan kebenaran**: perbedaan kecil (huruf kapital, subjudul, tahun terbit daring vs cetak, aksen nama) dianggap cocok, karena data di situs penerbit sering lebih baru. Hanya DOI yang tampaknya milik karya lain yang ditandai — tanpa penggantian otomatis, dengan tautan ke halaman penerbit untuk dicek.
- Referensi artikel/prosiding tanpa DOI dicarikan DOI-nya → tombol **Tambahkan DOI**.
- Sitasi yang tidak ada di daftar pustaka → **Cari referensi online** (penulis + tahun + kata kunci dari kalimat sekitarnya) → **Sisipkan ke daftar pustaka** sudah terformat APA 7/IEEE di posisi yang benar.
- Saat konversi, hanya data yang **belum ada** yang dilengkapi dari CrossRef (nama jurnal lengkap bila masih singkatan, DOI, penulis bila terpotong *et al.*, volume/halaman yang kosong). Setiap entri yang dilengkapi diberi tanda di pratinjau.

## Laporan konversi

Pratinjau konversi menampilkan **Perlu perhatian**: setiap entri yang tidak terbaca, entri yang gagal dibaca dengan benar, sitasi yang tidak cocok, dan entri yang perlu dicek — lengkap dengan alasannya dan tombol **Lihat** untuk melompat ke lokasinya di dokumen. Entri bermasalah tidak menghentikan konversi: entri itu dibiarkan apa adanya dan disorot kuning, sedangkan sisanya tetap dikonversi.

## Jenis saran yang dideteksi

**Sitasi di teks** — `and` vs `&`, bentuk *et al.*, urutan alfabetis dalam satu kurung, sitasi tanpa referensi, nomor IEEE yang tidak ada, kelompok IEEE (`[2, 4]` → `[2], [4]`, `[5]-[6]` → `[5], [6]`), gaya campuran.

**Daftar pustaka** — tanda pisah en (`1-20` → `1–20`), penghubung penulis, format DOI (`doi:` → `https://doi.org/` untuk APA), tanda kutip judul & titik akhir IEEE, label `[n]`, urutan nama IEEE, cetak miring nama jurnal/volume/judul buku (dibaca langsung dari format Word), urutan alfabetis APA (**Urutkan daftar referensi**), urutan kemunculan IEEE (**Nomori ulang otomatis**), referensi yang tidak disitasi.

Pengaman: saran tidak diterapkan otomatis bila akan menghilangkan kata dari referensi asli (mis. nama penulis yang gagal terbaca) — kartu menjelaskan bagian yang berisiko dan meminta perbaikan manual. Saat konversi, entri seperti itu ditulis apa adanya dan disorot kuning.

## Memasang

Add-in ini adalah halaman web statis yang di-host di GitHub Pages repo ini: `https://journal23abc-web.github.io/word-cek-sitasi/`.

1. Pastikan akar repo berisi `engine.js`, `bibliography-engine.js`, `converter-engine.js`, dan folder `word-addin/`. Tunggu GitHub Pages selesai memperbarui (± 1 menit).
2. Pasang `word-addin/manifest.xml` di Word:
   - **Word di web**: Beranda → **Add-in** → **Add-in Lainnya** → **Add-in Saya** → **Unggah Add-in Saya** → pilih `manifest.xml`.
   - **Word Windows**: taruh `manifest.xml` di folder bersama (mis. `\\NAMA-PC\addins`), daftarkan di *File → Opsi → Pusat Kepercayaan → Pengaturan Pusat Kepercayaan → Katalog Add-in Tepercaya* (centang *Tampilkan di Menu*), mulai ulang Word, lalu *Sisipkan → Add-in Saya → Folder Bersama*.
   - **Word Mac**: salin `manifest.xml` ke `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (buat foldernya bila belum ada), mulai ulang Word, lalu *Sisipkan → Add-in Saya*.
3. Tombol **Cek Sitasi** muncul di tab **Beranda**.

Host lain: `node word-addin/tools/make-manifest.js https://alamat-anda/ manifest.xml` (alamat = akar repo, wajib HTTPS).

### Mencoba di komputer sendiri (localhost)

```
npx office-addin-dev-certs install      # sekali saja, membuat sertifikat HTTPS lokal
node word-addin/tools/serve.js          # https://localhost:3000
```
Lalu pasang `word-addin/manifest.localhost.xml` dengan cara yang sama seperti di atas.

### Demo tanpa Word

Buka `word-addin/demo.html` (lewat server mana pun, mis. GitHub Pages atau `serve.js`). Halaman ini menampilkan dokumen tiruan di kiri dan panel add-in di kanan, dengan contoh naskah APA 7 dan IEEE. Semua tombol berfungsi pada dokumen tiruan itu.

## Kompatibilitas

| Kemampuan | Butuh | Catatan |
|---|---|---|
| Validator, Konversi, Terima saran | WordApi 1.3 | Word 2016+ / Microsoft 365 / Word di web / Mac |
| Tautan (bookmark), Track Changes | WordApi 1.4 | Word 2019+ / Microsoft 365 |
| Periksa otomatis saat mengetik | WordApi 1.6 | Versi lama: diperiksa ulang saat pindah kursor |
| Garis bawah saran di dokumen + menu Terima di popup | WordApi 1.7 / 1.8 | Hanya Word dengan langganan Microsoft 365; versi lain cukup lewat panel |

## Privasi

Teks naskah diproses di perangkat Anda. Saat **cek online** aktif (bisa dimatikan di *Pengaturan*), hanya judul/penulis/tahun/DOI referensi yang dikirim ke `api.crossref.org`.

## Struktur

```
word-addin/
  manifest.xml, manifest.localhost.xml   manifest add-in (dibuat oleh tools/make-manifest.js)
  taskpane.html / .css / .js             UI panel
  addin-core.js                          logika murni: struktur dokumen, saran, rencana tautan/konversi, CrossRef
  word-bridge.js                         semua panggilan Office.js (baca, cari, ganti, bookmark, anotasi)
  mock-bridge.js, demo.html              mode demo tanpa Word
  commands.html                          function file untuk tombol ribbon
  assets/                                ikon
  tools/                                 make-manifest.js, serve.js
  test/addin-core.test.js                tes Node (npm run test:addin)
```

## Keterbatasan

- Sama seperti versi web: deteksi berbasis pola teks. Selalu baca saran sebelum menerima.
- Sitasi/referensi yang dikelola Zotero/Mendeley/EndNote (field) bisa ditimpa ulang oleh pengelola sitasinya saat "Refresh". Untuk dokumen seperti itu, perbaiki di pengelola sitasinya atau lepas field-nya dulu.
- Konversi menulis ulang paragraf daftar pustaka (cetak miring dipertahankan sesuai gaya target; format karakter lain di dalam entri tidak dibawa).
- Hasil CrossRef bergantung pada keterindeksan sumber; "tidak ditemukan" bukan berarti referensinya salah.
