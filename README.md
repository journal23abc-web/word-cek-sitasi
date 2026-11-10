# Template Paste — Add-in Word

Tempel isi artikel per bagian (Introduction, Method, Results and Discussion, Conclusion, References, …), lalu add-in menyesuaikan formatnya dengan template jurnal secara otomatis:

- **Tempel keseluruhan**: salin seluruh naskah sekali (Ctrl+A, Ctrl+C), tempel sekali, dan isinya otomatis dipecah ke ABSTRACT, INTRODUCTION, METHOD, … REFERENCES menurut judul bagian (Heading 1, huruf KAPITAL; boleh bernomor seperti "1. INTRODUCTION" atau "IV. METHOD"). Judul, penulis, dan afiliasi di awal naskah dilewati; tabel halaman depan (Article info | Abstract) dibaca; RESULTS dan DISCUSSION yang terpisah digabung ke RESULTS AND DISCUSSION sebagai sub-heading; judul KAPITAL lain di badan artikel (mis. LITERATURE REVIEW) dijadikan sub-heading, sedangkan bagian penutup yang tidak ada di template (Conflict of Interest, Funding, Appendix, …) dilewati. Hasilnya bisa diperiksa dan diubah di **Tempel bagian** sebelum disisipkan
- **Font, ukuran, perataan, spasi sebelum/sesudah, spasi baris** untuk tiap jenis paragraf
- **Deteksi sub-heading otomatis** (kalimat pendek tanpa titik, teks bold, heading dari Google Docs/ChatGPT, heading bernomor seperti `2.1`) → dibuat bold sesuai template. Di Results and Discussion, *Results* dan *Discussion* jadi tingkat 2, sub-judul di bawahnya tingkat 3
- **Referensi**: hanging indent sesuai template, gaya **APA** atau **IEEE bernomor** ([1] …). Italic otomatis (APA: nama jurnal + volume / judul buku; IEEE: nama jurnal / judul buku) bila tempelan tidak membawa italic, DOI jadi hyperlink (APA) atau dibiarkan "doi: …" (IEEE), baris terputus dari PDF disambung
- **Caption** (`Table 1.` / `Figure 1.` → label bold, judul normal; caption gaya APA dua baris "Figure 1" + judul italic otomatis digabung, dan italic pada judul caption dihilangkan), **Note.**, dan **tabel** yang ditempel (garis gaya APA, font 10 pt)
- **Menambah, bukan mengganti**: isi masuk di bawah konten yang sudah ada pada bagian itu. Bila judul bagian belum ada, judulnya dibuat di urutan yang benar
- **Abstrak**: kartu ABSTRACT mengganti teks di kotak abstrak (bukan menambah). Abstrak terstruktur (*Purpose – / Methodology – / Findings – …*) dikenali otomatis dari tempelan (juga *Design/methodology/approach:*, *Tujuan:*, dll.); labelnya dibuat **bold** dan pemisahnya diseragamkan sesuai template. Abstrak satu paragraf diformat sesuai ukuran dan perataan template. Baris *Keywords* di tempelan dilewati. Abstrak lama bisa dikembalikan, dan jumlah kata dibandingkan dengan batas abstrak jurnal
- **Cek kesesuaian** (tab **Cek**): mencocokkan *To cite this article*, baik di badan dokumen, tabel, maupun footer, dengan artikelnya. Yang dicek: penulis (nama, urutan, inisial), judul, nama jurnal, volume/nomor, tahun, halaman (dibandingkan dengan nomor halaman awal dokumen), DOI (format, kelengkapan, kecocokan vX iY dan DOI di header), serta tanda baca (titik ganda, spasi ganda, placeholder). Juga memeriksa **article history**: tanggal terbaca dan valid, urutan Received ≤ Revised ≤ Accepted, dan jarak Received → Accepted **maks. 3 bulan** (bisa diubah per jurnal), tidak melewati bulan terbit, serta tanggal di masa depan. Selain itu memeriksa header/footer (placeholder "Author", "Title…", "XXXX"; running head penulis dan judul; nama jurnal) dan batas kata
- **Hapus isi bagian**: kosongkan seluruh isi satu bagian di dokumen (atau semua bagian sekaligus) dengan judul bagian tetap ada, misalnya untuk mengganti Introduction dengan versi baru. Ada konfirmasi dan tombol **Kembalikan isi yang dihapus** (selama panel masih terbuka), dan Ctrl+Z tetap bisa dipakai
- **Rapikan semua** (di menu Tempel keseluruhan): rapikan semua bagian di dokumen sekaligus, dengan ringkasan per bagian
- **Referensi ganda dihapus** saat Rapikan di bagian References: dikenali dari teks yang sama (walau beda huruf besar/tanda baca/DOI vs tanpa DOI), DOI yang sama, atau penulis pertama + tahun sama dengan isi ≥ 85% mirip. Yang dipertahankan versi yang lebih lengkap (ber-DOI). Pada gaya **IEEE bernomor**, yang pertama dipertahankan, nomor referensi diurutkan ulang, dan **sitasi [n] di badan artikel (termasuk tabel) ikut disesuaikan**, misalnya [1], [3] → [1] dan [2]–[4] → [1]–[3]
- **Rapikan isi yang ada**: terapkan format template ke paragraf yang sudah terlanjur ditempel manual (italic tidak diubah), **termasuk tabel** di bagian itu: garis vertikal dihapus, garis gaya APA (atas, bawah judul kolom, bawah tabel), font & ukuran tabel, judul kolom bold, perataan sel. Baris rumus yang masih berupa teks juga diubah menjadi persamaan Word
- **Kutipan langsung** (mis. kutipan wawancara “…” Guru, S1) dikenali dan diformat sesuai APA 7 (lihat di bawah)
- **Rumus / persamaan** (tab **Rumus**): membuat, menyalin, dan menempel rumus sebagai **persamaan Word asli** (bisa diedit dengan Equation Editor), lengkap dengan nomor (1), (2) rata kanan. Lihat bagian *Rumus* di bawah
- **Hitung kata** per bagian dan total terhadap batas jurnal (aturan bisa diatur per jurnal), plus abstrak
- **Profil template**: semua angka bisa diubah, dan profil bisa diduplikat untuk jurnal lain (ekspor/impor JSON)

**Kutipan langsung / wawancara mengikuti APA 7** (Publication Manual 8.27 & 8.36), bawaan untuk semua jurnal dan bisa diubah per jurnal di Template → Kutipan:
- **≥ 40 kata** → blok menjorok **1,27 cm**, **tanpa tanda kutip**, **tidak italic**; identitas partisipan **dalam kurung setelah tanda baca akhir**, misalnya *… improve things going forward. (Internal Audit Practitioner, subsidiary company, translated from Bahasa Indonesia)*.
- **< 40 kata** → APA menulisnya di dalam kalimat dengan tanda kutip. Add-in tetap menampilkannya sebagai paragraf terpisah, tetapi **tanda kutipnya dipertahankan**: “…” (Peran, kode).
- Atribusi yang ditulis di **baris terpisah** atau **menempel setelah tanda kutip** (“…” Teacher, S1) otomatis dipindahkan ke dalam kurung di akhir kutipan. **Rapikan** juga membetulkan kutipan yang sudah ada di dokumen.

**Rumus / persamaan** (tab **Rumus** di menu kanan):
- **Ketik atau tempel** rumus dalam bentuk apa pun: LaTeX (`\frac{a}{b}`, `\beta_0`), teks biasa (`Y = β0 + β1X1 + ε`, `n = N/(1+Ne²)`), atau **salin langsung** dari Word, ChatGPT, Wikipedia, dan situs lain yang memakai KaTeX/MathJax/MathML. Pratinjau muncul langsung.
- **Sisipkan di kursor**: sebagai rumus di baris sendiri (bernomor) atau di dalam kalimat. Nomor bisa diisi manual atau klik **Auto** untuk melanjutkan nomor terakhir di dokumen. Nomor naik otomatis setelah tiap sisipan.
- **Ubah teks terpilih jadi rumus**: blok baris rumus yang masih berupa teks, misalnya `Y = a + bX (1)`, lalu klik. Nomor (1) dipertahankan.
- **Ambil rumus dari dokumen**: klik di dalam rumus Word, lalu klik tombol ini. Rumus muncul sebagai LaTeX dan bisa diedit, disisipkan ulang dengan format template, atau **disalin** sebagai **LaTeX**, **MathML**, atau **teks**. MathML yang disalin bisa langsung ditempel (Ctrl+V) di Word mana pun dan otomatis menjadi persamaan.
- **Pustaka rumus penelitian**: rata-rata, simpangan baku, Z-score, Slovin, Lemeshow, Cronbach's alpha, korelasi Pearson, AVE, CR, uji t, chi-square, Cohen's d, N-Gain, regresi linear berganda, MRA, R², regresi logistik, Fuzzy C-Means, dan jarak Euclidean.
- **Saat menempel isi bagian** (tab Tempel), rumus ikut terbawa sebagai persamaan Word: rumus dari Word (termasuk yang diletakkan di tabel `| rumus | (1) |`), ChatGPT (`\[ … \]`, `$$ … $$`), dan web. Baris rumus berupa teks, seperti `Y = β0 + β1X1 + ε (1)`, juga diubah. Rumus di dalam kalimat menjadi teks biasa dengan subskrip dan superskrip Unicode (R², β₀).
- **Rapikan isi yang ada** juga membetulkan rumus di bagian itu: baris rumus yang masih berupa teks biasa (mis. `Y = β0 + β1X1 + ε (1)`, termasuk subskrip/superskrip dari naskah) diubah menjadi persamaan Word bernomor sesuai template, dan nomor "(1)" yang terpisah di baris berikutnya digabung. Rumus Word yang sudah benar tidak diubah. Rumus berupa **objek MathType/Equation 3.0** atau gambar tidak bisa diubah otomatis; jumlahnya dilaporkan.
- **Penomoran per jurnal** (Template → Rumus / persamaan): *(1) rata kanan* (bawaan; lebar halaman dibaca dari dokumen), *(1) langsung setelah rumus*, *penomoran Word 365 (#)*, atau tanpa nomor.

**Batas kata diatur sendiri per jurnal** di tab **Hitung kata → Aturan hitung kata**: maksimal/minimal kata isi, maksimal kata abstrak, bagian mana yang ikut dihitung (mis. referensi, Author Contribution, AI Disclosure), dan teks tabel. 0 = tidak dibatasi.

---

## 1. Hosting (sekali saja, gratis lewat GitHub Pages)

Add-in Word harus dibuka dari alamat **https**. Cara termudah:

1. Buat repositori baru di GitHub, misalnya `template-paste` (boleh *Public*).
2. Unggah **seluruh isi folder ini** (`manifest.xml`, `README.html`, folder `src/`, `assets/`) ke repositori itu.
3. Buka **Settings → Pages**, lalu pada *Source* pilih **Deploy from a branch**, branch `main`, folder `/ (root)`, dan klik **Save**.
4. Tunggu 1–2 menit. Alamat situs Anda akan seperti `https://NAMAANDA.github.io/template-paste`.
5. Buka `manifest.xml` di Notepad, lalu gunakan **Find & Replace**: ganti semua `https://YOUR-SITE` dengan alamat tadi (tanpa `/` di akhir). Simpan file itu di komputer Anda. (File di GitHub tidak perlu diubah.)

Cek: buka `https://NAMAANDA.github.io/template-paste/src/taskpane.html` di browser. Panel akan tampil dengan tulisan "Mode pratinjau".

## 2. Memasang di Word

### Word di Windows (Microsoft 365 / 2019 / 2021)
1. Buat folder, misalnya `C:\AddinWord`, lalu salin `manifest.xml` ke folder itu.
2. Klik kanan folder → **Properties → Sharing → Share…** → tambahkan diri Anda → **Share**. Catat alamat jaringannya, misalnya `\\NAMA-PC\AddinWord`.
3. Di Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Isi *Catalog Url* dengan `\\NAMA-PC\AddinWord`, klik **Add catalog**, centang **Show in Menu**, klik OK, lalu tutup dan buka Word lagi.
4. **Home → Add-ins** (atau **Insert → My Add-ins**) → tab **SHARED FOLDER** → **Template Paste** → **Add**.

### Word di Mac
1. Buka Finder → **Go → Go to Folder…** → `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (buat folder `wef` bila belum ada).
2. Salin `manifest.xml` ke folder itu, lalu buka ulang Word.
3. **Insert → Add-ins → My Add-ins** → pilih **Template Paste**.

### Word di web (office.com)
**Insert → Add-ins → My Add-ins → Upload My Add-in** → pilih `manifest.xml`.

Setelah terpasang, tombol **Template Paste** muncul di tab **Home**.

---

## 3. Cara pakai

1. Buka naskah template jurnal (yang sudah ada judul INTRODUCTION, METHOD, dst., atau masih kosong pun bisa).
2. Klik **Template Paste** di tab Home, lalu pilih jurnalnya di menu **Jurnal**.
3. **Cara cepat**: di menu **Tempel keseluruhan**, klik kotak bergaris putus-putus lalu tekan **Ctrl+V** dengan seluruh naskah. Periksa ringkasan per bagian, lalu klik **Sisipkan semua ke dokumen** (atau **Periksa di Tempel bagian** untuk mengubah hasil deteksi lebih dulu). Isi selalu ditambahkan di akhir tiap bagian, jadi kosongkan dulu isi lama dengan **Hapus isi semua bagian…** bila dokumen sudah berisi naskah.
   **Per bagian**: di menu **Tempel bagian**, buka kartu bagian (misalnya *INTRODUCTION*), klik kotak bergaris putus-putus, lalu tekan **Ctrl+V**. Anda bisa menempel berkali-kali untuk menambah isi.
4. Periksa pratinjau. Tiap baris punya menu jenis (*Isi / Heading 2 / Heading 3 / Kutipan / Caption / Catatan / Referensi / Lewati*). Ubah bila deteksi keliru.
5. Klik **Sisipkan ke dokumen**, atau **Sisipkan semua bagian** di bawah untuk memasukkan semua sekaligus.
6. Bila hasilnya tidak sesuai, klik **Hapus sisipan terakhir** (atau Ctrl+Z).
7. Untuk **mengganti** isi suatu bagian: klik **Hapus isi bagian ini di dokumen** di kartu bagian itu, lalu tempel isi barunya. Judul bagian tidak ikut terhapus. Perhatikan bahwa *semua* isi di antara judul itu dan judul bagian berikutnya ikut terhapus, termasuk tabel, gambar, dan section break.

**Mode "Posisi kursor"**: isi disisipkan di posisi kursor, bukan di akhir bagian.

**Gambar** di tempelan diabaikan. Sisipkan gambar langsung di Word, lalu tempel caption-nya lewat add-in.

**Tips akurasi deteksi**: menyalin dari Word, Google Docs, atau ChatGPT paling akurat karena format bold/italic ikut terbawa. Untuk teks polos (Notepad/PDF), sub-heading dikenali dari ciri "pendek, diawali huruf kapital, tanpa titik di akhir".

## 4. Template untuk jurnal lain

Tab **Template** → **Duplikat**, beri nama (misalnya "Jurnal X"), lalu ubah ukuran, spasi, hanging, kutipan, dan **Daftar bagian artikel** (judul bagian beserta nama lainnya, misalnya *Pendahuluan*, *Metode Penelitian*). Profil tersimpan di komputer itu. Gunakan **Ekspor** untuk menyalin profil ke komputer lain.

## 5. Struktur file

```
manifest.xml          definisi add-in (ganti YOUR-SITE)
src/taskpane.html     panel
src/taskpane.js       antarmuka
src/engine.js         deteksi heading/referensi & aturan format (tanpa Office.js)
src/math.js           rumus: LaTeX/teks/MathML/OMML Word -> persamaan Word (OMML), MathML, LaTeX
src/word-ops.js       operasi ke dokumen Word (Office.js, WordApi 1.3+)
assets/               ikon
tests/                uji otomatis (Playwright)
```

Kebutuhan: Word 2016 ke atas / Microsoft 365 / Word Mac / Word web. Fitur "Hapus sisipan terakhir" butuh WordApi 1.4 (Microsoft 365 / Word 2021).

## 6. Catatan versi

**v11.2:**
- Perbaikan error *GeneralException (Document._GetObjectByReferenceId)* saat "Hapus isi semua bagian…" / "Hapus isi bagian ini" pada bagian paling akhir dokumen (biasanya REFERENCES).

**v11.1 (tabel):**
- Tabel yang ditempel atau dirapikan otomatis **AutoFit to contents** dan diletakkan di tengah. Butuh Word desktop versi baru (Microsoft 365); di versi lama lebar tabel dibiarkan dan add-in memberi tahu.
- Baris judul kolom diberi **Repeat as header row** (diulang di atas tiap halaman). Judul kolom dua baris (sel gabungan) dikenali, sehingga kedua barisnya diulang dan garis APA ada di bawah baris kedua.
- Bisa diatur di Template → Tabel yang ditempel: Lebar tabel (contents / window / biarkan), ulangi baris judul, posisi tabel.

**v11:**
- Rapikan di References menghapus referensi ganda (IEEE: nomor & sitasi di teks disesuaikan).
- Tombol **Rapikan semua** di menu Tempel keseluruhan.

**v10.1 (perbaikan error "GeneralException (Paragraph.insertOoxml)"):**
- Rumus yang disalin dari persamaan Word sebagai teks (huruf matematika 𝑖 𝑡 𝜀) kini dibaca utuh; sebelumnya hurufnya terpotong sehingga Word menolak rumusnya. `SHIP𝑖𝑡`, `Yit` → SHIPᵢₜ, Yᵢₜ; `1.58×10−7` → 1.58 × 10⁻⁷.
- Teks biasa di dalam rumus (`\text{…}`) ditulis sesuai aturan format rumus Word. Semua rumus yang dibuat add-in kini diperiksa terhadap aturan itu di uji otomatis.
- Bila Word tetap menolak satu rumus, rumus itu disisipkan sebagai teks berformat (subskrip/superskrip) dan **sisipan lain tetap berjalan**; jumlahnya dilaporkan.
- Baris definisi/daftar seperti "Yit = District-level economic growth …" atau "Volume Loading (VOLLOADING = …; p = …)" tidak lagi dianggap rumus atau sub-heading.
- Tempel teks polos dari Word: tabel (sel dipisah Tab) disusun kembali jadi tabel, termasuk sel yang terpotong ganti baris; "a. Maritime …" tidak lagi tergabung dengan judul di atasnya.
- Tempel keseluruhan: judul bagian yang diketik huruf biasa tetapi tampil KAPITAL karena format Word (gaya Heading / All Caps) juga dikenali.

**v10:**
- Menu baru **Tempel keseluruhan** (paling atas): satu kali tempel untuk seluruh artikel, dipecah otomatis per Heading 1 sesuai daftar bagian template jurnal. Menu tempel lama berganti nama menjadi **Tempel bagian**.
- Teks polos yang tiap paragrafnya satu baris panjang tidak lagi tersambung menjadi satu paragraf.

**v9.1 (rumus):**
- Rumus dari ChatGPT yang **backslash-nya hilang** saat disalin (rumus diapit `[` … `]` di baris sendiri, `(x_i)` di dalam kalimat, `\sigma!\left(`, `{0,1}`) kini dikenali, baik ditempel sebagai teks polos maupun dari Word (satu baris per paragraf). Rumus baris menjadi persamaan Word bernomor (dari `\tag{1}`), rumus dalam kalimat menjadi huruf miring dengan subskrip/superskrip asli.
- **Rapikan** juga memperbaiki rumus seperti itu yang sudah terlanjur ada di dokumen.
- Huruf khusus ℝ, 𝒟, ℒ (`\mathbb`, `\mathcal`), simbol ⊙ ⊕ ⊗ dan lainnya, `aligned` dengan baris rata di tanda "=", serta `\log`/`\sum` di dalam rumus panjang ditampilkan dengan benar.
- Kalimat pengantar rumus ("The gating vector is computed as") tidak lagi dianggap sub-heading; "Algorithm 3." dikenali sebagai caption.
- Tabel yang tersalin sebagai baris-baris teks (satu sel per baris, setelah "Table 4. …") disusun kembali menjadi tabel; jumlah kolom ditebak otomatis.
- Rumus tanpa pembuka "[" tidak lagi menelan kalimat sebelumnya; `1000,t` (aslinya `1000\,t`) diperbaiki.
- Heading gaya IEEE (IV. → A. → 1)) mendapat tingkat yang benar.

**v9:**
- Template jurnal baru.
- Daftar pustaka gaya **IEEE** bernomor ([1] …): nomor dipertahankan, nama jurnal / judul buku dimiringkan otomatis bila tempelan tanpa italic, "doi: 10.…" tidak diubah, dan referensi yang terputus (salinan PDF) disambung dengan benar. Atur di Template → Referensi → *Gaya referensi*.
- Ukuran judul bisa diatur per bagian (mis. judul bagian penutup 11 pt sementara judul utama 12 pt) — Template → Daftar bagian artikel.
- Catatan tabel: pilihan label "Note." tebal dan seluruh catatan miring.
- Deteksi sub-heading dari teks polos: heading yang langsung diikuti sub-heading (mis. "Theoretical Background" lalu "Social media usage motivation") kini dikenali keduanya.
- Cek: perbandingan halaman awal cite dengan penomoran dokumen diperbaiki.

**v8.1:**
- **Rapikan** sekarang mengubah rumus yang masih berupa teks menjadi persamaan Word (lihat bagian *Rumus*).

**v8:**
- Tab baru **Rumus**: buat rumus dari LaTeX atau teks biasa, salin dari Word/ChatGPT/Wikipedia, sisipkan sebagai persamaan Word asli bernomor, ubah teks terpilih jadi rumus, ambil rumus dari dokumen, dan salin sebagai LaTeX/MathML/teks. Dilengkapi pustaka 20 rumus penelitian.
- Tempel isi bagian: rumus dari Word, ChatGPT, dan web tidak lagi hilang atau menjadi gambar, tetapi masuk sebagai persamaan Word. Baris rumus berupa teks ikut diubah (bisa dimatikan di Template → Rumus / persamaan).
- Pengaturan rumus per jurnal: penomoran, posisi, ukuran, dan spasi.
- Setelah memperbarui file di GitHub, panel otomatis memuat versi baru (tidak perlu hapus cache).

**v7:**
- Kutipan wawancara mengikuti **APA 7** untuk semua jurnal (lihat di atas).
- Baris persamaan (Y = β0 + …) dikenali dan tidak dianggap sub-heading.
- Sub-heading "Theme 1: …" yang panjang dikenali.
- Cek: article history dalam satu baris "(Received: …; Revised: …; Accepted: …)", rentang halaman di header ("Page 57-71") dibandingkan dengan penomoran dokumen, dan *cite* tidak wajib untuk jurnal yang tidak memakainya.

**v6.1:**
- Perbaikan error *GeneralException (TableCell.cellIndex)* saat mengganti/mengosongkan abstrak di tabel yang punya sel gabungan. Add-in tidak lagi memakai nomor baris/kolom sel; sel abstrak dikenali dari isinya.
- Rapikan tabel juga tidak lagi memakai nomor sel, jadi aman untuk tabel bersel gabungan.

**v6 (perbaikan bug):**
- Abstrak yang disalin dari dalam tabel sekarang terbaca. Kalau seluruh tabel halaman depan ikut tersalin, *Article info*, *Keywords* beserta butirnya, *Article history*, *cite*, dan lisensi otomatis dilewati.
- Mengganti abstrak yang terdiri dari beberapa paragraf memakai ulang paragraf yang ada, sehingga tidak error di sel tabel dan latar warnanya terjaga.
- Hitung kata abstrak tidak lagi ikut menghitung teks *To cite this article*.
- Panel tetap aktif walaupun Word lambat siap saat pertama dibuka.
- Kompatibel dengan Word for Mac versi lama (tanpa regex *lookbehind*).
- Rapikan: label caption ("Table 1.") dibold, baris persamaan "(1)" tidak diubah.
- Menyisipkan setelah daftar bernomor/berpoin tidak membuat isi baru ikut bernomor.
- Catatan kaki/komentar dari clipboard Word tidak ikut tertempel.
- "Figure 1" yang berdiri sendiri dikenali sebagai caption, bukan sub-heading.
