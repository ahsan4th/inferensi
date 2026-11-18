# INFERENSI Finalizer

Aplikasi Streamlit untuk **finalisasi artikel INFERENSI** (Departemen Statistika ITS) yang sudah diterima, sebelum diterbitkan. Aplikasi ini mengotomatiskan langkah editorial yang biasanya dikerjakan manual dan menghasilkan **file Word siap terbit** beserta laporan temuan untuk editor.

## Input
| Input | Keterangan |
|---|---|
| Artikel `.docx` | Naskah final yang sudah memakai template INFERENSI |
| Volume, nomor, bulan, tahun | Membentuk header, mis. `INFERENSI, Vol. 9(3), November. 2026. ISSN: 0216-308X (Print) 2721-3862 (Online)` |
| Riwayat OJS `.html` | Halaman submission OJS yang disimpan lewat *Save Page As → Webpage, HTML only* |
| DOI (opsional) | Diisi ke footer; bila kosong dipakai placeholder `10.12962/j27213862.v<vol>i<no>.xxxx` |

## Output tambahan
- `*_LOA.pdf` — Letter of Acceptance dari `templates/LOA.docx` (tanggal = keputusan Accept, nama penulis korespondensi, judul, Volume/Nomor/Tahun)
- `*_Receipt.pdf` — kuitansi pembayaran dari `templates/Receipt.docx` (tanggal kuitansi, nama & email penerima, Article ID OJS, judul, jumlah, tanggal pembayaran diterima)
- Semua data diisi otomatis dari naskah dan riwayat OJS, dan bisa diubah di tab **LOA & Receipt**. PDF dibuat dengan LibreOffice (`packages.txt`); bila LibreOffice tidak ada, LOA/Receipt diberikan dalam format Word.

> **Penting:** unggah `templates/LOA.docx` dan `templates/Receipt.docx` ke repositori (berkas biner tidak ikut dalam push otomatis). Tanpa berkas ini, template dapat diunggah langsung dari sidebar aplikasi.

## Output
- `*_final.docx` — Word siap terbit (field sitasi Mendeley, persamaan, dan gambar tetap utuh)
- `*_laporan.md` — daftar perbaikan otomatis dan temuan yang perlu ditinjau
- `*_references.ris` — daftar pustaka dari field Mendeley/Zotero (bila ada)

## Yang diperbaiki otomatis
- **Header & footer**: teks header terbitan, DOI, nama institusi footer; sisa header jurnal lain dibersihkan
- **Tanggal penting** dari riwayat OJS: *Received* = initial submission, *Revised* = unggahan revisi terakhir penulis sebelum diterima, *Accepted* = keputusan Accept Submission; kotak tanggal dilebarkan bila perlu
- **Lisensi**: tahun © disesuaikan dengan tahun terbit
- **Teks**: karakter tak terlihat (zero-width space), spasi ganda, spasi yang hilang di sekitar persamaan inline, koma desimal → titik (opsional), istilah hasil terjemahan yang aman (*data train → training data*, *residue → residual*, …)
- **Sitasi IEEE**: pemisah kelompok sitasi `[1] [3]` → `[1], [3]` (field Mendeley Desktop, Mendeley Cite, dan teks biasa)
- **Persamaan**: aksen pada simbol bersubskrip (Ŷₜ), simbol ∅ yang dimaksud φ, matriks 1-baris sel kosong, posisi nomor persamaan (kiri–kanan)
- **Referensi IEEE**: `no. <Bulan>` dihapus, `[Online]. Available: https://doi.org/…` → `doi: …`, kapitalisasi akronim (MEWMA, SARIMAX, …), kata sambung kecil di judul, tanda baca `.,`
- **Tata letak**: *keep with next* untuk judul bagian dan caption tabel
- **Dokumen**: terima track changes, hapus komentar, bersihkan metadata (pembuat, penyunting terakhir, Grammarly)

## Kesesuaian dengan template INFERENSI
- Highlight/sorotan teks dihapus; kotak teks, bingkai (frame) dan garis kotak paragraf di dalam tabel dibuka menjadi teks sel biasa
- Font & ukuran: isi Palatino Linotype 9 pt; subjudul Helvetica 9 pt; caption Helvetica 8 pt (label biru tebal); tabel dan referensi Palatino Linotype 8 pt; afiliasi Times New Roman 8 pt; email korespondensi 7,5 pt
- Warna: teks hitam; judul, judul bab dan label caption biru #00629B; bar header/footer biru #00629B dengan teks putih
- Paragraf isi, abstrak dan referensi rata kanan-kiri (justify)
- Section halaman: A4, margin atas/kanan/bawah 1,70 cm, kiri 2,49 cm, header/footer 0,76 cm, satu kolom; urutan bagian I–V dan REFERENCES diperiksa
- Opsional: unggah file template untuk menyalin header/footer persis

## Sitasi Mendeley
- Field Mendeley yang sudah ada (Mendeley Desktop maupun Mendeley Cite) tetap utuh dan dapat di-refresh
- Sitasi teks biasa `[n]` diubah menjadi field Mendeley (`ADDIN CSL_CITATION`) dengan data referensi tertanam, dan daftar pustaka dibungkus field `Mendeley Bibliography`, sehingga Mendeley dapat memperbarui sitasi dan daftar pustaka
- Properti dokumen Mendeley (gaya sitasi IEEE) dipertahankan saat metadata dibersihkan

## Yang dilaporkan untuk ditinjau editor
- Placeholder draf (`[seed]`, `00 October 2026`, `xx(x)`, cross-reference rusak)
- Ejaan (pyspellchecker + kamus istilah statistik), kata ganda, pola terjemahan mesin (*variety* untuk ragam, *at the time of-t*, “By … being …”)
- **Daftar semua persamaan** dalam teks linear + cek nomor ganda/hilang, kurung tak seimbang, bobot yang sama dipakai dua kali, ramalan yang memuat galat acak, persamaan yang diketik sebagai teks biasa, rujukan ke persamaan yang tidak ada
- Sitasi: nomor tidak ada di daftar pustaka, referensi tidak disitasi, urutan kemunculan pertama (IEEE), rujukan gaya *author-year*
- Referensi: judul huruf kapital semua, nama jurnal hilang, tahun hilang, DOI mencurigakan, nomor artikel di `no.`, nama penulis bukan inisial
- Gambar/tabel: nomor caption, rujukan ke gambar/tabel yang tidak ada, gambar mengambang, digit berlebihan
- Format template: judul ≤ 16 kata, abstrak ≤ 200 kata, kata kunci ≤ 5 dan urut abjad, baris *corresponding author*
- Konsistensi dengan OJS: judul, nama kontributor, kata kunci, abstrak

## Kompatibilitas
- Diterima: `.docx` (Word, LibreOffice, Google Docs), `.docm`/`.dotx` (otomatis diubah menjadi .docx, makro dibuang).
- Ditolak dengan pesan jelas: `.doc` lama, dokumen berpassword, *Strict Open XML*, file rusak/bukan Word.
- Setiap langkah berjalan terisolasi; hasil akhir divalidasi (semua XML terbaca, relasi lengkap, dapat dibuka python-docx, teks tidak hilang). Bila gagal, aplikasi beralih ke **mode aman** (hanya header, footer, tanggal, lisensi, metadata); bila tetap gagal, file asli dikembalikan.

## Menjalankan secara lokal
```bash
git clone https://github.com/ahsan4th/inferensi.git
cd inferensi
pip install -r requirements.txt
# untuk PDF LOA/Receipt: pasang LibreOffice (Linux: sudo apt install libreoffice-writer)
streamlit run app.py
```

Pengujian:
```bash
python tests/run_samples.py naskah.docx riwayat_ojs.html 10.12962/j27213862.v9i3.12345 hasil.docx   # satu naskah
python tests/test_compat.py folder_uji riwayat_ojs.html                                          # semua .docx di folder
python tests/test_app.py naskah.docx riwayat_ojs.html                                             # antarmuka Streamlit
```

## Deploy ke Streamlit Community Cloud
1. Buka https://share.streamlit.io → **Create app** → *Deploy a public app from GitHub*.
2. Repository `ahsan4th/inferensi`, branch `main`, main file `app.py`.
3. *Advanced settings* → Python 3.12, lalu **Deploy**. (Repo privat perlu memberi akses Streamlit ke GitHub.)
4. `packages.txt` otomatis memasang LibreOffice dan font pendukung untuk membuat PDF LOA/Receipt.

Panduan lengkap: `docs/Panduan_INFERENSI_Finalizer.docx`.

## Catatan
- Perbaikan otomatis sengaja dibatasi pada perubahan yang aman. Perubahan isi (rumus yang salah secara matematis, rujukan yang keliru, tabel yang tidak cocok dengan teks) hanya **dilaporkan**, tidak diubah.
- Jika field Mendeley di-*refresh* oleh penulis, perbaikan teks daftar pustaka akan kembali ke data di library Mendeley mereka.
- Naskah artikel tidak disimpan oleh aplikasi; semua pemrosesan terjadi di memori.
