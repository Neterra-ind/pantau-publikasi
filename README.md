# Pantau Publikasi

Aplikasi untuk vendor jasa publikasi: catat setiap konten lewat link, pantau capaian terhadap target kontrak,
lalu unduh laporan Excel dan PowerPoint untuk klien tanpa mengolah pivot manual. Dibangun dari `PRD Pantau Publikasi.pdf` (Fase 1 / MVP).

> Data bawaan adalah **data simulasi** (klien dan akun fiktif). Data asli dimasukkan lewat tempel atau impor workbook.

## Demo online
**https://neterra-ind.github.io/pantau-publikasi/** — langsung bisa dicoba tanpa memasang apa pun.

Di GitHub Pages tidak ada server, jadi aplikasi menjalankan `scripts/server.py` yang sama **di dalam browser** (Pyodide).
Semua fitur jalan, termasuk unduh Excel dan PowerPoint. Bedanya dengan versi lokal:
- tanpa login (masuk otomatis sebagai admin demo);
- data simulasi dan perubahan Anda **hanya tersimpan di browser Anda** (tidak terkirim ke mana pun, tidak terlihat pengunjung lain);
  tombol **Mode demo · Reset data** di kanan atas mengembalikan data awal;
- pemuatan pertama ±10–30 detik (mengunduh Python ke browser).

Jangan memasukkan data klien sungguhan ke demo; pakai versi lokal di bawah untuk pekerjaan nyata.

## Menjalankan
1. Sekali saja: `pip install openpyxl python-pptx pillow`
2. Klik dua kali `jalankan.bat`, lalu buka http://localhost:8082
3. Masuk dengan akun di `data/akun_awal.txt` (admin dan operator). Ganti kata sandi di menu **Pengaturan**.

Server hanya menerima koneksi dari komputer ini. Data tersimpan di folder `data/` (cadangkan folder ini secara rutin).

## Peran
| Peran | Boleh |
|---|---|
| **Admin** | Semua yang dilakukan operator, plus membuat/mengubah/menghapus klien, mengatur target, aturan nama, dan pengguna |
| **Operator** | Memasukkan, mengubah, menghapus data postingan; tempel, impor, rapikan nama; mengunduh laporan |

Setiap pengguna hanya melihat klien yang ia ikuti (diatur admin di Pengaturan → Pengguna).

## Menu
| Menu | Isi |
|---|---|
| **Ringkasan** | Skor "X dari Y target tercapai", 4 kartu (publikasi, views, interaksi, ER) dengan target dan selisih bulan lalu, grafik laju kumulatif harian terhadap jalur target, tabel per platform, target vs realisasi, per kelompok dengan kuota, akun paling aktif, konten views tertinggi |
| **Data** | Tabel yang bisa diedit langsung (tersimpan otomatis ±1 detik; indikator Menyimpan / Tersimpan / Gagal). Filter platform, kelompok, akun, jenis, tanggal, perlu dicek, dan pencarian; baris total selalu terlihat dan mencakup semua halaman |
| **Target** | Target per periode atau bawaan klien: postingan per platform, total views, total interaksi, ER minimum (total & per platform), kuota per kelompok dan per akun. Bisa disalin dari bulan lalu |
| **Ekspor** | Laporan Excel, laporan PowerPoint, template workbook. Peringatan muncul bila masih ada baris perlu dicek |
| **Pengaturan** | Klien (nama, istilah kelompok mis. "DPC", warna, logo), aturan nama, pengguna, ganti kata sandi |

Pemilih **klien** dan **periode** (bulan) selalu ada di atas. Tombol **+ Periode** menambah bulan baru.

## Memasukkan data
- **Tempel link / baris Excel** (tab Data): tempel daftar link satu per baris, atau salin baris dari Excel (dengan atau tanpa baris judul).
  Platform dikenali dari domain; tanggal Instagram dan TikTok diisi dari link bila kosong; akun diisi bila link memuatnya.
  Angka "1.234", "22.100", "3,7k", "12rb", "1,2 jt" terbaca benar. Link yang sudah ada **diperbarui, bukan digandakan**
  (parameter pelacakan seperti `?utm_source=` atau `?igsh=` diabaikan).
- **Impor workbook (.xlsx)**: lembar dikenali dari nama (IG/Instagram, TikTok, Facebook, YouTube, Artikel/Berita), kolom dari judulnya.
  Hanya baris bertanggal di periode aktif yang masuk; tanggal salah tahun di bulan yang sama (mis. 26 Agustus 2028) dikoreksi dan dilaporkan.
  Impor ulang file yang sama tidak menggandakan data. Bisa juga lewat baris perintah:
  `python scripts/impor_workbook.py --cek "file.xlsx" 2026-08` (hanya memeriksa) atau `--klien <id>` untuk menyimpan.
- **Rapikan nama**: menyatukan "Kab.Kendal", "kab. Kendal", "Kab. Kendal" ke penulisan yang paling sering. Pratinjau jumlah sel tampil
  sebelum diterapkan. Aturan "nama lama = nama baku" (mis. *Kota Solo = Kota Surakarta*, *PDI P = PDIP*) disimpan per klien dan
  otomatis dipakai saat tempel dan impor berikutnya.

## Penanda "Perlu dicek" (8 jenis)
Link tidak valid · link ganda · tanggal di luar periode · tanggal kosong · views kosong · akun kosong · kelompok kosong ·
metrik janggal (interaksi > views, atau ER > 50%). Arahkan kursor ke angka oranye di kolom "!" untuk penjelasannya.
Jumlahnya tampil di tab Data dan di Ringkasan; satu klik membuka tabel yang sudah tersaring.

## Rumus
- **Interaksi** = like + comment + repost + share. **ER** = interaksi ÷ views, dihitung dari jumlah (bukan rata-rata ER per baris).
- **Berita** tidak punya views: dihitung sebagai publikasi saja, tidak ikut views total dan ER.
- **Status** periode lewat (dan semua target ER): Tercapai ≥ 100%, Hampir 80–99%, Tertinggal < 80%.
  **Bulan berjalan**: Tercapai ≥ 100%; selain itu dibandingkan dengan laju (capaian ÷ porsi hari berlalu):
  Sesuai laju ≥ 90%, Perlu dikejar 70–89%, Tertinggal < 70%. Target kosong = tidak dinilai.
- Selisih "vs bulan lalu" memakai periode kalender sebelumnya milik klien yang sama (bulan berjalan dibandingkan dengan bulan lalu penuh).

## Laporan
- **Excel**: lembar Ringkasan (realisasi berupa **rumus** yang membaca lembar data; sel target kuning bisa diisi; status berwarna),
  Per [Kelompok] dan Per Akun (COUNTIF/SUMIF), lembar Instagram/TikTok/Facebook/YouTube/Berita (link bisa diklik, filter otomatis,
  rumus membaca sampai baris 3000), dan Catatan Data. Nilai hasil rumus ikut tersimpan agar tampil benar di penampil yang tidak menghitung ulang.
- **PowerPoint** (16:9, warna & logo klien): sampul, ringkasan, pencapaian target, sebaran dan rincian per kelompok, detail tiap platform
  (grafik asli PowerPoint yang datanya bisa diedit; tiga konten teratas dengan link), berita, produksi konten, penutup.
- **Unduh data terfilter** (tab Data): persis baris yang tampil, baris "Filter: …", kolom catatan masalah, total `SUBTOTAL(109, …)`.

## Uji
`python -m unittest discover -s tests -v` — 28 uji: aturan link/tanggal/angka/nama, 8 penanda, status & laju, ER agregat,
tempel & dedupe, filter + total, batas akses per klien, dan **regresi sampel Agustus 2026** (774 publikasi; IG 348 posting,
5.228.263 views, 114.606 interaksi). Uji regresi membaca workbook di folder `Dashboard Pekerjaan` (atau variabel `PANTAU_SAMPEL`);
bila tidak ada, dilewati.

Selisih dengan PDF lama yang perlu dikonfirmasi ke penyusunnya: interaksi Facebook menurut rumus **28.083** (PDF 25.580, kolom bergeser),
total interaksi **159.149** (PDF 156.524). Aplikasi mengikuti rumus.

## Struktur
- `index.html`, `assets/app.js`, `assets/style.css`, `assets/vendor/` — tampilan
- `assets/demo.js` — mode demo: bila tidak ada server, memuat Pyodide (dari jsDelivr) dan menjalankan server di browser
- `scripts/inti.py` — **semua aturan bisnis** (fungsi murni): parser link, tanggal dari link, angka, nama, 8 penanda, ringkasan, status, filter
- `scripts/server.py` — server lokal & API · `scripts/penyimpanan.py` — berkas JSON
- `scripts/ekspor_xlsx.py`, `scripts/ekspor_pptx.py` — laporan · `scripts/impor_workbook.py` — impor & template
- `scripts/buat_simulasi.py` — data simulasi (`--ulang` untuk membuat ulang; ini juga membuat kata sandi baru)
- `data/klien.json`, `data/pengguna.json`, `data/periode/<klien>/<YYYY-MM>.json`, `data/logo/`
- `tests/` — uji unit, regresi, dan server

## Perbedaan dengan rencana teknis PRD
PRD menyarankan Next.js + Supabase (Postgres, RLS) di Vercel. MVP ini dibuat sebagai aplikasi lokal (Python + HTML) agar langsung bisa
dipakai tanpa akun cloud; logika bisnis sudah terpisah di `scripts/inti.py` sehingga mudah dipindah. Pembatasan akses per klien
dilakukan di server (setara RLS; diuji di `tests/test_server.py`). Baris ganda di workbook lama tetap disimpan apa adanya
(dan ditandai "link ganda") supaya jumlah sama dengan laporan lama; tempel dan impor ulang tidak pernah menambah duplikat.

Belum dibuat (Fase 2–3 PRD): akses pembaca untuk klien, unggah ekspor Insights/CSV, metrik otomatis lewat API resmi,
narasi bantuan AI, pengingat target, templat PPTX per klien.
