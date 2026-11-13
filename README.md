# Dashboard WOK LamWay

Dashboard operasional sales IndiHome untuk **WOK Lampung Barat & Way Kanan**: performa IO/RE/PS, performa SF, peta order, prioritas kunjungan, dan perencanaan kunjungan ODP harian (**DJP — Daily Journey Plan**).

- Web: https://dashbordlamway.vercel.app
- Repo: `Rabil25/dashbordlamway` (branch `main` → otomatis deploy ke Vercel)

Dokumen ini adalah **catatan aturan resmi**. Bila aturan bisnis berubah, ubah kodenya **dan** dokumen ini bersamaan.

---

## 1. Arsitektur

```
Browser (index.html + assets/)  ──GET /api/data (Bearer session)──▶  Vercel Function (api/data.js)
        │                                                                  │  Google Sheets API (API key)
        │                                                                  ▼
        └──auth_* / djp_* / deepdive_odp (session_token)──▶  Google Apps Script  ──▶  Spreadsheet
```

| Komponen | Isi | Deploy |
|---|---|---|
| `index.html` | Markup & template Alpine.js semua halaman | Push ke `main` → Vercel |
| `assets/css/` | `theme.css` (token & tema) → `layout.css` → `responsive.css` → `components.css` (urutan wajib) | Vercel |
| `assets/js/` | `config.js` (aturan, kolom, helper) → `app.js` (dashboard) → `djp.js` (modul DJP). Script klasik, urutan wajib | Vercel |
| `api/data.js`, `lib/sheets.js`, `lib/session.js` | API baca data sales (wajib login) | Vercel |
| `gas/Code.gs`, `gas/DJP.gs` | Login, user, DJP, deepdive | **Manual**: tempel ke editor Apps Script → Deploy → Manage deployments → Edit → *New version* |
| `data/odp.json` | Snapshot master ODP dari tab `ODPBot` | Vercel (ekspor manual) |

Library dari CDN (versi dikunci): Alpine.js 3.14.9, ECharts 5.4.3, Leaflet 1.9.4.

### Spreadsheet

| Nama | ID | Dipakai untuk |
|---|---|---|
| Monitoring Daily HH LPG | `1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM` | `Bulan_Ini` (order), `Update SF` (master SF), `Target`, `Users` (akun) |
| Monitoring DJP SF LPG | `1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ` | `DJP_Plan`, `DJP_Realisasi`, `ODPBot` (master ODP) |

---

## 2. Aturan wilayah (cluster) — FONDASI, jangan diubah sembarangan

Didefinisikan di `assets/js/config.js` (`CLUSTER`) dan sama di `api/data.js` / `gas/Code.gs` (`WILAYAH`).

| Cluster | Kabupaten | STO | TL |
|---|---|---|---|
| **Lampung Barat** | LAMPUNG BARAT, TANGGAMUS, PESISIR BARAT | LIW, KRU, KOT, TGP | Andres Welsen |
| **Way Kanan** | LAMPUNG UTARA, WAY KANAN | BKM, KBI, BDU | Agus Setiawan |

- Order masuk cluster berdasarkan kolom **VALIDASI KAB** (`DQ`) di `Bulan_Ini`.
- SF masuk cluster berdasarkan kolom **CITY** (`U`) di `Update SF`, dan hanya yang **Status = ACTIVE**.
- ⚠️ Catatan terbuka: di data ODP ada 232 ODP berlabel cluster *WAY KANAN* tetapi berada di kabupaten *LAMPUNG BARAT* (STO BKM). Aplikasi tetap memakai aturan kabupaten di atas sampai diputuskan lain.

---

## 3. Data order — tab `Bulan_Ini`

Aplikasi membaca range **`CE2:DT`**. Indeks dihitung dari `CE = 0` (lihat `C` di `config.js`).

| Kolom | Isi | Dipakai untuk |
|---|---|---|
| CE | Status Order | — |
| CG | STO | Filter STO, peta |
| CH | Kategori produk | Filter produk |
| CJ / CM / CP | Validasi IO / RE / PS | Baris dihitung bila kolom ini **tidak kosong** |
| CK / CN / CQ | Bulan IO / RE / PS (angka 1–12, **tanpa tahun**) | ⚠️ Tidak dipakai lagi untuk menentukan bulan (lihat §4) |
| CL / CO / CR | Tanggal IO / RE / PS (**hari** 1–31) | Filter rentang tanggal (date picker) |
| CS / CT | Kode SF / Nama SF | Performa SF |
| CU | Channel | Deepdive channel |
| CX | Status final | Hitung **Batal** (mengandung CANCEL/BATAL) |
| CY / CZ | Detail / kategori kendala | Halaman Kendala |
| DC | Rentang jam IO | Grafik jam |
| **DH / DI / DJ** | Tanggal IO / RE / PS format **`yyyymmdd`** | **Sumber periode tahun-bulan** (§4) |
| DM | ID Order | Kunci deduplikasi |
| DP | Duplikat | Hanya baris **`Unik`** yang dihitung |
| DQ | VALIDASI KAB | Filter cluster |
| DR | Koordinat `lat,lng` | Peta |
| DT | Speed | Deepdive speed |

> ⚠️ Menyisipkan atau menggeser kolom di `CE:DT` akan menggeser semua angka. Bila struktur kolom berubah, ubah `C` di `config.js`, `IDX_DATE_*` di `lib/sheets.js` dan `gas/Code.gs`.

---

## 4. Aturan periode (bulan + tahun)

- Periode ditulis **`YYYY-MM`** (contoh `2026-11`) dan diambil dari kolom **DH/DI/DJ** (`yyyymmdd`), **bukan** dari CK/CN/CQ.
- Saat data dimuat, `normalizePeriods()` (`config.js`) mengganti isi kolom Bulan setiap baris menjadi `YYYY-MM`, sehingga semua halaman otomatis membedakan tahun.
- Bulan sebelumnya untuk MoM: `prevPeriod()` → **Jan 2027 dibandingkan Des 2026**.
- Tanggal `18991230` (tanggal kosong Excel) dan tahun di luar 2000–2099 diabaikan.
- API & Apps Script masih menerima angka bulan lama (`11`) dan menganggapnya tahun berjalan.
- **Syarat data:** kolom DH/DI/DJ harus tetap berformat `yyyymmdd` setiap pembaruan.

### Rentang hari & bulan lampau
- Bulan berjalan: "hari berjalan" = tanggal hari ini.
- Bulan lampau: "hari berjalan" = jumlah hari bulan itu (DRR & outlook memakai bulan penuh).

---

## 5. Rumus KPI

| KPI | Rumus |
|---|---|
| IO / RE / PS | Jumlah baris `Unik`, cluster terpilih, periode = bulan terpilih, tanggal hari di rentang, kolom Validasi tidak kosong |
| Target PS | Tab `Target` (kolom `KABUPATEN`, `PS`), dijumlah per cluster |
| Target RE / IO | **RE = PS × 1,2**; **IO = PS × 1,44** (dibulatkan) |
| ACV | aktual ÷ target × 100% |
| DRR | PS ÷ hari berjalan |
| MoM | dibandingkan bulan lalu pada **rentang tanggal yang sama** |
| WoW | PS 7 hari terakhir vs 7 hari sebelumnya |
| PS/RE, RE/IO | rasio konversi × 100% |
| Outlook akhir bulan | PS berjalan + DRR × sisa hari (DRR bisa disimulasikan) |

### Tier SF (berdasarkan PS bulan lalu)

| Tier | PS bulan lalu | Target PS |
|---|---|---|
| Black | 0 | 5 |
| Bronze | 1–9 | 15 |
| Silver | 10–15 | 20 |
| Gold | 16–25 | 30 |
| Platinum | 26–50 | 50 |
| Diamond | > 50 | 65 |

Aksi: ekspektasi = target ÷ 30 × hari berjalan. PS ≥ ekspektasi → **On Track**; kurang ≤ 20% → **Couching**; kurang > 20% → **Counseling**.

### Prioritas kunjungan (per desa)
- Potensi = √port avail × (1 − okupansi) × (1 − PS/total port), diskalakan 0–100.
- Jarak = jarak ke kantor referensi terdekat di kabupaten yang sama (`OFFICES` di `config.js`).
- Kuadran memakai **median** skor & median jarak: Prioritas (tinggi+dekat), Strategis (tinggi+jauh), Cadangan (rendah+dekat), Skip (rendah+jauh).

---

## 6. DJP — Daily Journey Plan

### Alur
1. **Setup:** tanggal → kabupaten (sesuai cluster) → kantor referensi (opsional) → desa (multi) → SF (multi).
2. **Generate draft**
   - Kumpulan ODP = ODP di desa terpilih dengan **port avail > 0** (dari `data/odp.json`).
   - **Rotasi** (`djp_plan_history`): ODP yang **belum pernah di-plan** didahulukan, lalu yang **paling lama** tidak di-plan, lalu port avail terbanyak.
   - Diambil **4 ODP × jumlah SF**.
   - Urutan: jika **Kantor Referensi** dipilih → urut **jarak dari kantor** (terdekat dulu); tanpa kantor → urut lintang/bujur.
   - Dibagi berurutan **4 ODP per SF**. Bila ODP kurang, muncul peringatan dan SF terakhir mendapat lebih sedikit.
3. **Review:** ODP bisa dibuang dari draft. Draft tersimpan di browser 2 jam.
4. **Simpan** (`djp_plan_save`, satu request untuk semua SF)
   - Simpan = **GANTI**: plan lama untuk pasangan *(tanggal, SF)* yang sama dihapus dulu, lalu yang baru ditulis. Satu SF = satu plan per tanggal.
   - Dibungkus `LockService` (aman dari dua orang menyimpan bersamaan).
   - Baris tanpa tanggal `yyyy-MM-dd`, kode SF, atau ODP → **seluruh simpan ditolak**.
   - Layar baru menulis "tersimpan" **setelah server mengonfirmasi**; bila gagal, draft tetap ada.
5. **Hapus** (`djp_plan_delete`): per SF per tanggal, dengan dialog konfirmasi. Mengembalikan error bila tidak ada baris yang cocok.
6. **Realisasi:** unggah CSV dari Web Telkomsel → `DJP_Realisasi`, deduplikasi berdasarkan `SCHEDULE_ID`.
7. **DJP Daily:** gabungan plan + realisasi per tanggal. **List ODP:** penanda Assign (di-plan bulan ini) & Visit (ada realisasi bulan ini).

### Kolom `DJP_Plan`
`TGL | USERNAME_SF | NAMA_SF | ODP | CLUSTER | KAB | KEC | DESA | LAT | LNG | AVAI | DIBUAT_OLEH | DIBUAT_AT | TAHUN`

Sheets otomatis mengubah teks tanggal menjadi sel bertipe tanggal. **Semua perbandingan tanggal wajib lewat `djpNormTgl()`** di `DJP.gs`.

### Kolom `DJP_Realisasi`
`SCHEDULE_ID | TUGAS_ID | NAMA_SF | TGL | ODP | STATUS | CHECKIN | CHECKOUT | LAT | LNG | ODP_CONDITION | PROSPECT_ESTIMATE | NOTES | UPLOADED_BY | UPLOADED_AT`

> ⚠️ Masalah terbuka: data realisasi yang ada berisi `00:00.0` di kolom `TGL` (kemungkinan CSV sempat dibuka/disimpan di Excel), sehingga penanda Visit dan DJP Daily belum bisa mencocokkan. Unggah CSV **asli** tanpa dibuka di Excel.

---

## 7. Login & keamanan

- Login di Apps Script: password di-hash SHA-256 di browser, dicocokkan dengan tab `Users`. Sesi berlaku **24 jam** (kolom `G` token, `H` expiry).
- Kolom `Users`: `Username | Password | Role | Status | Nama Lengkap | Tanggal Daftar | (token) | (expiry)`. Role: `admin` / `user`. Status: `pending` → `active` / `rejected` (disetujui admin).
- **Semua data wajib login:**
  - `/api/data` menerima header `Authorization: Bearer <session>`; diverifikasi di Vercel dari tab `Users` (cache 60 detik).
  - Endpoint data Apps Script memeriksa `session_token`.
  - Token bersama lama (`lpgLampung2026`) **tidak berlaku lagi**. `DATA_TOKEN` di Vercel berisi nilai acak yang tidak dipakai.
- Logout, ganti password, dan user ditolak langsung mengakhiri sesi (cache sesi dihapus).
- Teks dari sheet yang tampil di pop-up peta di-escape (`esc()`).
- ⚠️ Kedua spreadsheet dibagikan lewat link (keputusan pemilik). Siapa pun yang punya link bisa membuka tab `Users`.

---

## 8. Cache & kecepatan

| Lapisan | Durasi |
|---|---|
| Data sales di server Vercel (memori) | 10 menit |
| Daftar bulan di server | 1 jam |
| Verifikasi sesi di Vercel | 60 detik |
| Verifikasi sesi di Apps Script | 30 menit |
| Data di browser (memori) | 15 menit — tombol **Refresh** melewati cache |
| Daftar SF (Plan DJP, localStorage) | 30 menit |
| Plan tersimpan (localStorage) | 5 menit |
| `data/odp.json` | 30 menit (header di `vercel.json`) |

Data sales **tidak** di-cache CDN publik karena wajib login. Permintaan pertama setelah deploy bisa lambat (±20 detik) karena cache masih kosong.

---

## 9. Komponen tampilan

Gaya bersama ada di `assets/css/components.css`. Untuk elemen baru:

- **Tombol:** `.btn` + `.btn-primary` / `.btn-success` / `.btn-danger` / `.btn-ghost` / `.btn-outline-brand`; ukuran `.btn-xs` / `.btn-sm` / `.btn-lg`.
- **Tabel:** `<table>` biasa sudah seragam; kolom angka beri class `num`.
- **Konfirmasi:** `await uiConfirm({ title, message, confirmText, danger: true })` — jangan pakai `confirm()`.
- **Notifikasi:** `showToast(pesan, 'ok' | 'err' | 'warn' | 'info')`.
- **Daftar pilihan:** `<button class="pick-row" :class="{ on }" :aria-pressed>`.
- Warna lewat token (`--brand`, `--text-1/2/3`, `--success`, `--danger`, …) di `theme.css`. Teks redup `--text-3` sudah lolos kontras WCAG AA. Jangan pakai warna indigo/ungu.

---

## 10. Cara deploy

1. **Frontend & API:** commit lalu push ke `main`; Vercel otomatis membangun.
   - ⚠️ Vercel (paket Hobby) **melewati commit dari penulis yang tidak dikenal**. Bila deploy tidak muncul, cek status commit di GitHub.
2. **Apps Script:** salin isi `gas/Code.gs` dan `gas/DJP.gs` ke editor Apps Script → Deploy → Manage deployments → Edit → **New version**. Push ke GitHub **tidak** mengubah Apps Script.
3. **Master ODP:** `data/odp.json` diperbarui manual dari tab `ODPBot` (skrip ekspor belum ada di repo).

---

## 11. Riwayat perubahan utama (Nov 2026)

- Muat data 15,6 → ±1–4 detik (init ganda, verifikasi paralel, DJP dimuat saat dibuka).
- DJP: simpan mengganti (tidak dobel), hapus berfungsi, lock, status simpan jujur, kantor referensi aktif, rotasi aktif kembali.
- Semua data wajib login; token bersama dicabut.
- Periode tahun-bulan dari DH/DI/DJ (siap 2027).
- Tampilan mobile, kontras WCAG, fokus keyboard, komponen seragam.
- `index.html` dipecah menjadi `assets/css/*` dan `assets/js/*`.
- Menu **Timeline Desa** (`assets/js/timeline.js`): plan bulan berjalan per tanggal, warna kategori desa dari rumus Prioritas (`klasifikasiDesa()`, dihitung per cluster tanpa filter kab). Data plan via `/api/data?type=djp_plan` (`lib/plan.js`, read-only, wajib login).
- **Plan DJP jadi wizard 4 langkah** (Setup → Generate → Review dengan peta → Publish). Aturan generate dan simpan tidak berubah.
