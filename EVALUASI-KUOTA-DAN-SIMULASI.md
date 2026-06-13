# Evaluasi Penggunaan Kuota & Simulasi Beban — SIPEKA V2

Analisis konsumsi sumber daya Google Apps Script per peran pengguna, beserta simulasi beban satu siklus penilaian.

- **Tanggal analisis:** 2026-06-13
- **Arsitektur:** Frontend GitHub Pages (statis) → Backend Apps Script Web App (`/exec`) → Database Google Sheets (9 sheet)
- **Asumsi akun:** Gmail gratis (consumer) — kuota lebih ketat daripada Workspace
- **Skala (default README):** ~600 sekolah, ~10.000 guru. Turunan: ~50 pengawas (≈12 sekolah/pengawas), 22 kecamatan (Deli Serdang), 1 Admin Dinas
- **Pola beban:** penilaian bulanan, submit **terkonsentrasi di akhir periode**

> ⚠️ **Catatan metodologis penting.** Permintaan awal menyebut tiga metrik: HTTP Request (`UrlFetchApp`), serta baca/tulis **Cache** dan **Properti**. Pemeriksaan kode menunjukkan:
> - **`UrlFetchApp` = 0** — backend tidak pernah memanggil layanan luar. Yang ada adalah panggilan **masuk** (browser → Web App) = *executions*.
> - **`PropertiesService` = 0** — tidak dipakai sama sekali; pengaturan disimpan di sheet `PENGATURAN`.
> - **Cache** hanya dipakai untuk **token sesi Admin Dinas** (1× tulis saat login, 1× baca tiap aksi). Peran lain = 0.
>
> Karena itu, atas persetujuan, analisis difokuskan ke **kuota efektif**: jumlah *executions* Web App + operasi baca/tulis `SpreadsheetApp` + `LockService` — di sinilah beban & risiko sebenarnya.

---

## 1. Inventaris biaya per aksi backend

Operasi dihitung sebagai: **E** = 1 eksekusi Web App, **R** = operasi baca Sheet (`getRange().getValues()`), **W** = operasi tulis Sheet (`appendRow`/`setValue`/`setValues`/`deleteRow`), **L** = ambil `LockService` (kunci global), **CR/CW** = baca/tulis Cache.

| Aksi (endpoint) | Pemicu | E | R | W | L | CR | CW | Catatan ukuran |
|---|---|---|---|---|---|---|---|---|
| `master` (GET) | Form KS load, Rapor load | 1 | 2 | 0 | 0 | 0 | 0 | `getMaster_`: SEKOLAH (~600) + AKUN_PENGAWAS |
| `gurusekolah` (GET) | Form Guru load, Rapor guru | 1 | 1 | 0 | 0 | 0 | 0 | baca **seluruh** MASTER_GURU (~10.000 baris) lalu difilter |
| `pengaturan` (GET) | Rapor/Admin load | 1 | 1 | 0 | 0 | 0 | 0 | baca SETTING (kecil) |
| `rapor` (GET) | Cetak satu | 1 | 1 | 0 | 0 | 0 | 0 | baca seluruh sheet KS/Guru terkait |
| `raporbanyak` (GET) | Cetak banyak | 1 | 1 | 0 | 0 | 0 | 0 | baca seluruh sheet |
| `loginuser` (POST) | Login pengawas/KS | 1 | 3 | 1–2 | 0 | 0 | 0 | `cariAkun_`(1) + `getMaster_`(2) + tulis kunci/last-login |
| `loginuser` (kecamatan) | Login admin kec. | 1 | 1 | 1–2 | 0 | 0 | 0 | tanpa `getMaster_` |
| `login` (POST, admin) | Login Admin Dinas | 1 | 1 | 0 | 0 | 0 | **1** | baca adminHash + **tulis token ke Cache** |
| `submitks` / `submitpaud` | Submit penilaian KS | 1 | ~5 | ~2 | 1 | 0 | 0 | `cariAkun_`(1)+`getMaster_`(2)+`adaIdKirim_`(1)+`perbaruiMaster_`(1R+1W)+appendRow |
| `submitguru` | Submit penilaian guru | 1 | ~2 | 1 | 1 | 0 | 0 | `cariAkun_`(1)+`adaIdKirim_`(1)+appendRow — **tanpa** `getMaster_` |
| `hasilkecamatan` (kunci) | Admin kec. buka data | 1 | ~4 | 0 | 0 | 0 | 0 | `cariAkun_`(1)+baca SDSMP+PAUD+GURU (3 sheet penuh) |
| `hasilkecamatan` (token) | Admin Dinas buka data | 1 | ~3 | 0 | 0 | 1 | 0 | `cekToken_`(CR) + 3 sheet penuh |
| `editks`/`editguru` | Koreksi nilai | 1 | ~2 | 1 | 1 | 0–1 | 0 | resolusi akses + baca baris + tulis baris |
| `deleteks`/`deleteguru` | Hapus baris | 1 | ~1 | 1 | 1 | 0–1 | 0 | `deleteRow` |
| `laporan` (GET) | Dashboard | 1 | 4 | 0 | 0 | 1 | 0 | baca SDSMP+PAUD+GURU penuh + pengaturan |
| `tarik` (GET) | Tarik Data (snapshot) | 1 | ~7 | 0 | 0 | 1 | 0 | **terberat**: 3 sheet penuh + master + MASTER_GURU (~10.000) + detail |
| `admingenerateakun` | Generate akun massal | 1 | ~3 | **N** | 0 | 1 | 0 | 1 `appendRow` **per akun baru** (bisa ratusan) |

### Temuan kunci dari inventaris
1. **`getMaster_()` dipanggil di hampir setiap jalur** — tiap login pengawas/KS **dan** tiap `submitks`. Untuk 600 submit KS, itu 600× baca penuh MASTER_SEKOLAH + AKUN_PENGAWAS. Tidak ada cache.
2. **`LockService.getScriptLock()` adalah kunci tunggal se-skrip** (bukan per-sheet). Artinya **semua** submit KS, submit guru, edit, dan hapus **berebut satu kunci yang sama**. `waitLock(10000)` = jika menunggu >10 detik → submit gagal (lalu masuk Outbox).
3. **`tarik` & `laporan` membaca seluruh sheet besar dalam satu eksekusi** → risiko mendekati batas runtime 6 menit dan ukuran respons saat data membesar.
4. **Cache & Properti nyaris tak dipakai** — padahal keduanya adalah alat ideal untuk meredam beban baca Sheet (lihat rekomendasi).

---

## 2. Biaya per pengguna (per peran)

Panggilan API per **page-load** (dari pelacakan `apiGet`/`apiPost` di tiap halaman):

| Halaman / peran | Saat load | Saat aksi |
|---|---|---|
| Form KS (Pengawas) | `master` (1E) | tiap submit: `submitks` (1E) |
| Form Guru (Kepala Sekolah) | `gurusekolah` (1E) | tiap submit: `submitguru` (1E) |
| Rapor (Publik) | `master` (1E) + `pengaturan` (1E) | `rapor`/`raporbanyak` (+`gurusekolah` utk guru) |
| Hasil Penilaian (Admin Kec.) | `hasilkecamatan` (1E) | tiap edit/hapus (1E) |
| Dashboard (Admin Dinas) | `laporan` (1E) | `tarik` (1E) |
| Admin (Admin Dinas) | `pengaturan` (1E) | `login`/save/generate/reset (1E masing-masing) |

**Cache & Properti per peran:** hanya **Admin Dinas** yang menyentuh Cache (token). Pengawas, KS, Admin Kecamatan, dan Publik = **0 Cache, 0 Properti**.

---

## 3. Simulasi beban — satu siklus penilaian (1 bulan)

### Parameter
| Variabel | Nilai |
|---|---|
| Sekolah (akun KS) | 600 |
| Guru | 10.000 (≈16,7/sekolah) |
| Pengawas | 50 (≈12 sekolah/pengawas) |
| Kecamatan (admin) | 22 |
| Admin Dinas | 1 |
| Penilaian KS/bulan | 600 (1/sekolah) |
| Penilaian guru/bulan | 10.000 (1/guru) |

### Estimasi executions & operasi per peran (per bulan)

| Peran | Skenario aktivitas | Executions | Baca Sheet (R) | Tulis Sheet (W) | Lock (L) | Cache |
|---|---|---:|---:|---:|---:|---:|
| Pengawas (50) | 1 login + 1 load + 12 submit KS | ~700 | ~3.250 | ~1.275 | 600 | 0 |
| Kepala Sekolah (600) | 1 login + 1 load + 17 submit guru | ~11.400 | ~22.800 | ~11.100 | 10.200 | 0 |
| Publik (rapor) | ~1.000 sesi cetak | ~3.500 | ~5.000 | 0 | 0 | 0 |
| Admin Kecamatan (22) | 1 login + 5 buka + 5 edit | ~240 | ~700 | ~145 | ~110 | 0 |
| Admin Dinas (1) | ~30 sesi (login+laporan+tarik) | ~90 | ~berat | 0 | 0 | CW≈30 / CR≈70 |
| **TOTAL / bulan** | | **≈15.900** | **≈32.000** | **≈12.600** | **≈10.900** | **≈100** |

> Angka di atas adalah **estimasi orde besaran**, bukan presisi — bergantung pada berapa kali pengguna memuat ulang halaman & mencetak. Submit (10.600/bulan) bersifat pasti karena terikat populasi.

### Pembanding kuota Apps Script (akun Gmail gratis)
| Batas | Nilai (consumer) | Relevansi SIPEKA |
|---|---|---|
| URL Fetch calls / hari | 20.000 | **Tidak relevan** — `UrlFetchApp` tak dipakai |
| PropertiesService baca/tulis / hari | 500.000 / 50.000 | **Tidak relevan** — tak dipakai |
| Runtime per eksekusi | **6 menit** | Risiko di `tarik`/`laporan` saat data besar |
| **Eksekusi simultan** | **30** | **Batas pengikat utama saat puncak** |
| Runtime trigger / hari | 90 menit | Tidak relevan — tak ada trigger |

> 🔎 **Koreksi atas catatan README.** README menyebut "kuota Apps Script gratis ±20.000 panggilan/hari". Itu sebenarnya kuota **URL Fetch** (yang nilainya 0 di sini). Apps Script **tidak** mempublikasikan batas keras "jumlah pemanggilan Web App per hari" untuk akun konsumen. Maka ~15.900 executions/bulan **bukan** masalah secara jumlah. Yang mengikat adalah **konkurensi (30 eksekusi simultan)** dan **biaya per-eksekusi**.

### Analisis beban puncak (kritis)
Karena submit **terkonsentrasi di akhir periode**, asumsikan ~75% dari 10.600 submit terjadi dalam 2–3 hari terakhir; hari terburuk ≈ 60% → **~6.360 submit dalam 1 hari**.

- **Throughput kunci global:** seluruh submit melewati satu `getScriptLock()`. Bila satu tulis (di dalam kunci: `adaIdKirim_` membaca 500 baris + `appendRow`) memakan ~0,5–1,5 detik, kapasitas berkelanjutan ≈ **2.400–3.600 tulis/jam**.
  - 6.360 submit tersebar merata 8 jam = ~795/jam → **aman**.
  - 6.360 submit menumpuk dalam ~2 jam (sore hari) = ~3.180/jam → **mendekati/menembus kapasitas** → mulai muncul `waitLock` timeout.
  - menumpuk dalam ~1 jam = ~6.360/jam → **kelebihan beban**, banyak submit gagal seketika.
- **Penyelamat desain:** submit yang gagal **tidak hilang** — masuk **Outbox** (IndexedDB) dan dikirim ulang otomatis; `idKirim` + cek duplikat mencegah baris ganda. Jadi dampaknya = *tertunda*, bukan *hilang*. Ini titik kuat arsitektur.
- **Eksekusi simultan 30:** lonjakan sesaat dari 650 pengguna aktif berpotensi melampaui 30 → sebagian request error → masuk Outbox.

**Kesimpulan beban:** Dalam pemakaian normal tersebar, SIPEKA **nyaman di bawah kuota** akun gratis. Risiko nyata hanya muncul pada **puncak ekstrem** (semua menilai di jam yang sama di hari terakhir), dan itupun **degradasi anggun** lewat Outbox, bukan kehilangan data.

---

## 4. Verdикt evaluasi

| Aspek | Nilai | Keterangan |
|---|---|---|
| Efisiensi kuota luar (UrlFetch/Properti) | ✅ Sangat hemat | Nol penggunaan |
| Pemakaian Cache | ⚠️ Minim | Hanya token admin; peluang besar untuk caching master data |
| Konkurensi tulis | ⚠️ Berisiko di puncak | Kunci global tunggal menjadi serializer semua tulis |
| Skalabilitas baca | ⚠️ O(n) per operasi | `getMaster_`/`gurusekolah`/`tarik` baca sheet penuh berulang |
| Ketahanan kegagalan | ✅ Sangat baik | Outbox + anti-duplikat `idKirim` + `LockService` |
| Mode offline | ✅ Baik | Snapshot IndexedDB mengurangi executions dashboard/rapor |

---

## 5. Rekomendasi (urut dampak)

1. **Cache master data di server.** Simpan hasil `getMaster_()` dan `getGuruSekolah_()` ke `CacheService` (TTL 5–30 menit). Memangkas ratusan–ribuan baca Sheet penuh, terutama pada login & submit massal. *Inilah penggunaan Cache yang seharusnya — saat ini terbuang.*
2. **Hilangkan `getMaster_()` dari jalur `submitks`.** Validasi binaan cukup dengan membaca 1 baris master berdasarkan NPSN, bukan seluruh sheet.
3. **Indeks `idKirim` & lookup akun.** `adaIdKirim_` (500 baris) dan `cariAkun_` (sheet penuh) bisa dipercepat dengan kolom indeks atau `CacheService`.
4. **Paginasi `tarik`/`laporan`.** Saat `GURU` mendekati puluhan ribu baris, pecah respons atau gunakan kolom ringkas agar tak menembus 6 menit / ukuran respons. (README sudah menyarankan pecah sheet per tahun ajaran di >50.000 baris.)
5. **Pertimbangkan kunci per-jenis** (mis. lock terpisah submit-KS vs submit-guru) bila puncak benar-benar padat — mengurangi kontensi pada satu kunci global.
6. **Perbaiki klaim kuota di README** ("20.000 panggilan/hari") agar tidak menyesatkan operator.

---

*Semua angka simulasi adalah estimasi berbasis asumsi yang dinyatakan di Bagian 3. Untuk presisi, ganti parameter populasi dengan angka nyata dari `data.xlsx` dan ukur durasi eksekusi aktual via Apps Script Executions log.*
