# Bank QA Portfolio: ParaBank

Pengujian aplikasi perbankan online **[ParaBank](https://parabank.parasoft.com/parabank)** (aplikasi demo dari Parasoft): registrasi, login, rekening, transfer, bill pay, riwayat transaksi, profil, dan pinjaman.

Tahap pertama proyek ini adalah **QA manual yang terhubung otomatis dengan Jira**. Pengujian tetap dilakukan manusia di Excel, tetapi pembuatan tiket bug, pelacakan status, traceability, dan laporan berjalan sendiri lewat GitHub Actions. Tahap berikutnya adalah otomasi UI dan API di repo yang sama.

**[Dashboard](https://rivaldimhr.github.io/bank-qa-portfolio/)** · **[Workbook Excel](manual/ParaBank-Manual-QA.xlsx)** · **[Template kosong](manual/Template-Manual-QA.xlsx)**

## Alur kerja

```
 Tester (manual)                         GitHub Actions (otomatis)                     Jira
 ───────────────                         ─────────────────────────                     ────
 Uji ParaBank, isi Excel        push     Baca workbook
 (Status, Current Result,  ────────────► Baris Failed tanpa Bug ID ──────────────────► Tiket Bug baru
  Severity, screenshot)                    (cek dulu tiket terbuka yang sama)            + screenshot
                                         Tulis Bug ID ke Excel  ◄──────────────────────  key tiket
                                         Tarik status tiket     ◄──────────────────────  To Do / Done
                                         Perbarui RTM + Bug Report, commit Excel
                                         Terbitkan dashboard (gh-pages)
 git pull, lihat "Perlu retest"  ◄────── Setiap hari 06.00 WIB status disinkronkan lagi
 Retest, ubah Status, push
```

## Isi workbook

| Sheet | Isi |
|---|---|
| Summary | Hasil per modul dan bug per severity (rumus Excel) |
| Test Plan | Tujuan, ruang lingkup, pendekatan, kriteria mulai/selesai, risiko |
| Requirements | 12 requirement (REQ-01 s.d. REQ-12) beserta story Jira-nya |
| RTM | Requirement → test case → hasil → bug. **Diisi otomatis** |
| Autentikasi, Rekening, Transfer, Bill Pay, Transaksi, Profil & Pinjaman | 48 test case (21 positif, 27 negatif), dirancang dengan equivalence partitioning, boundary value, state transition, dan skenario negatif |
| Bug Report | Daftar bug dengan status dan link Jira. **Diisi otomatis** |
| Test Data | Akun uji per sesi (ParaBank publik bisa di-reset) |
| How to use | Cara pakai singkat |

Kolom test case mengikuti format lembar kerja QA: Modul, Feature, Test Case, Test Case ID, Test Type, Priority, Requirement ID, Pre-Condition, Test Step, Test Data, Expected Result, Current Result, Status, Severity, Bug ID, Bug Status, Notes, Capture.

## Cara pakai sehari-hari

1. `git pull` dulu (bot ikut menyimpan perubahan ke workbook).
2. Buka ParaBank, buat akun uji baru (PB-AUT-01), dan catat di sheet **Test Data**.
3. Jalankan test case. Isi **Current Result** dan pilih **Status** (Passed / Failed / Blocked / Not Run).
4. Untuk yang **Failed**: pilih **Severity**, simpan screenshot ke `evidence/` (misalnya `evidence/PB-TRF-06.png`), lalu tulis path-nya di kolom **Capture**. Current Result dipakai sebagai judul tiket, jadi tulis sejelas mungkin.
5. Kalau beberapa test case gagal karena bug yang sama, isi **Bug ID** yang sudah ada secara manual supaya tidak dibuat tiket ganda.
6. Commit dan push. Dalam beberapa menit:
   - tiket Bug muncul di Jira, lengkap dengan langkah, expected/actual, dan screenshot;
   - Bug ID dan status Jira tertulis di Excel;
   - RTM, Bug Report, dan dashboard diperbarui.
7. Saat tiket sudah **Done**, dashboard menandai test case-nya **Perlu retest**. Jalankan ulang, ubah Status, lalu push.

## Setup sekali saja

**1. Jira**
- Proyek `PB` sudah berisi epic **PB-1 QA manual ParaBank** dan 12 story requirement (PB-2 s.d. PB-13, berlabel `REQ-01` s.d. `REQ-12`) lengkap dengan kriteria penerimaan dan ID test case. Nomor story juga tercatat di sheet Requirements.
- Buat API token di <https://id.atlassian.com/manage-profile/security/api-tokens>.

**2. GitHub repo → Settings → Secrets and variables → Actions**

| Jenis | Nama | Isi |
|---|---|---|
| Variable | `JIRA_BASE_URL` | URL situs Jira, misalnya `https://nama-situs.atlassian.net` |
| Variable | `JIRA_PROJECT_KEY` | Key proyek, misalnya `PB` |
| Secret | `JIRA_EMAIL` | Email akun Atlassian |
| Secret | `JIRA_API_TOKEN` | API token dari langkah 1 |

Tanpa variabel ini workflow tetap berjalan dalam **mode lokal**: RTM, Bug Report, dan dashboard diperbarui, tetapi tiket belum dibuat. Dashboard akan memberi tanda "Jira belum terhubung".

**3. GitHub Pages:** Settings → Pages → Source: *Deploy from a branch*, branch `gh-pages` (muncul setelah workflow pertama selesai).

## Menjalankan di laptop

```powershell
pip install -r requirements.txt
python -m pytest -q tests                 # unit test skrip, memakai Jira palsu
python scripts/sync.py --local            # tanpa Jira: RTM, Bug Report, dashboard di folder site/
```

Untuk tersambung ke Jira dari laptop, isi dulu variabel lingkungan yang sama dengan tabel di atas (`$env:JIRA_BASE_URL = "..."` dan seterusnya di PowerShell), lalu jalankan `python scripts/sync.py`.

## Cara kerja skrip

[`scripts/sync.py`](scripts/sync.py):

- Membaca semua sheet yang memiliki kolom **Test Case ID**, jadi modul baru cukup ditambahkan sebagai sheet baru dengan header yang sama.
- **Tidak membuat tiket ganda.** Sebelum membuat tiket, skrip mencari tiket terbuka berlabel `parabank-manual` + ID test case. Baris yang sudah punya Bug ID dilewati.
- Setiap bug baru dihubungkan (link *Relates*) ke story requirement-nya, yaitu story berlabel `REQ-xx`.
- Severity dipetakan ke priority Jira (Critical → Highest). Jika proyek tidak memakai field priority, tiket tetap dibuat tanpa priority.
- Screenshot di kolom Capture diunggah sebagai lampiran, hanya jika file berada di dalam repo.
- **Workbook hanya disimpan jika ada nilai yang berubah,** jadi tidak ada commit kosong setiap hari.
- Status yang salah ketik (misalnya "Lulus") ditampilkan sebagai peringatan di dashboard.

[`tests/test_sync.py`](tests/test_sync.py) menguji perilaku di atas dengan Jira palsu: pembuatan dan penggunaan ulang tiket, status Done yang memicu retest, perhitungan RTM, mode lokal, dan penolakan lampiran di luar repo.

## Batasan

- ParaBank publik dipakai bersama banyak orang dan datanya bisa di-reset. Kegagalan karena reset atau server lambat dicatat di Notes, bukan sebagai bug.
- Requirement tidak berasal dari dokumen resmi ParaBank. Requirement disusun dari perilaku aplikasi dan praktik umum aplikasi perbankan.
- Screenshot disimpan sebagai file di `evidence/`, tidak ditempel di Excel, karena gambar yang ditempel akan hilang saat workbook disimpan ulang oleh skrip.
