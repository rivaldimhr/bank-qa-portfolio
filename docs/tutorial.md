# Tutorial: QA manual ParaBank yang terhubung ke Jira

Tutorial ini menjelaskan cara memakai repo ini dari nol: menyiapkan laptop, menjalankan test case, melaporkan bug, sampai retest. Di bagian akhir ada tiga latihan untuk mencoba seluruh alurnya.

## 1. Gambaran besar

Ada tiga pihak yang bekerja bergantian:

| Pihak | Tugas |
|---|---|
| **Kamu (tester)** | Menguji ParaBank di browser, mengisi Excel, lalu push |
| **GitHub Actions** | Membaca Excel, membuat tiket bug, menulis balik Bug ID dan status, menerbitkan dashboard |
| **Jira (proyek PB)** | Menyimpan requirement (story PB-2 s.d. PB-13) dan bug (PB-14 dan seterusnya) |

```
KAMU                          GITHUB ACTIONS (otomatis)          JIRA
1. git pull
2. Uji ParaBank
3. Isi Excel
4. Commit & push  ─────────►  5. Baca Excel
                              6. Ada Failed baru? ────────────► 7. Tiket Bug dibuat
                              8. Tulis Bug ID ke Excel ◄───────
                              9. Update RTM, Bug Report,
                                 dashboard; commit Excel
10. git pull
                                                                11. Bug diperbaiki → Done
                              12. Tiap pagi: tarik status ◄────
                                  → dashboard "Perlu retest"
13. Retest → Passed → push
```

Prinsipnya: **kamu hanya menguji dan mengisi Excel.** Membuat tiket, menyalin langkah reproduksi, mengecek status, dan menyusun laporan dikerjakan sistem.

## 2. Persiapan (sekali saja)

1. Pasang [Git](https://git-scm.com/downloads) dan Microsoft Excel.
2. Clone repo:
   ```powershell
   git clone https://github.com/rivaldimhr/bank-qa-portfolio.git
   cd bank-qa-portfolio
   ```
3. Siapkan tiga tab di browser:
   - ParaBank: <https://parabank.parasoft.com/parabank>
   - Jira proyek PB (board atau daftar issue)
   - Dashboard: <https://rivaldimhr.github.io/bank-qa-portfolio/>

Konfigurasi Jira di GitHub (secrets dan variables) sudah selesai, jadi tidak perlu diulang.

## 3. Mengenal workbook

Buka `manual/ParaBank-Manual-QA.xlsx`.

| Sheet | Kamu isi? | Fungsi |
|---|---|---|
| Summary | Tidak | Ringkasan otomatis (rumus) |
| Test Plan, Requirements | Jarang | Rencana uji dan 12 requirement beserta story Jira-nya |
| Autentikasi … Profil & Pinjaman | **Ya** | 48 test case, tempat mencatat hasil |
| RTM, Bug Report | **Jangan** | Diisi otomatis oleh GitHub Actions |
| Test Data | **Ya** | Akun uji yang dibuat setiap sesi |

Kolom yang kamu isi di sheet modul:

| Kolom | Diisi kapan | Contoh |
|---|---|---|
| Current Result | Setiap test | `Muncul konfirmasi transfer 100 dari 13344 ke 13455` |
| Status | Setiap test | `Passed` / `Failed` / `Blocked` / `Not Run` |
| Severity | Hanya jika Failed | `Critical` / `High` / `Medium` / `Low` |
| Capture | Jika ada screenshot | `evidence/PB-TRF-06.png` |
| Notes | Bebas | `Server lambat, diulang 2x` |
| Bug ID | **Biasanya tidak** | Diisi bot. Isi manual hanya jika bug yang sama sudah punya tiket |
| Bug Status | **Tidak** | Diisi bot dari Jira |

Isi kolom Device dan Tanggal tes di bagian atas setiap sheet modul yang kamu kerjakan.

### Panduan memilih Status

- **Passed**: hasil sesuai Expected Result.
- **Failed**: hasil berbeda dari Expected Result, dan penyebabnya aplikasi.
- **Blocked**: test tidak bisa dijalankan karena hal lain, misalnya server ParaBank down atau test sebelumnya gagal. Tulis alasannya di Notes.
- **Not Run**: belum dijalankan.

### Panduan memilih Severity

| Severity | Arti | Contoh di aplikasi bank |
|---|---|---|
| Critical | Uang atau data bisa salah / hilang | Transfer negatif menambah saldo |
| High | Fitur utama rusak atau validasi penting tidak ada | Transfer melebihi saldo diterima |
| Medium | Fitur berjalan tetapi ada perilaku salah | Pesan error tidak muncul |
| Low | Kosmetik atau hal kecil | Teks salah ketik |

### Menulis Current Result yang baik

Kalimat pertama Current Result dipakai sebagai **judul tiket bug**, jadi tulis faktanya dengan singkat dan spesifik.

- Kurang baik: `error`, `tidak sesuai`
- Baik: `Transfer -50 diterima, saldo rekening asal bertambah 50`

## 4. Alur satu sesi pengujian

1. `git pull`
2. Buka ParaBank, buat akun baru (PB-AUT-01), lalu catat di sheet **Test Data**.
3. Jalankan test case. Untuk setiap test case:
   1. Baca Pre-Condition, Test Step, dan Test Data.
   2. Lakukan langkahnya di ParaBank.
   3. Bandingkan hasilnya dengan Expected Result.
   4. Isi Current Result dan Status (serta Severity dan Capture jika Failed).
4. Simpan Excel, lalu **tutup Excel**. File yang masih terbuka dikunci oleh Windows.
5. Commit dan push:
   ```powershell
   git add manual evidence
   git commit -m "Eksekusi modul Transfer"
   git push
   ```
6. Tunggu 1–2 menit, lalu cek:
   - tab **Actions** di GitHub: workflow *Manual QA sync* berstatus hijau;
   - **Jira**: tiket bug baru untuk setiap baris Failed;
   - **dashboard**: angka terbaru.
7. `git pull` untuk mengambil Excel yang sudah berisi Bug ID.

> **Penting:** selalu `git pull` sebelum mulai mengedit Excel. Bot ikut menyimpan perubahan ke file yang sama. Excel adalah file biner, jadi konflik tidak bisa digabung otomatis.

## 5. Contoh lengkap satu bug

**Test case:** PB-TRF-06, transfer dengan nominal negatif.

1. Di ParaBank: Transfer Funds, Amount `-50`, klik Transfer. Misalkan transfernya **diterima**.
2. Ambil screenshot dan simpan sebagai `evidence/PB-TRF-06.png`.
3. Isi Excel di sheet Transfer, baris PB-TRF-06:

   | Current Result | Status | Severity | Capture |
   |---|---|---|---|
   | Transfer -50 diterima, saldo rekening asal bertambah 50 | Failed | Critical | evidence/PB-TRF-06.png |

4. Push. GitHub Actions lalu:
   - membuat **PB-14** di Jira berjudul `[PB-TRF-06] Transfer dengan nominal negatif: Transfer -50 diterima, saldo rekening asal bertambah 50`, berisi langkah, test data, expected, actual, dan screenshot, dengan priority Highest;
   - menghubungkan PB-14 ke story **REQ-07** (PB-8);
   - menulis `PB-14` dan `To Do` ke Excel;
   - menandai PB-TRF-06 sebagai "Menunggu perbaikan" di dashboard.
5. Ketika PB-14 dipindah ke **Done**, pada sinkronisasi berikutnya (push berikutnya atau pukul 06.00 WIB) dashboard menandai PB-TRF-06 sebagai **Perlu retest**.
6. Kamu menguji ulang. Jika sudah benar, ubah Status menjadi `Passed`, lalu push.

### Kasus khusus: dua test case, satu bug

Misalnya PB-TRF-06 dan PB-BIL-06 sama-sama gagal karena nominal negatif tidak divalidasi. Push PB-TRF-06 dulu sampai dapat `PB-14`, lalu isi **Bug ID** `PB-14` secara manual di baris PB-BIL-06 sebelum push. Dengan begitu tidak ada tiket ganda.

## 6. Arti tindakan di dashboard

| Tindakan | Artinya | Yang kamu lakukan |
|---|---|---|
| Buat tiket bug | Failed, tapi tiket belum dibuat | Biasanya hilang setelah push berikutnya |
| Menunggu perbaikan | Failed, tiket masih terbuka | Tunggu |
| Perlu retest | Failed, tiket sudah Done | Uji ulang, lalu ubah Status |
| Lulus, tiket masih terbuka | Passed, tapi tiket belum Done | Tutup tiket di Jira |

## 7. Masalah yang sering muncul

| Gejala | Penyebab | Solusi |
|---|---|---|
| `git push` ditolak (*rejected*) | Bot sudah mengubah Excel | `git pull`, lalu push lagi. Jika konflik di Excel, simpan dulu salinan file kamu, jalankan `git checkout --theirs manual/ParaBank-Manual-QA.xlsx`, lalu isi ulang perubahanmu |
| Workflow merah di langkah Jira | Token kedaluwarsa atau salah | Buat token baru dan perbarui secret `JIRA_API_TOKEN` |
| Status tertulis "Lulus" muncul sebagai peringatan | Status salah ketik | Pakai pilihan dari dropdown |
| Screenshot tidak ada di tiket | Path Capture salah, atau file belum di-push | Pastikan file ada di `evidence/` dan ikut di-`git add` |
| Data ParaBank hilang di tengah sesi | Server publik di-reset | Buat akun baru. Ini bukan bug, catat di Notes |

## 8. Latihan

Kerjakan berurutan. Setiap latihan punya tujuan dan tanda berhasil.

### Latihan 1: test yang lulus

**Tujuan:** mengenal alur isi Excel → push → dashboard, tanpa bug.

1. `git pull`, lalu buka Excel di sheet **Autentikasi**.
2. Jalankan **PB-AUT-01** (registrasi) dengan username `rivaldi_` + tanggal hari ini, misalnya `rivaldi_1117`.
3. Catat akun di sheet **Test Data**.
4. Jalankan juga **PB-AUT-06** (login) dan **PB-AUT-07** (password salah).
5. Isi Current Result dan Status ketiganya, isi Device dan Tanggal tes, lalu simpan dan tutup Excel.
6. Commit dan push.

**Tanda berhasil:** workflow hijau; dashboard menunjukkan *Dieksekusi 3*; tidak ada tiket baru di Jira (karena tidak ada yang Failed).

### Latihan 2: menemukan dan melaporkan bug

**Tujuan:** melihat tiket bug terbentuk otomatis.

1. `git pull`, lalu buka sheet **Transfer**.
2. Login dengan akun latihan 1, lalu jalankan **PB-TRF-01**, **PB-TRF-02**, dan **PB-TRF-08** (transfer melebihi saldo).
3. Pada PB-TRF-08, perhatikan apakah transfer ditolak dan apakah saldo menjadi negatif di Accounts Overview.
4. Jika hasilnya tidak sesuai Expected Result:
   - simpan screenshot Accounts Overview sebagai `evidence/PB-TRF-08.png`;
   - isi Current Result, `Failed`, Severity, dan Capture.
5. Simpan, tutup Excel, commit (`git add manual evidence`), lalu push.

**Tanda berhasil:** di Jira muncul tiket Bug baru dengan screenshot dan link ke story REQ-06; setelah `git pull`, kolom Bug ID di Excel terisi; dashboard menampilkan bug tersebut.

Jika PB-TRF-08 ternyata lulus, coba **PB-TRF-06** (nominal negatif) atau **PB-TRF-10** (tiga desimal) sampai menemukan yang gagal.

### Latihan 3: retest

**Tujuan:** menutup putaran bug.

1. Di Jira, pindahkan tiket dari latihan 2 ke **Done**. Ini mensimulasikan developer yang sudah memperbaiki bug.
2. Buka GitHub → Actions → **Manual QA sync** → *Run workflow* (atau tunggu jadwal pukul 06.00 WIB).
3. Buka dashboard: test case tersebut muncul di **Perlu retest**, dan Bug Status di Excel menjadi `Done` setelah `git pull`.
4. Uji ulang di ParaBank. Karena ParaBank tidak benar-benar diperbaiki, hasilnya tetap gagal. Biarkan Status `Failed`, lalu buka kembali tiketnya di Jira (*Reopen* / pindahkan ke To Do) dan tambahkan komentar "Retest gagal" beserta tanggalnya.

**Tanda berhasil:** kamu sudah melewati seluruh putaran: eksekusi → bug → perbaikan → retest → bug dibuka kembali. Ini persis yang terjadi di tim QA sungguhan.

## 9. Checklist sebelum push

- [ ] Sudah `git pull` sebelum mulai mengedit
- [ ] Device dan Tanggal tes terisi
- [ ] Setiap test yang dijalankan punya Current Result dan Status
- [ ] Setiap Failed punya Severity, dan Capture jika ada screenshot
- [ ] Screenshot ada di folder `evidence/`
- [ ] Excel sudah disimpan dan ditutup
- [ ] `git add manual evidence` → `git commit` → `git push`
