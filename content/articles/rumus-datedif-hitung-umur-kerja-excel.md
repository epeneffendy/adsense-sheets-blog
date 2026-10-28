---
title: "Menghitung Masa Kerja dan Selisih Tanggal dengan Rumus DATEDIF Excel"
description: "Panduan menghitung tahun, bulan, dan hari masa kerja karyawan secara akurat menggunakan fungsi tersembunyi DATEDIF."
category: "Excel"
date: "2026-10-28"
readTime: "8 menit baca"
---

# Menghitung Masa Kerja dan Selisih Tanggal dengan Rumus DATEDIF Excel

Kalkulasi masa kerja karyawan butuh akurasi mutlak. Departemen Human Resources (HR) dan Payroll hitung masa kerja penentu hak finansial: pesangon pensiun, hak cuti tahunan, kenaikan upah berkala, bonus loyalitas.

Kalkulasi manual selisih hari bagi 365 timbulkan deviasi sistemik. Tahun kabisat miliki 366 hari. Variasi panjang bulan kalender berkisar 28 hingga 31 hari. Pembagian statis hasilkan pembulatan bias. Dampak: sengketa hubungan industrial akibat selisih nominal kompensasi regulasi ketenagakerjaan.

Excel sediakan fungsi bawaan: `DATEDIF`. Fungsi tersembunyi sejak versi Excel kuno warisan Lotus 1-2-3. Tidak muncul pada menu Formula Wizard atau autocomplete IntelliSense. Fungsi evaluasi selisih temporal presisi berbasis kalender gregorian riil.

---

## Mekanisme Internal Serial Number Tanggal Excel

Excel simpan seluruh data tanggal bentuk bilangan bulat serial (Serial Number). 

- Titik awal Epoch: 1 Januari 1900 catat serial `1`.
- Serial bertambah `1` tiap 24 jam.
- 1 Januari 2024 catat serial `45292`.
- Waktu (jam, menit, detik) catat bilangan desimal pecahan.

Excel pertahankan kompatibilitas mundur sistem Lotus 1-2-3, termasuk replikasi bug tahun kabisat fiktif 1900 (Excel anggap 1900 tahun kabisat). Fungsi `DATEDIF` proses serial number tanggal dasar evaluasi rentang waktu tanpa terpengaruh anomali serial mentah.

---

## Anatomi Sintaks dan Argumen Teknis DATEDIF

Fungsi `DATEDIF` terima tiga argumen input:

```excel
=DATEDIF(start_date, end_date, unit)
```

1. `start_date`: Tanggal awal periode kerja. Wajib bernilai serial valid atau hasil fungsi tanggal. Nilai `start_date` wajib lebih kecil atau sama dengan `end_date`.
2. `end_date`: Tanggal akhir periode kerja. Wajib bernilai serial valid.
3. `unit`: Kode string teks tentukan format interval output. Wajib diapit tanda petik ganda.

```
+-----------------------------------------------------------------------------------------------+
| KODE UNIT | DESKRIPSI OUTPUT                                 | ALGORITMA EVALUASI             |
+-----------------------------------------------------------------------------------------------+
| "Y"       | Jumlah tahun penuh berlalu.                      | Floor(selisih tahun kalender)  |
| "M"       | Jumlah total bulan penuh berlalu.                | Total interval bulan diskrit   |
| "D"       | Jumlah total hari kalender berlalu.              | end_date - start_date           |
| "YM"      | Sisa bulan setelah kurangi total tahun penuh.    | Modulo bulan dalam tahun aktif |
| "YD"      | Sisa hari setelah kurangi total tahun penuh.     | Hari abaikan selisih tahun     |
| "MD"      | Sisa hari setelah kurangi tahun dan bulan penuh. | Sisa selisih hari tanggal      |
+-----------------------------------------------------------------------------------------------+
```

---

## Komparasi Metode Kalkulasi Selisih Tanggal

Evaluasi perbandingan kalkulasi masa kerja:

```
+----------------------------------------------------------------------------------------------------+
| METODE             | SINTAKS                       | AKURASI BULAN/HARI | RISIKO TEKNIS            |
+----------------------------------------------------------------------------------------------------+
| Pembagian Statis   | =(B2 - A2) / 365              | Nol (Hanya Pecahan)| Gagal tahun kabisat      |
| Fungsi YEARFRAC    | =YEARFRAC(A2, B2, 1)          | Menengah (Pecahan) | Konversi ke bulan ambigu |
| Selisih YEAR/MONTH | =(YEAR(B2)-YEAR(A2))*12+...   | Rendah             | Abaikan hari cut-off     |
| Fungsi DATEDIF     | =DATEDIF(A2, B2, "Y")         | Mutlak Diskrit     | Bug unit "MD" (versi lama)|
+----------------------------------------------------------------------------------------------------+
```

Keunggulan `DATEDIF`: konversi otomatis rentang tanggal menuju output integer diskrit. Tidak butuh fungsi pembulatan `INT` atau `ROUNDDOWN`.

---

## Studi Kasus 1: Kalkulasi Format Baku "X Tahun, Y Bulan, Z Hari"

Divisi HR evaluasi masa kerja karyawan per cut-off tanggal 31 Oktober 2026. 

Struktur data tabel spreadsheet:
- Kolom A: ID Karyawan
- Kolom B: Nama Karyawan
- Kolom C: Tanggal Bergabung (Join Date)
- Kolom D: Tanggal Evaluasi (Cut-off Date)
- Kolom E: Masa Kerja Tekstual Lengkap

```
+--------------------------------------------------------------------------------------------------------------------+
|   |      A      |         B         |      C      |      D      |                           E                      |
+--------------------------------------------------------------------------------------------------------------------+
| 1 | ID KARYAWAN | NAMA KARYAWAN     | JOIN DATE   | CUT-OFF     | MASA KERJA LENGKAP                                   |
| 2 | EMP-0101    | Anton Wijaya      | 2018-03-15  | 2026-10-31  | 8 Tahun 7 Bulan 16 Hari                              |
| 3 | EMP-0102    | Siti Rahma        | 2020-02-29  | 2026-10-31  | 6 Tahun 8 Bulan 2 Hari                               |
| 4 | EMP-0103    | Budi Hartono      | 2023-11-01  | 2026-10-31  | 2 Tahun 11 Bulan 30 Hari                             |
| 5 | EMP-0104    | Ratna Dewi        | 2026-05-18  | 2026-10-31  | 0 Tahun 5 Bulan 13 Hari                              |
| 6 | EMP-0105    | Farhan Nugraha    | 2015-08-31  | 2026-10-31  | 11 Tahun 2 Bulan 0 Hari                              |
+--------------------------------------------------------------------------------------------------------------------+
```

### Formula Gabungan Sel E2:

```excel
=DATEDIF(C2, D2, "Y") & " Tahun " & DATEDIF(C2, D2, "YM") & " Bulan " & DATEDIF(C2, D2, "MD") & " Hari"
```

### Bedah Eksekusi Komputasi:

1. Sub-formula 1: `DATEDIF(C2, D2, "Y")`
   - Input: `2018-03-15` dan `2026-10-31`.
   - Output: `8`. Karyawan lewati 8 kali siklus tahunan penuh per 15 Maret 2026.
2. Sub-formula 2: `DATEDIF(C2, D2, "YM")`
   - Hitung selisih bulan abaikan angka tahun.
   - Periode sisa: `2026-03-15` sampai `2026-10-15`.
   - Output: `7`.
3. Sub-formula 3: `DATEDIF(C2, D2, "MD")`
   - Hitung selisih sisa hari abaikan komponen tahun dan bulan penuh.
   - Periode sisa: `2026-10-15` sampai `2026-10-31`.
   - Output: `16`.
4. Operator Penggabung `&`: Rangkai nilai integer dan label teks jadi satu string utuh.

---

## Analisis Cacat Logika Unit "MD" dan Formula Alternatif

Unit kode `"MD"` miliki cacat kalkulasi internal (documented bug) pada build Microsoft Excel tertentu. Bug muncul saat tanggal awal jatuh pada akhir bulan panjang (31 hari) dan tanggal akhir jatuh pada bulan pendek (Februari/April/Juni/September/November). Hasil komputasi timbulkan angka negatif atau lompatan nilai tidak rasional.

Contoh kasus galat "MD":
- Tanggal Mulai: 31 Januari 2026
- Tanggal Akhir: 1 Maret 2026
- `DATEDIF(..., "MD")` menghasilkan nilai galat `-2` atau `29` pada versi Excel tertentu karena algoritma kurangi 31 dari 28 hari.

### Solusi Perbaikan Mandiri (Bug-Free Formula)

Ganti unit `"MD"` pakai kombinasi kalkulasi dinamis fungsi `EDATE`:

```excel
=D2 - EDATE(C2, DATEDIF(C2, D2, "M"))
```

Logika formula pengganti:
- `DATEDIF(C2, D2, "M")` hitung akumulasi total bulan penuh berlalu.
- `EDATE(C2, ...)` proyeksikan tanggal masuk maju sebanyak total bulan penuh.
- `D2 - EDATE(...)` kurangi serial cut-off dengan serial tanggal proyeksi. Hasilkan selisih hari absolut bebas galat interpolasi.

Formula produksi tahan cacat logika untuk sel E2:

```excel
=DATEDIF(C2, D2, "Y") & " Tahun " & DATEDIF(C2, D2, "YM") & " Bulan " & (D2 - EDATE(C2, DATEDIF(C2, D2, "M"))) & " Hari"
```

---

## Studi Kasus 2: Kalkulasi Kompleks Hak Kompensasi Pesangon

Regulasi ketenagakerjaan tentukan ambang batas uang pesangon berdasar masa kerja tahunan bulat. Karyawan aktif gunakan acuan tanggal sistem lokal (`TODAY()`). Karyawan berhenti kerja gunakan tanggal keluar efektif.

Matriks Hak Pesangon Pokok:
- Masa kerja < 1 tahun: 1 bulan upah.
- Masa kerja 1 tahun s.d. < 2 tahun: 2 bulan upah.
- Masa kerja 2 tahun s.d. < 3 tahun: 3 bulan upah.
- Masa kerja bertambah: tambah 1 bulan upah tiap tahun masa kerja.
- Batas atas maksimum: 9 bulan upah (masa kerja >= 8 tahun).

Struktur tabel data sel A1:H6:
- Kolom A: ID Karyawan
- Kolom B: Tanggal Masuk
- Kolom C: Tanggal Resign (Kosong = Masih Aktif)
- Kolom D: Gaji Pokok
- Kolom E: Tanggal Evaluasi Efektif
- Kolom F: Masa Kerja Tahun Penuh (Komputasi)
- Kolom G: Faktor Pengali Upah Pesangon
- Kolom H: Estimasi Kewajiban Nominal Pesangon

```
+-----------------------------------------------------------------------------------------------------------------------+
|   |    A     |     B      |     C      |     D      |     E      |    F    |    G    |              H               |
+-----------------------------------------------------------------------------------------------------------------------+
| 1 | ID EMP   | JOIN       | RESIGN     | BASE SAL   | EFF DATE   | THN WJK | MULTIP  | TOTAL PESANGON               |
| 2 | EMP-201  | 2014-06-01 |            | 12000000   | 2026-10-28 | 12      | 9       | 108000000                    |
| 3 | EMP-202  | 2021-02-15 | 2024-08-20 | 8500000    | 2024-08-20 | 3       | 4       | 34000000                     |
| 4 | EMP-203  | 2025-11-10 |            | 6000000    | 2026-10-28 | 0       | 1       | 6000000                      |
| 5 | EMP-204  | 2019-04-01 | 2026-01-15 | 15000000   | 2026-01-15 | 6       | 7       | 105000000                    |
+-----------------------------------------------------------------------------------------------------------------------+
```

### 1. Formula Kolom E (Tanggal Evaluasi Efektif):
Gunakan fungsi `IF` periksa status keterisian sel tanggal keluar:

```excel
=IF(ISBLANK(C2), TODAY(), C2)
```

### 2. Formula Kolom F (Masa Kerja Tahun Penuh):
Ekstraksi tahun bulat lewati validasi tanggal:

```excel
=DATEDIF(B2, E2, "Y")
```

### 3. Formula Kolom G (Faktor Multiplier Hak Upah Pesangon):
Gunakan fungsi `MIN` batasi pengali tertinggi plafon 9 bulan upah sesuai batas undang-undang:

```excel
=MIN(9, F2 + 1)
```

Logika matematis: Masa kerja 0 tahun beri 1 bulan upah. Masa kerja 1 tahun beri 2 bulan upah. Pola linier `Tahun + 1`. Fungsi `MIN(9, ...)` potong nilai jika hasil penjumlahan lampaui 9.

### 4. Formula Kolom H (Total Nilai Finansial Pesangon):
Perkalian upah dasar dengan pengali:

```excel
=D2 * G2
```

Formula tunggal gabungan tanpa kolom bantu (Single-Cell Dynamic Formula):

```excel
=D2 * MIN(9, DATEDIF(B2, IF(ISBLANK(C2), TODAY(), C2), "Y") + 1)
```

---

## Penanganan Galat dan Jebakan Teknis (Pitfalls)

### 1. Galat `#NUM!`
Penyebab teknis: Nilai `start_date` lebih besar secara serial dibanding nilai `end_date`. Excel tolak selisih waktu bernilai negatif.

Pencegahan defensif:
```excel
=IF(B2 > C2, "Input Invalid: Terbalik", DATEDIF(B2, C2, "Y"))
```

### 2. Galat `#VALUE!`
Penyebab teknis: Salah satu argumen tanggal tersimpan format teks non-numerik (string). Sering terjadi akibat proses impor file CSV beda regional.

Verifikasi status sel:
```excel
=ISNUMBER(B2)
```
Jika formula hasilkan `FALSE`, bersihkan teks pakai `DATEVALUE` atau potong karakter spasi:
```excel
=DATEDIF(DATEVALUE(TRIM(B2)), DATEVALUE(TRIM(C2)), "Y")
```

### 3. Delimiter Regional: Koma `,` vs Titik Koma `;`
Excel patuhi Regional Settings sistem operasi host:
- Format English / US: Pemisah argumen gunakan koma `,`.
- Format Indonesia / Latin Europe: Pemisah argumen gunakan titik koma `;`.

Jika salin formula format US ke Excel format lokal Indonesia:
Ganti: `=DATEDIF(A2, B2, "Y")`
Menjadi: `=DATEDIF(A2; B2; "Y")`

### 4. Referensi Sel Absolut vs Relatif
Tanggal cut-off audit evaluasi statis wajib kunci referensi sel absolut pakai simbol dollar `$`.

Salah (Relatif geser baris):
```excel
=DATEDIF(C2, D1, "Y")
```
Benar (Absolut terkunci):
```excel
=DATEDIF(C2, $D$1, "Y")
```

---

## Implementasi Modern: Excel 365 LAMBDA dan Google Sheets

Pengguna arsitektur cloud spreadsheet butuh standardisasi formula reusable tanpa makro VBA.

### 1. Formula Kustom Microsoft Excel 365 (LAMBDA)
Daftarkan fungsi bernama `MASA_KERJA` lewat Name Manager (Ctrl + F3):

```excel
=LAMBDA(join_date, cut_off, LET(end_d, IF(ISBLANK(cut_off), TODAY(), cut_off), thn, DATEDIF(join_date, end_d, "Y"), bln, DATEDIF(join_date, end_d, "YM"), hari, end_d - EDATE(join_date, DATEDIF(join_date, end_d, "M")), thn & " Thn, " & bln & " Bln, " & hari & " Hr"))
```

Pemanggilan pada lembar kerja:
```excel
=MASA_KERJA(C2, D2)
```

### 2. Google Sheets Apps Script (Alternatif Cloud Backend)
Google Sheets dukung fungsi `DATEDIF` asli pada antarmuka sel. Pemrosesan batch skala besar pada script otomasi gunakan JavaScript Google Apps Script:

```javascript
/**
 * Hitung selisih masa kerja karyawan format terstruktur.
 * @param {Date} startDate Tanggal mulai kerja.
 * @param {Date} endDate Tanggal akhir evaluasi.
 * @return {string} Masa kerja format teks.
 * @customfunction
 */
function HITUNG_MASA_KERJA_JS(startDate, endDate) {
  if (!startDate || !(startDate instanceof Date)) {
    throw new Error("Parameter startDate tidak valid.");
  }
  
  var start = new Date(startDate);
  var end = endDate ? new Date(endDate) : new Date();
  
  if (start > end) {
    throw new Error("startDate melebihi endDate.");
  }
  
  var years = end.getFullYear() - start.getFullYear();
  var months = end.getMonth() - start.getMonth();
  var days = end.getDate() - start.getDate();
  
  if (days < 0) {
    months -= 1;
    var prevMonthLastDay = new Date(end.getFullYear(), end.getMonth(), 0).getDate();
    days += prevMonthLastDay;
  }