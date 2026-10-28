---
title: "Menguasai SUMIFS dan COUNTIFS Excel: Agregasi Data dengan Kriteria Majemuk"
description: "Panduan kalkulasi total dan frekuensi data berdasarkan rentang tanggal, status transaksi, dan teks dinamis menggunakan SUMIFS dan COUNTIFS."
category: "Excel"
date: "2026-10-28"
readTime: "8 menit baca"
---

## 1. Pendahuluan

Pertumbuhan volume data transaksional pada lingkungan enterprise menuntut metode pemrosesan data yang deterministik, cepat, dan otomatis. Masalah umum pada pengolahan data spreadsheet skala besar adalah ketergantungan pada filter manual atau pembuatan kolom pembantu (*helper columns*) secara berulang. Pendekatan konvensional tersebut memperbesar ukuran file, memicu fragmentasi data, dan meningkatkan risiko kesalahan manusia (*human error*).

Pivot Table kerap digunakan untuk agregasi. Namun, Pivot Table bersifat semi-statis: agregasi memerlukan tindakan penyegaran (*refresh*) manual atau pemicu VBA/Macro setiap kali basis data berubah. Formula agregasi bersyarat multi-kriteria—spesifiknya `SUMIFS` dan `COUNTIFS`—menjadi solusi defakto untuk kalkulasi dinamis waktu nyata (*real-time*). 

`SUMIFS` dan `COUNTIFS` dirancang untuk memproses evaluasi logika `AND` lintas rentang baris secara simultan tanpa memerlukan matriks array kompleks seperti formula legasi `SUMPRODUCT`. Memahami arsitektur argumen, interaksi tipe data, penanganan rentang tanggal, karakter *wildcard*, serta mekanika referensi sel merupakan syarat mutlak dalam membangun model data finansial dan operasional yang bebas galat.

---

## 2. Anatomi Sintaks Fungsi dan Parameter

Kedua fungsi agregasi ini beroperasi berdasarkan evaluasi multivariat. Setiap elemen kriteria dievaluasi berpasangan terhadap rentang sel targetnya.

### 2.1 Sintaks SUMIFS

Fungsi `SUMIFS` menjumlahkan nilai dalam rentang yang memenuhi serangkaian kriteria majemuk.

```excel
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

Penjelasan Parameter:
*   `sum_range` (Wajib): Rentang sel numerik aktual yang akan dijumlahkan. Parameter ini berada di posisi paling depan.
*   `criteria_range1` (Wajib): Rentang sel pertama yang dievaluasi menggunakan kriteria uji terkait.
*   `criteria1` (Wajib): Kondisi atau pengujian logika yang diterapkan pada `criteria_range1`. Dapat berupa angka, ekspresi logika, referensi sel, teks, atau penggunaan karakter *wildcard*.
*   `criteria_range2, criteria2, ...` (Opsional): Pasangan rentang dan kriteria tambahan. Excel mendukung hingga 127 pasangan rentang/kriteria.

### 2.2 Sintaks COUNTIFS

Fungsi `COUNTIFS` menghitung frekuensi kemunculan baris data yang memenuhi satu atau beberapa kondisi secara bersamaan.

```excel
=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

Penjelasan Parameter:
*   `criteria_range1` (Wajib): Rentang sel pertama yang diuji.
*   `criteria1` (Wajib): Kondisi batas frekuensi dalam bentuk nilai numerik, teks, perbandingan komparatif, atau ekspresi dinamis.
*   `criteria_range2, criteria2, ...` (Opsional): Rentang dan kriteria lanjutan hingga batas maksimal 127 pasangan.

### 2.3 Perbedaan Struktural Terhadap Formula Tunggal

Pada fungsi versi kriteria tunggal (`SUMIF`), urutan argumen meletakkan rentang penjumlahan di akhir:

```excel
=SUMIF(criteria_range, criteria, [sum_range])
```

Pada `SUMIFS`, `sum_range` digeser ke posisi parameter pertama. Pergeseran ini wajib diperhatikan saat melakukan migrasi formula lama ke formula multi-kriteria untuk menghindari kegagalan kalkulasi logika.

---

## 3. Komparasi Metode Agregasi Data Spreadsheet

Tabel berikut membandingkan karakteristik performa dan operasional antara `SUMIFS`/`COUNTIFS` dengan pendekatan alternatif di Microsoft Excel:

| Kriteria Metrik | SUMIFS / COUNTIFS | SUMIF / COUNTIF Legasi | Pivot Table | Kolom Bantu + SUBTOTAL / Filter | SUMPRODUCT |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Kriteria Majemuk** | Didukung bawaan (Native AND) | Tidak didukung (Butuh manipulasi `&`) | Didukung | Didukung secara visual | Didukung penuh (AND/OR logic) |
| **Pembaruan Kalkulasi** | Otomatis (*Real-time Recalc*) | Otomatis (*Real-time Recalc*) | Manual (*Needs Data Refresh*) | Manual (Klik filter ulang) | Otomatis (*Real-time Recalc*) |
| **Beban Komputasi** | Rendah (Dioptimalkan oleh kalkulator internal) | Rendah | Sangat Rendah (Cached memory) | Sedang (Tergantung thread visual) | Tinggi (Operasi perkalian matriks) |
| **Kompleksitas Sintaks** | Rendah - Menengah | Rendah | Tanpa formula | Rendah | Tinggi |
| **Integritas Struktur** | Tinggi (Tidak memodifikasi tabel data) | Sedang (Membutuhkan modifikasi kolom) | Terpisah dari tabel sumber | Rentan (Mengubah susunan baris tersembunyi) | Tinggi |

---

## 4. Studi Kasus 1: Pelaporan Penjualan Cabang Reguler

Skenario: Analisis performa tim penjualan berdasarkan kategori produk dan status penyelesaian transaksi.

### 4.1 Struktur Data Mentah (Rentang A1:E11)

| Baris | A (Tanggal) | B (Kategori) | C (Status) | D (Sales Rep) | E (Nilai Transaksi) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **Tanggal** | **Kategori** | **Status** | **Sales Rep** | **Nilai Transaksi** |
| **2** | 2026-01-05 | Elektronik | Selesai | Budi | 15000000 |
| **3** | 2026-01-07 | Furniture | Selesai | Siti | 4500000 |
| **4** | 2026-01-12 | Elektronik | Pending | Budi | 8000000 |
| **5** | 2026-01-15 | Pakaian | Selesai | Joko | 1200000 |
| **6** | 2026-01-18 | Elektronik | Selesai | Siti | 22000000 |
| **7** | 2026-01-20 | Furniture | Batal | Joko | 3000000 |
| **8** | 2026-01-22 | Elektronik | Selesai | Budi | 12500000 |
| **9** | 2026-01-25 | Pakaian | Pending | Siti | 950000 |
| **10** | 2026-01-28 | Elektronik | Selesai | Joko | 17000000 |
| **11** | 2026-01-30 | Furniture | Selesai | Budi | 6000000 |

### 4.2 Masalah Bisnis
1. Hitung total nilai nominal transaksi yang memenuhi tiga kondisi: Kategori adalah "Elektronik", Status adalah "Selesai", dan ditangani oleh Sales Rep bernama "Budi".
2. Hitung jumlah total pesanan yang berhasil diselesaikan oleh Sales Rep "Budi" tanpa memandang kategori barang.

### 4.3 Implementasi Formula

Formula kalkulasi total nominal:

```excel
=SUMIFS(E2:E11, B2:B11, "Elektronik", C2:C11, "Selesai", D2:D11, "Budi")
```

Formula kalkulasi frekuensi transaksi selesai:

```excel
=COUNTIFS(C2:C11, "Selesai", D2:D11, "Budi")
```

### 4.4 Bedah Logika Eksekusi Internal
Pada formula `SUMIFS`, mesin Excel mengevaluasi setiap baris data menggunakan operasi Boolean:
*   Kriteria 1 (`B2:B11 = "Elektronik"`): `{TRUE; FALSE; TRUE; FALSE; TRUE; FALSE; TRUE; FALSE; TRUE; FALSE}`
*   Kriteria 2 (`C2:C11 = "Selesai"`): `{TRUE; TRUE; FALSE; TRUE; TRUE; FALSE; TRUE; FALSE; TRUE; TRUE}`
*   Kriteria 3 (`D2:D11 = "Budi"`): `{TRUE; FALSE; TRUE; FALSE; FALSE; FALSE; TRUE; FALSE; FALSE; TRUE}`

Operasi `AND` antar-array menghasilkan matriks filter:
*   Baris 2: `TRUE * TRUE * TRUE = 1` (Nilai: 15.000.000)
*   Baris 4: `TRUE * FALSE * TRUE = 0` (Diabaikan)
*   Baris 6: `TRUE * TRUE * FALSE = 0` (Diabaikan)
*   Baris 8: `TRUE * TRUE * TRUE = 1` (Nilai: 12.500.000)
*   Baris lainnya menghasilkan nilai `0`.

Hasil kalkulasi: `15.000.000 + 12.500.000 = 27.500.000`.

Untuk `COUNTIFS`, proses evaluasi logika identik. Namun, alih-alih menjumlahkan nilai pada kolom `E`, mesin menjumlahkan total kemunculan kondisi `1`, menghasilkan output nilai `2`.

---

## 5. Studi Kasus 2: Analisis Lanjutan Berdasarkan Rentang Tanggal, Wildcard, dan Logika Pengecualian

Skenario korporat membutuhkan filtering berbasis rentang waktu dinamis (*date bounded range*), identifikasi string parsial, serta pengecualian status pembatalan.

### 5.1 Struktur Data Baru (Rentang A1:E11)

| Baris | A (Tanggal) | B (Kode SKU) | C (Kanal) | D (Status) | E (Omzet Bersih) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **Tanggal** | **Kode SKU** | **Kanal** | **Status** | **Omzet Bersih** |
| **2** | 2026-02-01 | ELK-LPT-01 | Online | Sukses | 25000000 |
| **3** | 2026-02-05 | FUR-MEJ-09 | Offline | Sukses | 1800000 |
| **4** | 2026-02-12 | ELK-MNT-04 | Online | Batal | 4200000 |
| **5** | 2026-02-15 | ELK-LPT-02 | Corporate | Sukses | 52000000 |
| **6** | 2026-02-20 | PKA-KJD-01 | Online | Sukses | 750000 |
| **7** | 2026-03-01 | ELK-ACC-11 | Online | Sukses | 1200000 |
| **8** | 2026-03-05 | FUR-KUR-02 | Offline | Sukses | 3100000 |
| **9** | 2026-03-10 | ELK-LPT-03 | Corporate | Pending | 34000000 |
| **10** | 2026-03-15 | PKA-KJM-02 | Offline | Sukses | 620000 |
| **11** | 2026-03-22 | ELK-ACC-12 | Online | Sukses | 850000 |

### 5.2 Kriteria Parameter Uji
Kalkulasi omzet agregat dengan batasan:
1. Rentang tanggal transaksi: Antara `2026-02-01` sampai dengan `2026-02-28` (Bulan Februari 2026).
2. Tipe barang: Spesifik perangkat elektronik (*Prefix* SKU diawali teks `ELK`).
3. Status: Mengabaikan transaksi yang berstatus `Batal`.
4. Nilai referensi batas tanggal diletakkan di sel `G1` (`2026-02-01`) dan `G2` (`2026-02-28`).

### 5.3 Implementasi Formula Dinamis

Formula dengan referensi sel dinamis dan operator komparasi:

```excel
=SUMIFS(E2:E11, A2:A11, ">="&G1, A2:A11, "<="&G2, B2:B11, "ELK*", D2:D11, "<>Batal")
```

Formula menghitung jumlah transaksi unik yang memenuhi kondisi di atas:

```excel
=COUNTIFS(A2:A11, ">="&G1, A2:A11, "<="&G2, B2:B11, "ELK*", D2:D11, "<>Batal")
```

### 5.4 Penjelasan Mekanisme Operator

1.  **Penggabungan String dan Operator (`">="&G1`)**:
    Operator perbandingan (`>`, `<`, `>=`, `<=`, `<>`) wajib diapit tanda kutip dua dan digabungkan dengan referensi sel menggunakan ampersand (`&`). Penulisan langsung seperti `">=G1"` menyebabkan Excel mencari teks harfiah "G1" alih-alih mengekstrak nilai tanggal serial di dalam sel tersebut.
2.  **Karakter Wildcard Tanda Bintang (`"ELK*"`)**:
    Simbol asteris (`*`) mewakili jumlah karakter alfanumerik bebas tanpa batas panjang. Kriteria `"ELK*"` memvalidasi kode `ELK-LPT-01`, `ELK-MNT-04`, dan `ELK-LPT-02`. Jika pengujian memerlukan pembatasan tepat satu karakter, gunakan karakter tanda tanya (`?`).
3.  **Operator Pertidaksamaan (`"<>Batal"`)**:
    Kondisi `<>` mengecualikan nilai teks tertentu dari evaluasi penjumlahan.

### 5.5 Eksekusi Logika Skenario OR Majemuk Menggunakan Array Constant

Secara bawaan, `SUMIFS` hanya menjalankan operasi berbasis `AND`. Apabila kalkulasi membutuhkan agregasi dengan logika `OR` (misalnya: Kategori "Online" **ATAU** "Corporate"), gunakan konstanta array berpasangan dengan fungsi pembungkus `SUM`:

```excel
=SUM(SUMIFS(E2:E11, C2:C11, {"Online","Corporate"}, D2:D11, "Sukses"))
```

Mekanika eksekusi:
1. `SUMIFS` dijalankan dua kali secara independen untuk menghasilkan array perantara: `{Total Omzet Online, Total Omzet Corporate}`.
2. Fungsi `SUM` mengeksekusi penjumlahan skalar terluar dari elemen array perantara tersebut.

---

## 6. Diagnosis Pitfall dan Analisis Kesalahan Umum

Implementasi agregasi bersyarat sering kali memicu galat fungsional atau kesalahan kalkulasi data senyap (*silent calculation errors*). Berikut rincian masalah dan mitigasinya:

### 6.1 Galat #VALUE! Akibat Dimensi Rentang Tidak Simetris
Penyebab utama munculnya galat `#VALUE!` pada `SUMIFS` dan `COUNTIFS` adalah ketidakcocokan jumlah baris atau kolom antar-parameter rentang.

```excel
SALAH   : =SUMIFS(E2:E11, A2:A11, ">="&G1, B2:B10, "ELK*")
BENAR   : =SUMIFS(E2:E11, A2:A11, ">="&G1, B2:B11, "ELK*")
```
Analisis: Rentang `B2:B10` memiliki 9 baris, sedangkan `E2:E11` memiliki 10 baris. Mesin kalkulasi Excel mewajibkan seluruh parameter `criteria_range` memiliki dimensi matriks yang identik dengan `sum_range`.

### 6.2 Kegagalan Sintaks Delimiter Regional
Penggunaan tanda pemisah (*list separator*) bergantung pada konfigurasi bahasa sistem operasi:
*   Format Regional Indonesia / Eropa: Menggunakan titik koma (`;`).
*   Format Regional Amerika / Internasional: Menggunakan koma (`,`).

```excel
Sistem ID   : =SUMIFS(E2:E11; B2:B11; "Elektronik")
Sistem EN   : =SUMIFS(E2:E11, B2:B11, "Elektronik")
```
Pemanggilan formula dengan delimiter yang salah mengakibatkan *error* parsing sintaks formula spreadsheet.

### 6.3 Inkonsistensi Referensi Sel Saat Replikasi Formula
Penyalinan formula antar-baris atau antar-kolom tanpa penguncian referensi absolut memicu pergeseran batas rentang agregasi.

```excel
SALAH (Relatif) : =SUMIFS(E2:E11, B2:B11, H2)
BENAR (Absolut) : =SUMIFS($E$2:$E$11, $B$2:$B$11, H2)
```
Tanda dolar (`$`) mengunci baris dan kolom agar tidak mengalami pergeseran indeks saat formula diseret (*drag*) ke sel lain.

### 6.4 Mismatch Tipe Data Numerik dan Teks
Angka yang tersimpan sebagai teks (*numbers formatted as text*) tidak dihitung oleh `sum_range` pada `SUMIFS`. Nilai tersebut diabaikan secara diam-diam tanpa memunculkan pesan galat, sehingga menghasilkan jumlah total di bawah nilai riil (*underreported metrics*).
Mitigasi: Konversi kolom numerik menggunakan fitur *Text to Columns*, operasi matematika paksa (`=A2*1`), atau penerapan fungsi validasi tipe data `=ISNUMBER(E2)`.

---

## 7. Checklist Praktik Terbaik dan Ringkasan Eksekutif

Terapkan prosedur kerja berikut sebelum mempublikasikan model data:

```
[ ] Dimensi Seluruh Rentang Identik:
    Pastikan sum_range dan seluruh criteria_range memiliki indeks awal dan akhir baris yang persis sama (misal: baris 2 hingga 1000).

[ ] Kunci Referensi Rentang Database:
    Gunakan referensi absolut ($A$2:$E$1000) atau terapkan Excel Tables (Struktur Structured Reference seperti Table1[Nilai]).

[ ] Validasi Tipe Data Tanggal:
    Pastikan sel target tanggal berupa serial date resmi Excel, bukan format string teks murni (verifikasi dengan fungsi =ISNUMBER(A2)).

[ ] Pola Operator Komparasi Dinamis:
    Pastikan operator perbandingan teks diapit tanda kutip dan disambung dengan ampersand (contoh: ">="&K1, bukan ">=K1").

[ ] Bersihkan Spasi Tersembunyi:
    Gunakan fungsi =TRIM() dan =CLEAN() pada data mentah untuk mencegah ketidakterbacaan karakter akibat trailing spaces.

[ ] Hindari Rentang Satu Kolom Utuh (Full-Column References):
    Batasi rentang seperti A2:A50000 daripada A:A pada dataset masif untuk menekan konsumsi alokasi memori recalculation thread.
```

### Ringkasan Eksekutif

*   `SUMIFS` dan `COUNTIFS` adalah solusi agregasi bersyarat multi-kriteria berbasis logika `AND` tanpa dependensi skrip atau pembaruan manual.
*   Parameter `sum_range` selalu ditempatkan pada urutan pertama di `SUMIFS`, berbeda dengan formula legasi `SUMIF`.
*   Ekspresi komparatif dinamis memerlukan pemisahan operator string dari referensi sel target (`"<="&Cell`).
*   Penerapan karakter *wildcard* (`*` dan `?`) menyediakan fleksibilitas pencarian string parsial berkecepatan tinggi.
*   Logika majemuk bercabang `OR` dapat diimplementasikan secara efisien melalui perpaduan fungsi `SUM` dengan konstanta array.