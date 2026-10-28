---
title: "Panduan Lengkap Rumus XLOOKUP Excel: Pengganti VLOOKUP dan HLOOKUP"
description: "Pelajari cara kerja fungsi XLOOKUP di Microsoft Excel, keunggulan dibanding VLOOKUP tradisional, dan cara pencarian horizontal serta backward lookup tanpa ribet."
category: "Excel"
date: "2025-02-17"
readTime: "12 min"
---

## 1. Evolusi Mesin Pencarian Excel: Masalah Formula Warisan

Excel andalkan `VLOOKUP` sejak 1985. Fungsi bawa kelemahan arsitektur puluhan tahun.

```
Masalah struktural VLOOKUP:
1. Lookup vector terikat rentang table_array. Kolom kunci wajib di indeks 1 (paling kiri).
2. Pencarian backward (kiri) butuh trik CHOOSE() atau kombinasi INDEX/MATCH.
3. Parameter col_index_num pakai integer statis. Tambah/hapus kolom bikin referensi geser, data korup.
4. Default match_mode bernilai TRUE (approximate). Lupa pasang FALSE hasilkan data halusinasi tanpa trigger error.
5. Operasi boros komputasi. Excel simpan referensi seluruh blok data ke memory tree, bukan kolom target saja.
```

`HLOOKUP` bawa masalah simetris. Kunci wajib baris teratas. Transposisi data manual boros waktu.

Formula `INDEX/MATCH` tuntaskan batasan arah lookup. Namun punya limitasi teknis:
- Butuh dua fungsi bertingkat.
- Beban kognitif tinggi bagi tim audit.
- Parsing formula lambat pada eksekusi jutaan baris data mentah.

Tahun 2019, Microsoft rombak Calculation Engine Excel. Arsitektur Dynamic Array rilis. Mesin hitung dukung kalkulasi vektor native. `XLOOKUP` hadir jadi solusi terpadu pengganti `VLOOKUP`, `HLOOKUP`, `LOOKUP`, dan `INDEX/MATCH`.

---

## 2. Anatomi Sintaksis XLOOKUP

Sintaks resmi `XLOOKUP`:

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

Fungsi terima 6 argumen. 3 wajib, 3 opsional.

```
                      +-------------------+
                      |   lookup_value    | -> Nilai dicari (skalar / array)
                      +---------+---------+
                                |
             +------------------+------------------+
             v                                     v
   +-------------------+                 +-------------------+
   |   lookup_array    |                 |   return_array    |
   | (Rentang referensi|                 |  (Rentang target  |
   |    pencarian)     |                 |  ekstraksi data)  |
   +---------+---------+                 +---------+---------+
             |                                     |
             +------------------+------------------+
                                |
             +------------------+------------------+
             | ARGUMEN OPSIONAL                    |
             v                                     v
   +-------------------+                 +-------------------+
   |   if_not_found    |                 |    match_mode     |
   | (Fallback error   |                 | (0=Exact, -1=<=,  |
   |  tanpa IFERROR)   |                 |  1=>=, 2=Wildcard)|
   +-------------------+                 +---------+---------+
                                                   |
                                                   v
                                         +-------------------+
                                         |    search_mode    |
                                         | (1=Depan, -1=Blkg,|
                                         |  2=Bin+, -2=Bin-) |
                                         +-------------------+
```

### 2.1. Parameter Wajib

#### `lookup_value`
Nilai identitas pencarian. Terima tipe data skalar (teks, angka, boolean) atau array dinamis referensi sel. Bila diisi rentang multi-sel (vektor), formula hasilkan output spill array.

#### `lookup_array`
Rentang data asal nilai `lookup_value`. Berupa larik satu dimensi (vektor 1D) vertikal (`A2:A100`) atau horizontal (`A2:Z2`). Larik tidak butuh dimensi sama dengan tabel sumber. Hanya butuh panjang larik sinkron dengan `return_array`.

#### `return_array`
Rentang data diambil. Vektor 1D atau matriks 2D (`B2:D100`). Jika dimensi kolom lebih dari satu, Excel lemparkan nilai seluruh kolom secara dinamis lewat fitur *dynamic spill*.

### 2.2. Parameter Opsional

#### `[if_not_found]`
Ganti peran `IFERROR()` atau `IFNA()`. Nilai otomatis keluar jika kunci tidak cocok. Evaluasi internal fungsi hemat alokasi heap memori kalkulasi. String kosong diisi `""`, nilai default teks diisi `"Data Tidak Ada"`.

#### `[match_mode]`
Atur logika perbandingan bit/karakter. Nilai default `0`.
- `0`: Pencarian persis (Exact Match). Kunci sama persis atau kembalikan error `#N/A`.
- `-1`: Pencarian persis. Jika nihil, ambil elemen berikutnya lebih kecil dari `lookup_value`. Cocok untuk skema pajak/diskon bertingkat.
- `1`: Pencarian persis. Jika nihil, ambil elemen berikutnya lebih besar dari `lookup_value`. Cocok untuk batas toleransi minimum produksi.
- `2`: Pencarian pola teks pakai karakter wildcard: tanda bintang `*` (karakter majemuk), tanda tanya `?` (karakter tunggal), dan tilde `~` (escape character).

#### `[search_mode]`
Atur arah traversal pointer algoritma pencarian. Nilai default `1`.
- `1`: Traversal linier indeks awal ke akhir (First-to-last).
- `-1`: Traversal linier indeks akhir ke awal (Last-to-first / Reverse search). Ambil histori transaksi paling mutakhir.
- `2`: Pencarian biner (Binary Search). Rentang `lookup_array` **wajib** terurut naik (Ascending / $A-Z$). Kompleksitas waktu turun dari $O(N)$ ke $O(\log N)$.
- `-2`: Pencarian biner. Rentang `lookup_array` **wajib** terurut turun (Descending / $Z-A$).

---

## 3. Matriks Komparasi Fitur Teknis

Tabel evaluasi teknis metode lookup pada core engine Excel:

| Dimensi Arsitektur | VLOOKUP | HLOOKUP | INDEX/MATCH | XLOOKUP |
| :--- | :--- | :--- | :--- | :--- |
| **Arah Traversal** | Vertikal (Kiri ke Kanan) | Horizontal (Atas ke Bawah) | Bebas (2 Vektor Bebas) | Bebas (Vertikal & Horizontal) |
| **Pencarian Mundur (Left/Upward)** | Tidak (Wajib trik `CHOOSE`) | Tidak | Ya | Ya (Native) |
| **Integritas Skema Data** | Rusak bila kolom geser | Rusak bila baris geser | Tahan perubahan layout | Tahan perubahan layout |
| **Default Matching** | `TRUE` (Approximate - Rawan) | `TRUE` (Approximate - Rawan) | `0` (Manual parameter) | `0` (Exact Match native) |
| **Error Handling Bawaan** | Butuh bungkus `IFNA()` | Butuh bungkus `IFNA()` | Butuh bungkus `IFNA()` | Terpasang (`[if_not_found]`) |
| **Output Array 2D (Spill)** | Tidak (Hanya skalar sel tunggal) | Tidak | Butuh formula array legasi | Ya (Native Multi-column spill) |
| **Algoritma Binary Search** | Terbatas (Implicit via `TRUE`) | Terbatas | Butuh perlakuan khusus | Ya (`search_mode` 2 / -2) |
| **Konsumsi Memori Cache** | Berat (Cache seluruh tabel) | Berat | Ringan (2 larik vektor) | Ringan (2 larik vektor) |
| **Dukungan Wildcard** | Ya | Ya | Ya | Ya (`match_mode = 2`) |

---

## 4. Studi Kasus 1: Ekstraksi Data Multi-Kolom Transaksi Gudang

### 4.1. Masalah
Sistem POS toko material catat master inventaris gudang. Analis butuh ambil data detail: `Nama Produk`, `Kategori`, `Stok`, dan `Harga Satuan` hanya masukkan parameter tunggal `SKU ID`. 

Pakai `VLOOKUP`: Butuh 4 formula berbeda di 4 sel terpisah. Perlu ubah nomor indeks kolom manual.
Pakai `XLOOKUP`: Cukup 1 formula tunggal. Manfaatkan fitur dynamic array spill.

### 4.2. Dataset Master Inventaris

Dataset tersimpan pada range `$A$1:$E$7` (Tabel referensi `tbl_Gudang`):

| Row | A (SKU_ID) | B (Nama_Barang) | C (Kategori) | D (Stok) | E (Harga_Satuan) |
|:---|:---|:---|:---|:---|:---|
| **1** | **SKU_ID** | **Nama_Barang** | **Kategori** | **Stok** | **Harga_Satuan** |
| **2** | SKU-001 | Semen Portland 50kg | Material | 120 | 65000 |
| **3** | SKU-002 | Cat Dinding Putih 5L | Finishing | 45 | 115000 |
| **4** | SKU-003 | Besi Beton 10mm | Struktur | 310 | 82000 |
| **5** | SKU-004 | Pipa PVC 3/4 Dim | Plumbing | 88 | 42000 |
| **6** | SKU-005 | Triplek 12mm | Kayu | 64 | 145000 |
| **7** | SKU-006 | Pasir Pasang 1 PickUp | Material | 12 | 650000 |

Area pencarian analisis target:
- Sel input kunci pencarian: `G2` (Input: `"SKU-003"`)
- Sel target formula: `H2` (Kolom target luapan: `H2:K2`)

```
Layout Ekstraksi Data:
[ Sel G2 ]  -> Input "SKU-003"
[ Sel H2 ]  -> [ Nama_Barang ] (Formula induk XLOOKUP ditulis di sini)
[ Sel I2 ]  -> [ Kategori ]    (Spill otomatis)
[ Sel J2 ]  -> [ Stok ]        (Spill otomatis)
[ Sel K2 ]  -> [ Harga_Satuan] (Spill otomatis)
```

### 4.3. Implementasi Formula

Formula ditulis pada sel `H2`:

```excel
=XLOOKUP(G2, $A$2:$A$7, $B$2:$E$7, "SKU Tidak Terdaftar", 0)
```

Jika data terformat sebagai Excel Structured Table (`Table1`):

```excel
=XLOOKUP(@G2, Table1[SKU_ID], Table1[[Nama_Barang]:[Harga_Satuan]], "SKU Tidak Terdaftar", 0)
```

### 4.4. Analisis Teknis Alur Eksekusi

```
Alur Eksekusi Formula:
1. Engine baca G2 ("SKU-003").
2. Traversal lookup_array ($A$2:$A$7):
   Row 2: "SKU-001" != "SKU-003"
   Row 3: "SKU-002" != "SKU-003"
   Row 4: "SKU-003" == "SKU-003" -> MATCH FOUND pada index baris ke-3.
3. Pointer kunci baris ke-3 diekstrapolasi ke return_array ($B$2:$E$7).
4. Sektor memory mengisolasi baris data 4: {"Besi Beton 10mm", "Struktur", 310, 82000}.
5. Alokasi Dynamic Array:
   - Nilai [0] -> Tulis ke sel H2 ("Besi Beton 10mm")
   - Nilai [1] -> Tulis ke sel I2 ("Struktur")
   - Nilai [2] -> Tulis ke sel J2 (310)
   - Nilai [3] -> Tulis ke sel K2 (82000)
6. Bounding box spill aktif. H2 bawa tanda formula aktif, I2:K2 terisi implisit ghost-formula (disabled editing).
```

Keuntungan teknis: Hilangkan ketergantungan hitung kolom manual. Perubahan posisi kolom `B` ke `C` di tabel sumber otomatis adaptif tanpa kalkulasi ulang manual.

---

## 5. Studi Kasus 2: Kondisi Ganda dan Backward Lookup

`XLOOKUP` tangani pencarian arah terbalik dan pemfilteran kriteria majemuk tanpa manipulasi struktur basis data.

### 5.1. Skenario 2A: Backward Lookup (Kunci Berada di Sisi Kanan Data)

Masalah: Temukan `ID_Karyawan` berdasarkan `No_KTP`. Pada tabel data legasi, kolom `No_KTP` ditaruh di kolom kanan, sedangkan `ID_Karyawan` ditaruh di kolom kiri. `VLOOKUP` gagal tanpa modifikasi urutan kolom sheet.

Dataset master karyawan (`$A$1:$C$6`):

| Row | A (ID_Karyawan) | B (Nama_Lengkap) | C (No_KTP) |
|:---|:---|:---|:---|
| **1** | **ID_Karyawan** | **Nama_Lengkap** | **No_KTP** |
| **2** | EMP-9011 | Bambang Hartono | 3201020405890001 |
| **3** | EMP-9012 | Siti Aminah | 3201020405900002 |
| **4** | EMP-9013 | Hendra Setiawan | 3201020405850003 |
| **5** | EMP-9014 | Ratna Juwita | 3201020405920004 |
| **6** | EMP-9015 | Doni Prasetyo | 3201020405880005 |

Nilai dicari: Sel `E2` berisi nilai `3201020405850003`.
Target ekstraksi: Sel `F2` mengambil data `ID_Karyawan` (Kolom A).

Formula sel `F2`:

```excel
=XLOOKUP(E2, $C$2:$C$6, $A$2:$A$6, "KTP Tidak Ditemukan", 0)
```

Mekanisme internal:
`lookup_array` diarahkan ke `$C$2:$C$6`. `return_array` diarahkan ke `$A$2:$A$6`. Engine Excel isolasi array independen. Posisi koordinat sumbu X kolom tidak dibaca. Hanya indeks array sumbu Y dicocokkan. Operasi backward lookup selesai tanpa overhead alokasi memory matrix.

---

### 5.2. Skenario 2B: Pencarian Kondisi Ganda (Multi-Condition Lookup)

Masalah: Temukan harga per liter pelumas mesin pada transaksi logistik berdasarkan kombinasi dua parameter variabel: `Tipe_Kendaraan` DAN `Kode_Viskositas`.

Dataset variasi pelumas (`$A$1:$D$7`):

| Row | A (Tipe_Kendaraan) | B (Kode_Viskositas) | C (Kemasan) | D (Harga_Liter) |
|:---|:---|:---|:---|:---|
| **1** | **Tipe_Kendaraan** | **Kode_Viskositas** | **Kemasan** | **Harga_Liter** |
| **2** | Diesel Truk | 15W-40 | Drum 200L | 48000 |
| **3** | Diesel Truk | 10W-40 | Drum 200L | 62000 |
| **4** | Bensin Mobil | 0W-20 | Botol 1L | 135000 |
| **5** | Bensin Mobil | 5W-30 | Galon 4L | 98000 |
| **6** | Motor Matic | 10W-30 | Botol 0.8L | 52000 |
| **7** | Motor Matic | 10W-40 | Botol 0.8L | 55000 |

Input kriteria pencarian:
- Sel `F2`: Tipe kendaraan (`"Diesel Truk"`)
- Sel `G2`: Viskositas (`"10W-40"`)
- Output sel `H2`: Mengambil `Harga_Liter`

Formula logika boolean sel `H2`:

```excel
=XLOOKUP(1, ($A$2:$A$7=F2) * ($B$2:$B$7=G2), $D$2:$D$7, "Kombinasi Spesifikasi Kosong", 0)
```

Bedah kalkulasi Boolean Array di memori heap:

```
Evaluasi Kondisi 1 ($A$2:$A$7 = "Diesel Truk"):
Row 2: TRUE
Row 3: TRUE
Row 4: FALSE
Row 5: FALSE
Row 6: FALSE
Row 7: FALSE
Array Kondisi 1 -> {TRUE; TRUE; FALSE; FALSE; FALSE; FALSE}

Evaluasi Kondisi 2 ($B$2:$B$7 = "10W-40"):
Row 2: FALSE
Row 3: TRUE
Row 4: FALSE
Row 5: FALSE
Row 6: FALSE
Row 7: TRUE
Array Kondisi 2 -> {FALSE; TRUE; FALSE; FALSE; FALSE; TRUE}

Perkalian Matriks Logika Boolean (Operasi AND -> TRUE=1, FALSE=0):
  {1; 1; 0; 0; 0; 0}  (Kondisi 1)
* {0; 1; 0; 0; 0; 1}  (Kondisi 2)
--------------------
= {0; 1; 0; 0; 0; 0}  -> Hasil array evaluasi intermediate

Evaluasi XLOOKUP:
lookup_value = 1
lookup_array = {0; 1; 0; 0; 0; 0}
Pencocokan nilai 1 ditemukan pada baris ke-2 (relatif terhadap dataset baris 3).
Ambil elemen baris ke-2 dari return_array $D$2:$D$7 -> Output: 62000.
```

Keunggulan metode: Eliminasi kebutuhan membuat kolom helper concate (`KolomA & KolomB`) pada lembar kerja sumber data.

---

### 5.3. Skenario 2C: Reverse Search (Audit Log Transaksi Terakhir)

Masalah: Buku kas logistik catat pergerakan kas masuk dan keluar berulang kali untuk vendor sama. Auditor ingin ekstrak transaksi **paling mutakhir** (data urutan baris terbawah) dari vendor tertentu tanpa mengurutkan ulang (sorting) tabel utama.

Dataset log transaksi mutasi kas (`$A$1:$D$8`):

| Row | A (Trx_ID) | B (Nama_Vendor) | C (Jenis_Mutasi) | D (Nominal) |
|:---|:---|:---|:---|:---|
| **1** | **Trx_ID** | **Nama_Vendor** | **Jenis_Mutasi** | **Nominal** |
| **2** | TRX-0101 | PT Abadi Jaya | Keluar | 1500000 |
| **3** | TRX-0102 | CV Sumber Makmur | Keluar | 4200000 |
| **4** | TRX-0103 | PT Abadi Jaya | Masuk | 800000 |
| **5** | TRX-0104 | UD Berkah Tani | Keluar | 1200000 |
| **6** | TRX-0105 | CV Sumber Makmur | Masuk | 3000000 |
| **7** | TRX-0106 | PT Abadi Jaya | Keluar | 2500000 |
| **8** | TRX-0107 | CV Mandiri Logistik | Keluar | 950000 |

Input pencarian:
- Sel `F2`: Nama vendor target (`"PT Abadi Jaya"`)
- Output sel `G2`: Ambil `Nominal` transaksi paling terakhir dibukukan.

Formula traversal balik sel `G2`:

```excel
=XLOOKUP(F2, $B$2:$B$8, $D$2:$D$8, "Data Mutasi Nihil", 0, -1)
```

Parameter `search_mode = -1` eksekusi algoritma traversal mundur:
- Pointer kalkulasi mulai pindai data dari baris paling bawah (`Row 8` -> `CV Mandiri Logistik`).
- Naik ke `Row 7` -> Terbaca `"PT Abadi Jaya"`.
- Pencocokan sukses (`PT Abadi Jaya` == `PT Abadi Jaya`).
- Mesin langsung hentikan scanning, ekstrak sel `$D$7` (`2500000`).
- Efisiensi: Scanning hanya telusuri 2 baris data dari bawah, hindari scanning dari baris teratas (lewat 6 baris data).

---

## 6. Dekonstruksi Kesalahan Teknis dan Root Cause

Penanganan bug dan exception saat deploy formula `XLOOKUP` pada file enterprise.

```
+---------------+-------------------------------------+---------------------------------------+
| Error Code    | Root Cause                          | Solusi Teknis                         |
+---------------+-------------------------------------+---------------------------------------+
| #SPILL!       | Blok sel output terhalang teks/data | Kosongkan rentang jalur sel output    |
| #N/A          | Format tipe data mismatch           | Normalisasi tipe data (VALUE / TRIM)  |
| #VALUE!       | Dimensi array tidak simetris        | Samakan panjang vektor 1D rentang     |
| Memory Bloat  | Referensi whole-column (A:A)        | Gunakan Dynamic Range / Excel Tables  |
+---------------+-------------------------------------+---------------------------------------+
```

### 6.1. Kasus Error `#SPILL!`

```
Penyebab: 
Formula return_array menghasilkan multi-kolom atau multi-baris dinamis, 
namun terdapat sel terisi data, karakter spasi, atau merge cells 
pada blok grid lintasan output.
```

Gejala: Formula sel `H2` kembalikan error `#SPILL!`.

Diagnosa:
Formula butuh 4 kolom kosong (`H2:K2`). Jika sel `J2` tidak sengaja terisi karakter spasi atau nilai statis, engine dynamic array batalkan render data seluruh blok untuk cegah overwrite data tanpa izin user.

Solusi:
1. Klik sel yang berstatus `#SPILL!`.
2. Perhatikan indikator batas garis putus-putus biru pada lembar kerja (bounding box).
3. Hapus seluruh data statis yang menghalangi blok bounding box.
4. Unmerge seluruh sel yang masuk dalam cakupan target output data.

### 6.2. Kasus Error `#N/A` (Data Format Mismatch)

```
Penyebab: 
Nilai lookup_value dan nilai lookup_array tampak identik di layar,
namun beda representasi biner pada memory layer 
(String Teks vs Integer Biner atau Whitespace tak kasat mata).
```

Diagnosa Kasus:
- Sel `A2` berisi nilai teks: `'10023` (String).
- Sel `lookup_value` berisi angka: `10023` (Integer numeric).
- Komparasi biner `10023 == "10023"` hasilkan nilai `FALSE`. `XLOOKUP` kembalikan error `#N/A`.

Solusi sanitasi format langsung di dalam formula:

```excel
=XLOOKUP(TRIM(TEXT(G2, "@")), TRIM($A$2:$A$100), $B$2:$B$100, "Tidak Ditemukan", 0)
```

Atau normalisasi teks ke integer numeric pakai operator double unary (`--`):

```excel
=XLOOKUP(--G2, --($A$2:$A$100), $B$2:$B$100, "Tidak Ditemukan", 0)
```

Pembersihan karakter non-breaking space (ASCII 160) asal sistem ERP web:

```excel
=XLOOKUP(SUBSTITUTE(G2, CHAR(160), ""), SUBSTITUTE($A$2:$A$100, CHAR(160), ""), $B$2:$B$100, "Data Bersih", 0)
```

### 6.3. Kasus Error `#VALUE!` (Asymmetric Vector Dimensions)

```
Penyebab:
Panjang rentang larik antara lookup_array dan return_array tidak sinkron
(Jumlah baris atau kolom tidak sama persis).
```

Contoh formula pemicu error:

```excel
=XLOOKUP(G2, $A$2:$A$100, $B$2:$B$80)
```

Analisis kegagalan:
- `lookup_array` memiliki panjang 99 elemen baris ($A2:A100).
- `return_array` hanya memiliki panjang 79 elemen baris ($B2:B80).
- Engine pemetaan memori gagal proyeksikan pasangan offset indeks. Excel langsung lempar eksepsi kode `#VALUE!`.

Solusi:
Pastikan batas awal baris dan batas akhir baris terkunci pada angka identik:

```excel
=XLOOKUP(G2, $A$2:$A$100, $B$2:$B$100)
```

### 6.4. Isu Performa: Whole-Column Reference Memory Bloat

Kesalahan fatal pemula spreadsheet enterprise:

```excel
=XLOOKUP(G2, A:A, B:B)
```

Analisis komputasi:
Referensi kolom penuh `A:A` paksa Excel siapkan alokasi pointer metadata kalkulasi sebanyak 1.048.576 baris data kosong. Jika formula ini disalin ke 20.000 baris sheet, file alami gejala *laggy*, calculation thread hung, dan file size membengkak puluhan megabyte tanpa isi.

Solusi optimasi:
Pakai referensi batas eksplisit rentang data aktif:

```excel
=XLOOKUP(G2, $A$2:$A$15000, $B$2:$B$15000)
```

Atau konversi range data menjadi format native Excel Table (`Ctrl + T`):

```excel
=XLOOKUP(G2, MasterData[ID], MasterData[Nilai])
```

---

## 7. Standar Implementasi dan Best Practice Checklist

Gunakan daftar kepatuhan teknis ini sebelum merilis dokumen spreadsheet ke lingkungan produksi:

```
[ ] 1. HAPUS FUNGSI WRAPPER REDUNDAN
       Tinggalkan pola lama pemanggilan formula:
       =IFERROR(XLOOKUP(...), "Kosong") -> TIDAK EFISIEN!
       Gunakan parameter resmi ke-4:
       =XLOOKUP(..., "Kosong")           -> NATIVE RUNTIME HANDLING.

[ ] 2. PRIORITASKAN PENGGUNAAN EXCEL STRUCTURED REFERENCES
       Format data mentah sebagai Excel Table (ListObject).
       Ganti:
       =XLOOKUP(A2, $D$2:$D$5000, $E$2:$E$5000)
       Menjadi:
       =XLOOKUP(A2, tbl_Transaksi[SKU], tbl_Transaksi[Harga])
       Alasan: Mengeliminasi risiko formula asymmetric dimension saat append data baru.

[ ] 3. HINDARI PENGGUNAAN WHOLE COLUMN REFERENCES
       Jangan gunakan parameter referensi baris penuh (A:A, B:B).
       Kunci rentang referensi hanya pada area efektif sel data.

[ ] 4. OPTIMASI AKSES DATA BESAR DENGAN BINARY SEARCH
       Data transaksi historis berukuran >500.000 baris:
       - Lakukan sorting ascending (A-Z) pada kolom kunci.
       - Aktifkan search_mode = 2.
       - Kompleksitas waktu terpangkas dari O(N) ke O(log N). Kalkulasi instan.

[ ] 5. CEK INTERSEKSI SPILL RANGE
       Pastikan sel sebelah kanan formula bebas dari data statis apa pun
       jika return_array melibatkan multi-kolom matriks.

[ ] 6. PERIKSA DISTRIBUSI LISENSI PENGGUNA TERAKHIR
       XLOOKUP hanya jalan di:
       - Microsoft 365 (Desktop, Web, Mac)
       - Office 2021
       - Office 2024
       Jika workbook dikirim ke klien pengguna Excel 2019, Excel 2016, 
       atau Excel 2013, formula hasilkan error fatal #NAME?. 
       Tetap gunakan INDEX/MATCH untuk backwards compatibility jika target 
       lingkungan sistem belum terstandardisasi.
```

Kerapian arsitektur data spreadsheet percepat laju analitik bisnis, cegah anomali kalkulasi finansial, dan minimalisir downtime akibat debugging formula warisan yang rapuh. `XLOOKUP` standarisasi alur kerja lookup Excel modern secara optimal, stabil, dan presisi tinggi.