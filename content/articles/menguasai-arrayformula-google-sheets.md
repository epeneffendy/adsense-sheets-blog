---
title: "Menguasai ARRAYFORMULA Google Sheets untuk Otomasi Kalkulasi Otomatis"
description: "Panduan praktis menggunakan ARRAYFORMULA di Google Sheets: kalkulasi ratusan baris secara otomatis dari 1 sel tanpa perlu drag-and-drop manual."
category: "Google Sheets"
date: "2023-10-25"
readTime: "14 min"
---

## 1. Pendahuluan dan Latar Belakang

Model komputasi spreadsheet tradisional andalkan replikasi sel skalar. User tulis formula pada sel target. User seret *fill handle* ke bawah. Cara ini ciptakan duplikasi metadata identik ribuan kali dalam file.

```
Traditional:
[Row 2] Cell C2 = A2 * B2  --> Parse tree C2 simpan string "=A2*B2"
[Row 3] Cell C3 = A3 * B3  --> Parse tree C3 simpan string "=A3*B3"
...
[Row N] Cell CN = AN * BN  --> Parse tree CN simpan string "=AN*BN"
```

Masalah operasional muncul cepat:
* **Memory Bloat**: Google Sheets alokasikan memori untuk tiap instance formula. Ratusan ribu sel dengan formula terduplikasi picu konsumsi RAM browser tinggi. UI freeze sering terjadi.
* **Integrity Drift**: User tidak sengaja edit satu sel tengah kolom. Formula baris rusak. Audit logika jadi sulit.
* **Form & API Sync Failure**: Data masuk via Google Forms atau webhook AppSheet sisip baris baru (*row insertion*). Baris baru geser range target. Rumus seret manual tidak eksekusi baris baru.
* **Network Payload Besar**: Download JSON payload sheet lewat Google Drive API kirimkan formula string berulang tiap baris. Bandwidth boros.

Solusi: Array Processing Engine. Engine bangun array multidimensi dalam memory buffer tunggal. Output tumpah (*spill*) ke sel-sel bawah atau samping secara virtual.

`ARRAYFORMULA` transformasikan operasi skalar jadi eksekusi vektor:

```
Array Model:
[Row 2] Cell C2 = ARRAYFORMULA(A2:A * B2:B)
        |
        +--> Memory buffer alokasikan 1 formula tree.
        +--> Output ditumpahkan ke C2:C otomatis.
        +--> C3:CN kosong dari teks formula, data tetap muncul.
```

Karakteristik komputasi:
* Single point of maintenance. Edit sel master ubah kalkulasi seluruh kolom.
* Data baru masuk via append langsung masuk kalkulasi tanpa intervensi manual.
* Ukuran file XML internal Google Sheets ramping drastis.

---

## 2. Anatomi Sintaks Fungsi ARRAYFORMULA

### Sintaks Dasar

```excel
=ARRAYFORMULA(array_formula)
```

Shortcut keyboard terapkan fungsi cepat:
* Windows / Linux: `Ctrl + Shift + Enter`
* macOS: `Cmd + Shift + Return`

Engine Google Sheets parse argument `array_formula` sebagai rentang array. Scalar operator diubah paksa jadi array operator.

### Parameter Breakdown

* `array_formula`: Ekspresi matematika, rentang sel (*range*), atau kombinasi fungsi yang dukung input array. Menerima rentang 1D (misal `A2:A100`) atau rentang 2D (misal `A2:B100`).

### Karakteristik Eksekusi

#### Operasi Vektor Sejajar (Element-wise Operation)
Operasi antara dua rentang dengan dimensi sama dieksekusi paralel baris demi baris:
```excel
=ARRAYFORMULA(A2:A10 + B2:B10)
```
Engine ambil elemen `A2` tambah `B2`, tumpahkan ke output index 0. Lanjut `A3` tambah `B3`, tumpahkan ke index 1, sampai batas index array tercapai.

#### Scalar Expansion (Broadcasting)
Jika satukan array dengan nilai tunggal (skalar), engine duplikasi nilai skalar secara virtual ke seluruh panjang array:
```excel
=ARRAYFORMULA(A2:A10 * 0.11)
```
Nilai `0.11` dikalikan ke setiap elemen `A2` hingga `A10`.

#### Penanganan Array Literals `{}`
Array literal bangun struktur data baru tanpa sel fisik:
* Koma `,` bangun pemisah kolom (horizontal array).
* Titik koma `;` bangun pemisah baris (vertical array).

Sematkan header langsung dari formula root:
```excel
=ARRAYFORMULA({"Total Harga"; A2:INDEX(A:A, COUNTA(A:A)) * B2:INDEX(B:B, COUNTA(B:B))})
```
Formula atas taruh teks "Total Harga" di baris 1. Baris 2 ke bawah isi hasil perkalian. Seluruh kolom kendali penuh dari sel baris 1. Baris 2 ke bawah kebal timpa formula tak sengaja.

### Inkompatibilitas Fungsi Agregasi Bawaan

Fungsi agregasi skalar Google Sheets runtuh dalam `ARRAYFORMULA`:
* `AND()`: Evaluasi seluruh rentang jadi satu output skalar Boolean tunggal (`TRUE` atau `FALSE`). Tidak evaluasi baris per baris.
* `OR()`: Evaluasi seluruh sel range, hasilkan satu output Boolean skalar tunggal.
* `SUM()`: Jumlahkan seluruh angka jadi satu angka tunggal, bukan akumulasi baris demi baris.
* `AVERAGE()`: Hasilkan rata-rata total rentang.

Pengganti operasi logika skalar:
* Ganti `AND` pakai perkalian matematika Boolean: `*`
* Ganti `OR` pakai penjumlahan matematika Boolean: `+`
* Ganti `NOT` pakai negasi biner: `1 - (kondisi)`

Evaluasi Boolean internal:
* `TRUE` bernilai numerik `1`
* `FALSE` bernilai numerik `0`

---

## 3. Komparasi: Formula Tradisional vs ARRAYFORMULA

| Dimensi Evaluasi | Formula Tradisional (Drag-and-Drop) | ARRAYFORMULA (Array Vectorized) |
| :--- | :--- | :--- |
| **Penyimpanan Metadata** | Berat. Formula string disimpan pada setiap sel fisik baris. | Ramping. Formula disimpan hanya di 1 sel root. |
| **Konsumsi Memory Client** | Alokasikan DOM node dan parse tree terpisah per sel. | Alokasikan satu output buffer array. Rendah overhead engine. |
| **Keamanan Logika** | Rentan. User dapat timpa rumus di baris 45 tanpa deteksi visual. | Kuat. Sel anak dilindungi; user hapus data sel anak langsung ditolak/spill balik. |
| **Respon Input Dinamis** | Butuh drag ulang manual atau auto-fill script tiap ada record baru. | Otomatis kalkulasi saat sel target terisi nilai baru. |
| **Kompatibilitas Import** | Bebani proses sinkronisasi CSV via third-party integration. | Data stream via API langsung diproses real-time. |
| **Audit dan Debugging** | Lambat. Harus verifikasi inkonsistensi formula di sepanjang baris. | Cepat. Cukup periksa validitas satu formula pada sel root. |
| **Batas Spill Area** | Tidak ada konsep spill. Rentan timpah data diam-diam. | Blokir output dan lempar error `#REF!` jika jalur sel terhalang. |

---

## 4. Studi Kasus Nyata 1: Sistem Faktur Penjualan Multi-Item

### Permasalahan Bisnis

Sistem POS e-commerce kirim log transaksi ke sheet penjualan. Data baru masuk terus via API. 

Kebutuhan perhitungan:
1. Hitung Subtotal Kotor: `Qty * Harga Satuan`.
2. Hitung Diskon Progresif:
   * Jika Qty >= 100, diskon 15%.
   * Jika Qty >= 50, diskon 10%.
   * Jika Qty < 50, diskon 0%.
3. Hitung Subtotal Bersih: `Subtotal Kotor - Diskon`.
4. Hitung PPN 11%: `Subtotal Bersih * 0.11`.
5. Hitung Total Akhir: `Subtotal Bersih + PPN`.
6. Cegah kalkulasi pada baris kosong. Range terbuka `A2:A` tanpa guard clause cetak angka 0 hingga baris batas sheet (10.000+ baris). Kondisi ini buat sheet berat tak perlu.

### Skema Dataset (`Sheet_Sales`)

| Kolom A (Invoice_ID) | Kolom B (Item_Name) | Kolom C (Qty) | Kolom D (Price) |
| :--- | :--- | :--- | :--- |
| INV-1001 | Keyboard Mekanikal | 120 | 850000 |
| INV-1002 | Mouse Wireless | 35 | 320000 |
| INV-1003 | Monitor 24 Inch | 60 | 1950000 |
| INV-1004 | USB Hub Type-C | 15 | 150000 |
| INV-1005 | Web Camera 1080p | 80 | 650000 |
| INV-1006 | Headset Gaming | 45 | 450000 |
| INV-1007 | Mousepad Deskmat | 110 | 120000 |
| INV-1008 | Standing Desk | 10 | 3200000 |
| INV-1009 | Kursi Ergonomis | 52 | 2100000 |
| INV-1010 | Lampu Monitor | 25 | 280000 |

### Implementasi Formula Kolom Mandiri

Tempatkan formula pada sel baris 2 masing-masing kolom output.

#### Kolom E: Subtotal Kotor
```excel
=ARRAYFORMULA(IF(ISBLANK(A2:A), "", C2:C * D2:D))
```

#### Kolom F: Nilai Diskon
```excel
=ARRAYFORMULA(IF(ISBLANK(A2:A), "", (C2:C * D2:D) * IF(C2:C >= 100, 0.15, IF(C2:C >= 50, 0.1, 0))))
```

#### Kolom G: Subtotal Bersih
```excel
=ARRAYFORMULA(IF(ISBLANK(A2:A), "", (C2:C * D2:D) - ((C2:C * D2:D) * IF(C2:C >= 100, 0.15, IF(C2:C >= 50, 0.1, 0)))))
```

#### Kolom H: PPN 11%
```excel
=ARRAYFORMULA(IF(ISBLANK(A2:A), "", ((C2:C * D2:D) - ((C2:C * D2:D) * IF(C2:C >= 100, 0.15, IF(C2:C >= 50, 0.1, 0)))) * 0.11))
```

#### Kolom I: Total Akhir (Master Consolidated Column)
Formula ringkas satukan seluruh logika transaksi langsung dari sel root `I2`:
```excel
=ARRAYFORMULA(
  IF(ISBLANK(A2:A), "",
    LET(
      gross, C2:C * D2:D,
      discount_rate, IF(C2:C >= 100, 0.15, IF(C2:C >= 50, 0.10, 0)),
      discount_val, gross * discount_rate,
      net, gross - discount_val,
      vat, net * 0.11,
      net + vat
    )
  )
)
```

### Bedah Logika Eksekusi

```
Pemeriksaan Evaluasi Sel:
Baris 2: INV-1001
  ISBLANK(A2) -> FALSE.
  gross       -> 120 * 850000 = 102000000
  discount    -> C2 (120) >= 100 -> 0.15 * 102000000 = 15300000
  net         -> 102000000 - 15300000 = 86700000
  vat         -> 86700000 * 0.11 = 9537000
  total       -> 86700000 + 9537000 = 96237000. Tumpah ke I2.

Baris 11: Kosong (EOF)
  ISBLANK(A11) -> TRUE.
  Output      -> "" (String kosong). Array berhenti alokasi memori kalkulasi.
```

Kombinasi `ARRAYFORMULA` dan `LET` eliminasi redundansi kalkulasi berulang. Evaluasi subtotal kotor hanya dihitung satu kali per elemen di tingkat memori cache internal.

---

## 5. Studi Kasus Nyata 2: Multi-Criteria Dynamic Reverse Lookup

### Permasalahan Bisnis

Tabel inventaris logistik catat riwayat mutasi stok. Kolom paling kiri bukan primary key. `VLOOKUP` standar gagal lakukan pencarian ke arah kiri (*reverse lookup*) tanpa modifikasi array virtual `{}`. 

`VLOOKUP` standar juga tidak mampu evaluasi kondisi ganda (Multi-Criteria) secara native tanpa helper column.

Tantangan: 
Ambil data `SKU_CODE` (Kolom A) berdasarkan kombinasi dua kriteria input dinamis:
1. `Warehouse_ID` (Kolom C) cocok dengan sel target pencarian.
2. `Status` (Kolom D) bernilai `"Available"`.
3. `Stock_Qty` (Kolom E) lebih besar dari ambang batas minimum.

### Skema Dataset Logistik (`Sheet_Inventory`)

| Kolom A (SKU_Code) | Kolom B (Batch_ID) | Kolom C (Warehouse) | Kolom D (Status) | Kolom E (Qty) |
| :--- | :--- | :--- | :--- | :--- |
| SKU-JKT-01 | B-901 | WH-WEST | Available | 450 |
| SKU-SBY-02 | B-902 | WH-EAST | Out_Of_Stock | 0 |
| SKU-BDG-03 | B-903 | WH-WEST | Damaged | 12 |
| SKU-SMG-04 | B-904 | WH-CENTRAL | Available | 85 |
| SKU-JKT-05 | B-905 | WH-WEST | Available | 1200 |
| SKU-SBY-06 | B-906 | WH-EAST | Available | 30 |
| SKU-BDG-07 | B-907 | WH-CENTRAL | Quarantined | 500 |
| SKU-JKT-08 | B-908 | WH-WEST | Available | 75 |

### Parameter Pencarian

* Target Gudang: Sel `G2` = `"WH-WEST"`
* Target Ambang Stok Minimum: Sel `H2` = `100`

### Pendekatan 1: Boolean Filter Array (High Performance)

```excel
=ARRAYFORMULA(
  FILTER(
    A2:A, 
    (C2:C = G2) * (D2:D = "Available") * (E2:E >= H2)
  )
)
```

### Pendekatan 2: Reverse Lookup Indexing Native via Boolean Matching

Jika sheet harus return baris pertama yang cocok ke dalam template laporan tunggal:

```excel
=ARRAYFORMULA(
  INDEX(
    A2:A, 
    MATCH(
      1, 
      (C2:C = G2) * (D2:D = "Available") * (E2:E >= H2), 
      0
    )
  )
)
```

### Pendekatan 3: Reverse VLOOKUP Virtual Matrix Injection

Jika struktur formula harus pertahankan fungsi `VLOOKUP`:

```excel
=ARRAYFORMULA(
  VLOOKUP(
    G2 & "_" & "Available",
    {C2:C & "_" & D2:D, A2:A},
    2,
    FALSE
  )
)
```

### Analisis Mekanisme Eksekusi Boolean

Lihat proses resolusi aljabar matriks ekspresi `(C2:C = G2) * (D2:D = "Available") * (E2:E >= H2)`:

```
Step 1: Evaluasi Vektor 1 (Warehouse Match G2 = "WH-WEST")
Row 2: "WH-WEST"    = "WH-WEST" -> TRUE  (1)
Row 3: "WH-EAST"    = "WH-WEST" -> FALSE (0)
Row 4: "WH-WEST"    = "WH-WEST" -> TRUE  (1)
Row 5: "WH-CENTRAL" = "WH-WEST" -> FALSE (0)
Row 6: "WH-WEST"    = "WH-WEST" -> TRUE  (1)
Row 7: "WH-EAST"    = "WH-WEST" -> FALSE (0)
Row 8: "WH-CENTRAL" = "WH-WEST" -> FALSE (0)
Row 9: "WH-WEST"    = "WH-WEST" -> TRUE  (1)
Vektor 1 = {1; 0; 1; 0; 1; 0; 0; 1}

Step 2: Evaluasi Vektor 2 (Status = "Available")
Row 2: "Available"   = "Available" -> 1
Row 3: "Out_Of_Stock"= "Available" -> 0
Row 4: "Damaged"     = "Available" -> 0
Row 5: "Available"   = "Available" -> 1
Row 6: "Available"   = "Available" -> 1
Row 7: "Available"   = "Available" -> 1
Row 8: "Quarantined" = "Available" -> 0
Row 9: "Available"   = "Available" -> 1
Vektor 2 = {1; 0; 0; 1; 1; 1; 0; 1}

Step 3: Evaluasi Vektor 3 (Stock >= 100)
Row 2: 450  >= 100 -> 1
Row 3: 0    >= 100 -> 0
Row 4: 12   >= 100 -> 0
Row 5: 85   >= 100 -> 0
Row 6: 1200 >= 100 -> 1
Row 7: 30   >= 100 -> 0
Row 8: 500  >= 100 -> 1
Row 9: 75   >= 100 -> 0
Vektor 3 = {1; 0; 0; 0; 1; 0; 1; 0}

Step 4: Perkalian Hadamard Antar Vektor (Bitwise AND)
  {1; 0; 1; 0; 1; 0; 0; 1}  (Warehouse)
* {1; 0; 0; 1; 1; 1; 0; 1}  (Status)
* {1; 0; 0; 0; 1; 0; 1; 0}  (Stock)
--------------------------------------
= {1; 0; 0; 0; 1; 0; 0; 0}

Hasil Filter:
Index 0 (Row 2): 1 -> Ambil A2: "SKU-JKT-01"
Index 4 (Row 6): 1 -> Ambil A6: "SKU-JKT-05"
Output Array tumpahkan 2 SKU yang valid.
```

Metode aljabar Boolean eliminasi kebutuhan kolom helper manual. Kalkulasi terjadi murni di memori RAM tanpa rendering visual sementara.

---

## 6. Kesalahan Umum dan Troubleshooting Deep-Dive

### 1. Error `#REF!` (Spill Range Blocked)

```
Error details:
"Array result was not expanded because it would overwrite data in [Target Cell]."
```

Penyebab: 
ARRAYFORMULA butuh area kosong untuk tumpahkan data. Sel di bawahnya isi teks, angka, spasi tersembunyi, atau karakter non-cetak.

Mitigasi dan Penanganan:
1. Identifikasi koordinat sel penghalang dari tooltip error.
2. Kosongkan seluruh sel di bawah formula root sampai batas data.
3. Cek spasi kosong tidak sengaja pakai shortcut: Tekan `Ctrl + Shift + Down` dari bawah sel root, tekan `Delete`.
4. Jika baris bawah perlu catatan kaki, batasi range output formula:
   ```excel
   =ARRAYFORMULA(A2:INDEX(A:A, 100) * B2:INDEX(B:B, 100))
   ```

### 2. Error `#VALUE!` (Mismatched Dimensions)

```
Error details:
"Function ARRAYFORMULA parameter 1 expects array inputs of equal size."
```

Penyebab:
Operasi array libatkan dua atau lebih rentang dengan dimensi baris atau kolom beda panjang.

Contoh salah:
```excel
=ARRAYFORMULA(A2:A100 * B2:B50)  <-- Crash. Baris 100 dikalikan baris 50.
```

Aturan dimensi Google Sheets:
* Panjang rentang vertikal harus identik: `A2:A` dengan `B2:B`.
* Jika pakai skalar, skalar harus sel tunggal absolut: `A2:A * $D$1`.

Perbaikan:
```excel
=ARRAYFORMULA(A2:A100 * B2:B100)
```

### 3. Logical Collapse pada Fungsi Agregator Tradisional

Contoh salah penggunaan `AND`/`OR`:
```excel
=ARRAYFORMULA(IF(AND(A2:A > 0, B2:B > 0), "Valid", "Invalid"))
```

Hasil runtime:
Output cuma 1 sel di posisi root. Seluruh baris bawah kosong atau salah evaluasi. Fungsi `AND()` telan seluruh range `A2:A` dan `B2:B`, kompresi range jadi satu nilai Boolean tunggal untuk seluruh sheet.

Solusi: Konversi ke Aljabar Boolean:
```excel
=ARRAYFORMULA(IF((A2:A > 0) * (B2:B > 0), "Valid", "Invalid"))
```

Konversi Operator:

| Operasi Logika | Sintaks Konvensional (Gagal) | Sintaks Vektor Valid |
| :--- | :--- | :--- |
| **Logika AND** | `AND(Cond1, Cond2)` | `(Cond1) * (Cond2)` |
| **Logika OR** | `OR(Cond1, Cond2)` | `(Cond1) + (Cond2) > 0` |
| **Logika NOT** | `NOT(Cond1)` | `1 - (Cond1)` |
| **Logika XOR** | `XOR(Cond1, Cond2)` | `MOD((Cond1) + (Cond2), 2)` |

### 4. Overcomputation Spike (Range Terbuka Tanpa Filter)

Contoh masalah:
```excel
=ARRAYFORMULA(A2:A * B2:B)
```

Dampak sistemik:
Sheet punya 50.000 baris. Data isi hanya baris 2 sampai 100. Formula di atas paksa engine alokasikan perkalian skalar untuk 49.900 baris kosong. File lemot, konsumsi API kuota bengkak, kalkulasi mutar terus (*calculating loop*).

Solusi proteksi awal baris kosong (*Guard Clause*):

```excel
=ARRAYFORMULA(IF(LEN(A2:A) = 0, "", A2:A * B2:B))
```

Fungsi `LEN(A2:A) = 0` lebih cepat dieksekusi engine Google Sheets dibanding `ISBLANK(A2:A)`. Karakter evaluasi string kosong dihentikan pada siklus pertama komputasi.

### 5. Error Penanganan Nilai Missing (`#N/A`)

Penyebab:
Operasi `VLOOKUP` atau `MATCH` dalam array temukan data yang tidak terdaftar di master tabel. Error `#N/A` tumpah ke sel anak dan stop rantai kalkulasi downstream.

Solusi pembungkusan protektif:
```excel
=ARRAYFORMULA(
  IF(LEN(A2:A) = 0, "", 
    IFNA(
      VLOOKUP(A2:A, MasterData!A:B, 2, FALSE), 
      "Data Tidak Ditemukan"
    )
  )
)
```

Gunakan `IFNA()` daripada `IFERROR()`. `IFNA()` tangkap data tidak ketemu saja. `IFERROR()` telan error sintaks fatal lain sehingga bug logika tersembunyi.

---

## 7. Checklist dan Best Practice Standar Industri

Terapkan standard operating procedure berikut sebelum rilis spreadsheet ke environment produksi.

### A. Konvensi Header Dinamis
Jangan letakkan `ARRAYFORMULA` di baris data (`Row 2`). Taruh di baris header (`Row 1`) pakai array literal vertikal `{}`.

```excel
=ARRAYFORMULA(
  {
    "Nama Kolom Header"; 
    IF(LEN(A2:A) = 0, "", A2:A * B2:B)
  }
)
```
Keuntungan:
* Baris 2 sepenuhnya berisi hasil output tumpahan virtual.
* User yang tidak sengaja hapus baris 2 tidak akan rusak logika formula sistem.
* Formula terkunci aman di level Header (Row 1).

### B. Optimalisasi Eksekusi dan Memory Profiling
1. **Gunakan Batas Range Riil Bila Memungkinkan**:
   Jika rentang transaksi bulanan pasti di bawah 1.000 baris, ganti `A2:A` jadi `A2:A1000`. Stop engine alokasi virtual cell tak terpakai sampai batas 10.000.000 limit sel Google Sheets.
2. **Kombinasikan `LET` untuk Rumus Kompleks**:
   Hindari hitung ekspresi array identik dua kali. Simpan di variabel lokal engine:
   ```excel
   =ARRAYFORMULA(
     LET(
       qty, A2:A,
       harga, B2:B,
       subtotal, qty * harga,
       IF(LEN(qty)=0, "", subtotal + (subtotal * 0.11))
     )
   )
   ```
3. **Audit Ketergantungan Sirkular (Circular Dependency)**:
   Hindari formula di sel `C1` mereferensikan kolom `C:C`. Selalu batasi referensi input dari kolom lain.

### C. Standard Operasional Troubleshooting (SOP Debugging)

Jika formula hasilkan error atau perilaku lemot:
1. **Spill check**: Navigasi ke sel root error `#REF!`. Hapus isi sel di bawahnya.
2. **Dimension check**: Cocokkan batas awal dan akhir range:
   * Valid: `A2:A50` dan `B2:B50`
   * Rusak: `A2:A50` dan `B2:B51`
3. **Data type check**: Pastikan kolom input angka tidak diformat sebagai plain text. Karakter teks pada perkalian array hasilkan error runtime `#VALUE!`.
4. **Boolean check**: Hapus fungsi `AND()` atau `OR()`. Ganti operator `*` atau `+`.

### D. Lembar Kendali Mutu Produksi (Pre-Deployment Checklist)

* [ ] Formula master diletakkan pada Baris 1 (`Row 1`) dengan array literal header `{}`.
* [ ] Seluruh input range terbuka dibungkus pelindung `IF(LEN(...) = 0, "", ...)`.
* [ ] Tidak ada fungsi agregasi tunggal (`SUM`, `AND`, `OR`, `AVERAGE`) di dalam ekspresi vektor.
* [ ] Error penelusuran lookup ditangani oleh `IFNA()`.
* [ ] Format kolom output disetel via menu format Google Sheets, bukan di-cast manual via formula text string.
* [ ] Kolom target tumpahan bersih dari data residu manual. Error `#REF!` nihil.