---
title: "Menampilkan Data Dinamis dengan Fungsi FILTER dan SORTN Google Sheets"
description: "Filter data multi-kriteria dan tampilkan ranking Top 5 atau Top 10 transaksi secara otomatis tanpa perlu mengubah tabel sumber."
category: "Google Sheets"
date: "2026-10-28"
readTime: "8 menit baca"
---

## Pendahuluan

Pengelolaan data transaksi berskala besar pada Google Sheets menuntut pemisahan tegas antara tabel data mentah (*raw data*) dan lapisan pelaporan (*reporting layer*). Praktik konvensional berbasis manipulasi langsung—seperti pengurutan manual melalui menu GUI (*Data > Sort range*) atau penyalinan baris secara manual—menimbulkan fragmentasi data, merusak urutan kronologis asli, serta meningkatkan risiko *human error*. 

Laporan operasional membutuhkan daftar peringkat berkala, seperti 5 atau 10 transaksi terbesar berdasarkan segmen spesifik. Kebutuhan ini harus terpenuhi secara dinamis: setiap penambahan baris transaksi baru pada tabel sumber wajib terrefleksi langsung pada tabel ringkasan tanpa intervensi pengguna. 

Kombinasi fungsi `FILTER` dan `SORTN` membentuk kalkulasi *dynamic array* terisolasi. Formula membaca data mentah, mengevaluasi kriteria multi-kondisi, menyortir nilai, dan memotong hasil sesuai limit kuantitas yang ditentukan dalam satu proses komputasi di memori tanpa mengubah struktur tabel sumber.

---

## Anatomi Sintaks dan Argumen Teknis

Integrasi formula ini memanfaatkan output matriks `FILTER` sebagai input data mentah bagi fungsi `SORTN`.

```excel
=SORTN(FILTER(range; kondisi1; [kondisi2; ...]); [n]; [display_ties_mode]; [indeks_kolom_sort]; [is_ascending])
```

### 1. Fungsi FILTER

Fungsi `FILTER` mengekstrak subset data yang memenuhi parameter pengujian logika boolean.

```excel
FILTER(range; kondisi1; [kondisi2; ...])
```

*   **`range`**: Rentang data sumber yang akan diekstraksi (contoh: `A2:E100`).
*   **`kondisi1`**: Rentang evaluasi satu kolom dengan dimensi baris identik terhadap `range`, menghasilkan array boolean (`TRUE` atau `FALSE`).
*   **`[kondisi2; ...]`**: Parameter kondisi tambahan (opsional). Kondisi tambahan yang dipisahkan tanda pemisah argumen diperlakukan dengan operasi logika `AND`.

Evaluasi logika array pada `FILTER` bekerja berdasarkan aljabar boolean:
*   Logika **AND**: Mengalikan matriks kondisi `(Kondisi_A) * (Kondisi_B)`.
*   Logika **OR**: Menjumlahkan matriks kondisi `(Kondisi_A) + (Kondisi_B)`.

### 2. Fungsi SORTN

Fungsi `SORTN` memperluas kapabilitas `SORT` standar dengan membatasi jumlah output data (*N-records*) dan mengendalikan perilaku terhadap nilai duplikat (*ties*).

```excel
SORTN(range; [n]; [display_ties_mode]; [sort_kolom1]; [is_ascending1]; ...)
```

*   **`range`**: Matriks data yang akan diurutkan dan dibatasi. Dalam arsitektur ini, parameter diisi oleh keluaran formula `FILTER`.
*   **`n`**: (Opsional, *default* = 1) Jumlah record baris data yang ingin ditampilkan.
*   **`display_ties_mode`**: (Opsional, *default* = 0) Integer penentu mitigasi benturan nilai kembar (*ties*):
    *   `0`: Tampilkan maksimal `n` baris secara ketat tanpa toleransi nilai kembar.
    *   `1`: Tampilkan `n` baris plus seluruh baris tambahan yang memiliki nilai identik dengan baris ke-`n`.
    *   `2`: Tampilkan `n` baris setelah mengeliminasi baris duplikat (deduplikasi data).
    *   `3`: Tampilkan `n` baris unik beserta seluruh duplikat dari record tersebut.
*   **`sort_kolom1`**: Indeks numerik berbasis 1 (*1-based index*) yang merujuk pada posisi kolom di dalam rentang input, bukan referensi kolom lembar kerja global.
*   **`is_ascending1`**: Arah pengurutan. Bernilai `TRUE` untuk urutan menaik (*ascending* / A-Z / 0-9) atau `FALSE` untuk urutan menurun (*descending* / Z-A / 9-0).

---

## Komparasi Metode Pengolahan Data

Kombinasi `FILTER` dan `SORTN` memberikan keunggulan spesifik dibandingkan metode alternatif:

| Parameter Evaluasi | Sort & Filter GUI | Pivot Table | Formula QUERY | Kombinasi FILTER + SORTN |
| :--- | :--- | :--- | :--- | :--- |
| **Sifat Pembaruan** | Statis (manual) | Semi-otomatis (butuh refresh) | Dinamis real-time | Dinamis real-time |
| **Integritas Data Asli** | Rawan teracak | Aman | Terisolasi | Terisolasi |
| **Penanganan Nilai Seri** | Manual | Terbatas | Sulit diatur otomatis | Otomatis via `display_ties_mode` |
| **Konsistensi Tipe Data** | Mengikuti sel | Fleksibel | Merusak data bertipe campuran | Mempertahankan tipe data asli |
| **Beban Komputasi** | Nol | Ringan | Sedang - Berat | Ringan (native array execution) |

Fungsi `QUERY` rentan mengalami galat ketika satu kolom berisi tipe data campuran (numerik dan teks), karena mesin Google Visualization API hanya mengadopsi tipe data mayoritas dan mengonversi tipe minoritas menjadi *null*. `FILTER` dan `SORTN` memproses data mentah tanpa transformasi skema, sehingga integritas data bertipe campuran tetap terjaga.

---

## Studi Kasus 1: Top 5 Transaksi Penjualan Terbesar Regional

### Deskripsi Masalah
Departemen Audit Operasional memerlukan data 5 transaksi terbesar khusus wilayah **"Jawa Barat"** untuk kategori **"Elektronik"**. Data baru terus bertambah pada tabel lembar kerja `Transaksi_Q3`.

### Dataset Sumber (`Transaksi_Q3`, Rentang `A1:E16`)

| Baris | A (ID_Trx) | B (Sales_Rep) | C (Wilayah) | D (Kategori) | E (Nominal) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **ID_Trx** | **Sales_Rep** | **Wilayah** | **Kategori** | **Nominal** |
| **2** | TRX-101 | Andi | Jawa Barat | Elektronik | 15500000 |
| **3** | TRX-102 | Budi | Jawa Tengah | Elektronik | 22000000 |
| **4** | TRX-103 | Citra | Jawa Barat | Fashion | 4500000 |
| **5** | TRX-104 | Dewi | Jawa Barat | Elektronik | 31000000 |
| **6** | TRX-105 | Eko | Jawa Timur | Elektronik | 18000000 |
| **7** | TRX-106 | Farhan | Jawa Barat | Elektronik | 8500000 |
| **8** | TRX-107 | Gita | Jawa Barat | Makanan | 2100000 |
| **9** | TRX-108 | Hadi | Jawa Barat | Elektronik | 42000000 |
| **10** | TRX-109 | Intan | DKI Jakarta | Elektronik | 55000000 |
| **11** | TRX-110 | Joko | Jawa Barat | Elektronik | 12000000 |
| **12** | TRX-111 | Kevin | Jawa Barat | Elektronik | 27500000 |
| **13** | TRX-112 | Lina | Jawa Barat | Fashion | 6000000 |
| **14** | TRX-113 | Mega | Jawa Barat | Elektronik | 19500000 |
| **15** | TRX-114 | Nanda | Jawa Barat | Elektronik | 9000000 |
| **16** | TRX-115 | Oscar | Jawa Tengah | Elektronik | 33000000 |

### Implementasi Formula

Tuliskan formula berikut pada sel target lembar kerja pelaporan:

```excel
=SORTN(FILTER(A2:E16; C2:C16 = "Jawa Barat"; D2:D16 = "Elektronik"); 5; 0; 5; FALSE)
```

```
+---------------------------------------------------------------------------------------+
|                                ALUR EKSEKUSI FORMULA                                 |
+---------------------------------------------------------------------------------------+
|  1. EVALUASI LOGIKA FILTER                                                            |
|     C2:C16 = "Jawa Barat"  --> [T, F, T, T, F, T, T, T, F, T, T, T, T, T, F]          |
|     D2:D16 = "Elektronik"  --> [T, T, F, T, T, T, F, T, T, T, T, F, T, T, T]          |
|     Interseksi Boolean AND --> Baris 2, 5, 7, 9, 11, 12, 14, 15 lolos validasi        |
+---------------------------------------------------------------------------------------+
                                           |
                                           v
+---------------------------------------------------------------------------------------+
|  2. OUTPUT MATRIKS FILTER (8 Baris x 5 Kolom)                                         |
|     {TRX-101; TRX-104; TRX-106; TRX-108; TRX-110; TRX-111; TRX-113; TRX-114}         |
+---------------------------------------------------------------------------------------+
                                           |
                                           v
+---------------------------------------------------------------------------------------+
|  3. PENGURUTAN SORTN                                                                  |
|     Kolom acuan: 5 (Nominal) | Arah: FALSE (Descending)                               |
|     Peringkat nilai: 42jt, 31jt, 27.5jt, 19.5jt, 15.5jt, 12jt, 9jt, 8.5jt             |
+---------------------------------------------------------------------------------------+
                                           |
                                           v
+---------------------------------------------------------------------------------------+
|  4. TRUNKASI REKORD                                                                   |
|     Argumen n = 5, ties_mode = 0                                                      |
|     Mengambil tepat 5 baris data teratas                                              |
+---------------------------------------------------------------------------------------+
```

### Hasil Keluaran Eksekusi

| ID_Trx | Sales_Rep | Wilayah | Kategori | Nominal |
| :--- | :--- | :--- | :--- | :--- |
| TRX-108 | Hadi | Jawa Barat | Elektronik | 42000000 |
| TRX-104 | Dewi | Jawa Barat | Elektronik | 31000000 |
| TRX-111 | Kevin | Jawa Barat | Elektronik | 27500000 |
| TRX-113 | Mega | Jawa Barat | Elektronik | 19500000 |
| TRX-102 | Andi | Jawa Barat | Elektronik | 15500000 |

Data langsung menempati rentang 5 baris secara otomatis melalui mekanisme *array spill*.

---

## Studi Kasus 2: Multi-Kriteria Lanjutan dengan Penanganan Nilai Seri (*Ties*)

### Deskripsi Masalah
Manajemen keuangan mewajibkan pelaporan Top 3 transaksi B2B dengan kriteria:
1. Status pemesanan bernilai **"Completed"**.
2. Nilai transaksi minimum **Rp 50.000.000**.
3. Jika terdapat transaksi pada peringkat ke-3 dan seterusnya yang memiliki nilai transaksi sama persis (*draw/ties*), baris tersebut wajib ditampilkan seluruhnya guna transparansi audit finansial.
4. Parameter filter dikendalikan dinamis melalui sel input:
   *   Sel `H1`: Parameter Status (`Completed`)
   *   Sel `H2`: Parameter Ambang Nominal Minimal (`50000000`)

### Dataset Transaksi Korporat (`Data_B2B`, Rentang `A1:F13`)

| Baris | A (Klien_ID) | B (Entitas) | C (Tgl_Order) | D (Nominal) | E (Termin_Hari) | F (Status) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **Klien_ID** | **Entitas** | **Tgl_Order** | **Nominal** | **Termin_Hari** | **Status** |
| **2** | CL-901 | PT Alpha | 2026-10-01 | 85000000 | 30 | Completed |
| **3** | CL-902 | CV Beta | 2026-10-03 | 45000000 | 14 | Completed |
| **4** | CL-903 | PT Gamma | 2026-10-05 | 60000000 | 45 | Pending |
| **5** | CL-904 | PT Delta | 2026-10-08 | 75000000 | 30 | Completed |
| **6** | CL-905 | UD Epsilon | 2026-10-10 | 50000000 | 14 | Completed |
| **7** | CL-906 | PT Zeta | 2026-10-12 | 92000000 | 60 | Cancelled |
| **8** | CL-907 | PT Eta | 2026-10-15 | 60000000 | 30 | Completed |
| **9** | CL-908 | CV Theta | 2026-10-18 | 60000000 | 30 | Completed |
| **10** | CL-909 | PT Iota | 2026-10-20 | 110000000 | 60 | Completed |
| **11** | CL-910 | PT Kappa | 2026-10-22 | 50000000 | 30 | Completed |
| **12** | CL-911 | UD Lambda | 2026-10-25 | 75000000 | 14 | Pending |
| **13** | CL-912 | PT Mu | 2026-10-28 | 40000000 | 30 | Completed |

### Implementasi Formula

Formula ditulis dengan proteksi kondisi data kosong menggunakan fungsi `IFNA`:

```excel
=IFNA(SORTN(FILTER(A2:F13; (F2:F13 = H1) * (D2:D13 >= H2)); 3; 1; 4; FALSE); "Tidak Ada Rekord yang Cocok")
```

### Analisis Logika Komputasi

1. **Perkalian Vektor Boolean `(F2:F13 = H1) * (D2:D13 >= H2)`**:
   * Evaluasi Status: Mengecek kesesuaian nilai kolom `F` dengan sel `H1` (`"Completed"`).
   * Evaluasi Batas: Mengecek apakah nilai kolom `D` lebih besar atau sama dengan sel `H2` (`50000000`).
   * Operasi perkalian skalar `*` bertindak sebagai gerbang logika `AND`. Operasi ini menghasilkan representasi integer `1` (keduanya valid) atau `0` (salah satu/keduanya tidak valid).

2. **Daftar Baris yang Lolos Filter**:
   * Baris 2: PT Alpha (Nominal 85.000.000) -> Valid
   * Baris 5: PT Delta (Nominal 75.000.000) -> Valid
   * Baris 6: UD Epsilon (Nominal 50.000.000) -> Valid
   * Baris 8: PT Eta (Nominal 60.000.000) -> Valid
   * Baris 9: CV Theta (Nominal 60.000.000) -> Valid
   * Baris 10: PT Iota (Nominal 110.000.000) -> Valid
   * Baris 11: PT Kappa (Nominal 50.000.000) -> Valid

3. **Perilaku Argumen `display_ties_mode = 1` pada `SORTN`**:
   * Kolom pengurutan diatur pada indeks ke-`4` (Nominal), urutan menurun (`FALSE`).
   * Nilai urutan teratas:
     * Peringkat 1: PT Iota (110.000.000)
     * Peringkat 2: PT Alpha (85.000.000)
     * Peringkat 3: PT Delta (75.000.000)
   * Jika parameter `n` disetel `3` dan ambang diubah ke nominal yang memicu nilai seri:
     
Uji skenario jika mencari Top 3 dengan batas Rp 60.000.000 ke bawah, atau ketika terjadi benturan peringkat. Perhatikan kasus data jika kuota dibatasi hingga peringkat nominal 60.000.000:

Misalkan formula disesuaikan untuk mengambil Top 4 (`n=4`):
* Peringkat 1: 110.000.000 (PT Iota)
* Peringkat 2: 85.000.000 (PT Alpha)
* Peringkat 3: 75.000.000 (PT Delta)
* Peringkat 4: 60.000.000 (PT Eta)
* Peringkat 4 (Seri): 60.000.000 (CV Theta)

Karena mode diatur ke `1`, sistem mendeteksi CV Theta bernilai sama dengan batas peringkat ke-4 (PT Eta). Sistem secara otomatis mengekspansi *array output* menjadi 5 baris, mencegah eliminasi data sepihak.

### Hasil Keluaran Eksekusi (Parameter: n = 4, ties_mode = 1)

| Klien_ID | Entitas | Tgl_Order | Nominal | Termin_Hari | Status |
| :--- | :--- | :--- | :--- | :--- | :--- |
| CL-909 | PT Iota | 2026-10-20 | 110000000 | 60 | Completed |
| CL-901 | PT Alpha | 2026-10-01 | 85000000 | 30 | Completed |
| CL-904 | PT Delta | 2026-10-08 | 75000000 | 30 | Completed |
| CL-907 | PT Eta | 2026-10-15 | 60000000 | 30 | Completed |
| CL-908 | CV Theta | 2026-10-18 | 60000000 | 30 | Completed |

---

## Mitigasi Kesalahan Teknis (Pitfalls)

### 1. Penanganan Galat `#N/A` (No Matches Found)
*   **Penyebab**: Fungsi `FILTER` tidak menemukan satu pun baris data yang memenuhi kriteria pengujian boolean.
*   **Dampak**: Menampilkan galat visual dan merusak kalkulasi hilir (*downstream dependencies*).
*   **Solusi**: Bungkus formula menggunakan fungsi `IFNA`:
    ```excel
    =IFNA(SORTN(FILTER(A2:E; C2:C = "Kriteria_Kosong"); 5; 0; 5; FALSE); "Data Tidak Ditemukan")
    ```

### 2. Penanganan Galat `#VALUE!` (Array Dimensions Mismatch)
*   **Penyebab**: Ketidaksesuaian jumlah baris antara rentang data utama dengan rentang kondisi pengujian logika.
*   **Contoh Salah**:
    ```excel
    =FILTER(A2:E100; C2:C50 = "Jawa Barat")
    ```
*   **Solusi**: Pastikan batas awal dan akhir baris bernilai identik pada setiap rentang:
    ```excel
    =FILTER(A2:E100; C2:C100 = "Jawa Barat")
    ```

### 3. Penanganan Galat `#REF!` (Spill Range Obstruction)
*   **Penyebab**: *Dynamic Array* membutuhkan sejumlah sel kosong di bawah dan di kanan sel jangkar formula. Galat terjadi jika terdapat teks, formula lain, atau spasi kosong (*whitespace*) yang menghalangi rentang ekspansi.
*   **Solusi**: Kosongkan seluruh sel di sekitar area proyeksi tumpahan formula (*spill area*).

### 4. Perbedaan Delimiter Regional (Koma vs Titik Koma)
*   **Penyebab**: Konfigurasi lokal (*File > Settings > Locale*) memengaruhi karakter pemisah argumen formula.
*   **Sistem Metrik / Indonesia**: Menggunakan titik koma (`;`) sebagai pemisah argumen fungsi karena koma (`,`) dialokasikan sebagai pemisah desimal.
*   **Sistem Anglo-Saxon / Amerika Serikat**: Menggunakan koma (`,`) sebagai pemisah argumen.
*   **Contoh Komparasi**:
    *   *Locale Indonesia*: `=SORTN(FILTER(A2:E; C2:C="X"); 5; 0; 5; FALSE)`
    *   *Locale US*: `=SORTN(FILTER(A2:E, C2:C="X"), 5, 0, 5, FALSE)`

### 5. Kesalahan Referensi Kolom Urut pada SORTN
*   **Penyebab**: Menuliskan notasi rentang absolut seperti `E2:E` atau `E` pada parameter `sort_kolom` di dalam fungsi `SORTN`.
*   **Penjelasan Teknis**: Input `SORTN` adalah hasil matriks dari `FILTER`, bukan referensi lembar kerja langsung.
*   **Koreksi**: Gunakan bilangan bulat murni (*integer*) yang merepresentasikan indeks kolom relatif terhadap matriks input (kolom ke-1, ke-2, dst.).

---

## Checklist Praktik Terbaik

*   **Pemisahan Tabel Secara Fisik**: Selalu tempatkan tabel sumber data mentah (*raw data*) pada tab terpisah dari tabel laporan atau ringkasan eksekutif (*dashboard*).
*   **Penggunaan Rentang Terbuka (*Open-Ended Range*)**: Gunakan notasi rentang terbuka (contoh: `A2:E`) untuk mengakomodasi penambahan data transaksi baru tanpa konfigurasi ulang formula.
*   **Pencegahan Penurunan Performa (*Calculation Overhead*)**: Batasi rentang terbuka jika baris data mencapai ratusan ribu baris. Formula dinamis yang membaca baris kosong di luar batas data riil tetap memakan alokasi memori kalkulasi.
*   **Dokumentasikan Mode Nilai Seri**: Selalu tentukan nilai eksplisit pada argumen `display_ties_mode` (jangan biarkan kosong) untuk menghindari ambiguitas penafsiran data saat pelaporan resmi.

---

## Ringkasan Eksekutif

Kombinasi `FILTER` dan `SORTN` menyediakan arsitektur pelaporan data otomatis berkinerja tinggi pada Google Sheets. Penggabungan kedua fungsi ini menyelesaikan masalah pemrosesan data multi-kriteria sekaligus mengeliminasi tahapan sortir manual.

Dengan membedah matriks data secara langsung di memori lokal spreadsheet:
1. `FILTER` bertindak sebagai gerbang isolasi logika untuk menyaring data yang relevan tanpa merusak struktur data mentah.
2. `SORTN` mengontrol urutan dan volume data keluaran serta memberikan kepastian tata kelola terhadap data bernilai seri (*ties*).

Penerapan metode ini memastikan laporan manajerial, analitik penjualan, dan audit internal berjalan konsisten, presisi, serta terbarui otomatis secara real-time.