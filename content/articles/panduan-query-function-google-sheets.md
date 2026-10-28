---
title: "Panduan Lengkap Fungsi QUERY di Google Sheets: Olah Data Layaknya SQL"
description: "Gunakan perintah SELECT, WHERE, GROUP BY, PIVOT, dan ORDER BY langsung di dalam Google Sheets menggunakan fungsi sakti QUERY."
category: "Google Sheets"
date: "2026-10-28"
readTime: "8 menit baca"
---

# Panduan Lengkap Fungsi QUERY di Google Sheets: Olah Data Layaknya SQL

Pemrosesan dataset berskala ribuan baris pada Google Sheets sering terkendala penurunan performa komputasi. Pendekatan konvensional yang menggabungkan fungsi `FILTER`, `SORT`, `UNIQUE`, dan `SUMIFS` secara bertingkat (nested formula) memaksa spreadsheet engine mengevaluasi array kalkulasi berkali-kali secara simultan. Dampaknya: konsumsi memori browser membengkak, latensi hitung ulang meningkat, dan struktur logika rentan rusak saat skema kolom bergeser.

Fungsi `QUERY` menyelesaikan limitasi tersebut. Berjalan di atas **Google Visualization API Query Language**, `QUERY` mengeksekusi transformasi, manipulasi, kalkulasi agregasi, dan restrukturisasi data multidimensi dalam satu baris instruksi deklaratif dengan sintaksis menyerupai SQL (*Structured Query Language*).

---

## Anatomi Sintaksis dan Argumen Fungsi

Fungsi `QUERY` menerima tiga parameter:

```excel
=QUERY(data; query_string; [headers])
```

### Parameter:
1. `data`: Rentang sel sumber (*range* array). Format: rentang sel langsung (`A1:E500`) atau bentukan array vertikal/horizontal (`{Sheet1!A1:E; Sheet2!A2:E}`).
2. `query_string`: Kumpulan klausa deklaratif berbasis teks yang dibungkus tanda petik ganda (`"..."`). Menginstruksikan transformasi data yang akan dieksekusi engine API.
3. `[headers]` *(Opsional)*: Integritas jumlah baris header di bagian atas `data`. Jika tidak dideklarasikan, engine menebak secara otomatis berdasarkan konsistensi tipe data sel teratas. Praktik terbaik: eksplisitkan nilai `1` jika data memiliki satu baris header, atau `0` jika data mentah tanpa header.

### Klausul Utama Query Language (Urutan Evaluasi Wajib)

Query engine memproses instruksi berdasarkan urutan eksekusi tetap (*strict evaluation order*). Pelanggaran urutan memicu galat `#VALUE!`.

| Klausul | Urutan | Fungsi Teknis |
| :--- | :--- | :--- |
| `SELECT` | 1 | Menentukan kolom yang diekstrak, dimanipulasi, atau diagregasi. |
| `WHERE` | 2 | Memfilter baris berdasarkan kondisi logika/perbandingan nilai. |
| `GROUP BY` | 3 | Mengelompokkan record data berdasarkan nilai identik untuk agregasi. |
| `PIVOT` | 4 | Mentransformasikan record unik kolom menjadi layout matriks horizontal. |
| `ORDER BY` | 5 | Mengurutkan baris output secara *ascending* (`ASC`) atau *descending* (`DESC`). |
| `LIMIT` | 6 | Membatasi kuantitas record baris yang dirender pada output. |
| `OFFSET` | 7 | Melompati baris data teratas sebanyak nilai integer tertentu. |
| `LABEL` | 8 | Merombak atau mengosongkan nama header kolom pada output. |
| `FORMAT` | 9 | Mengatur format numerik, tanggal, atau mata uang secara langsung di dalam instruksi query. |

---

## Analisis Komparasi: QUERY vs Formula Konvensional

| Parameter Evaluasi | Metode Konvensional (`FILTER` / `SUMIFS` / `SORT`) | Native Pivot Table | Fungsi `QUERY` |
| :--- | :--- | :--- | :--- |
| **Beban Memory** | Tinggi (evaluasi formula terpisah per baris/kolom) | Sedang | Rendah (pemrosesan terpusat via API engine) |
| **Dinamisme Output** | Memerlukan duplikasi formula ke baris baru | Memerlukan trigger *Refresh Data* manual | Menghasilkan array otomatis (*spill range*) |
| **Fleksibilitas Logika** | Sintaks rumit jika melibatkan filter multikolom | Terbatas pada konfigurasi UI | Sangat fleksibel via klausa `AND`, `OR`, `MATCHES` |
| **Manipulasi Header** | Membutuhkan baris manual tambahan | Terikat layout pivot | Langsung disematkan via klausa `LABEL` |
| **Resiko Referensi** | Rawan desinkronisasi sel ($ lock error) | Terisolasi | Stabil jika didefinisikan dengan notasi tepat |

---

## Studi Kasus 1: Konsolidasi dan Agregasi Transaksi Retail

### Struktur Sumber Data (Dataset Transaksi: Sel `A1:E1000`)
- **Kolom A**: `ID_Transaksi` (String)
- **Kolom B**: `Tanggal` (Date: YYYY-MM-DD)
- **Kolom C**: `Wilayah` (String: "Jawa", "Sumatera", "Kalimantan", "Bali")
- **Kolom D**: `Kategori` (String: "Elektronik", "Fashion", "FMCG")
- **Kolom E**: `Total_Omset` (Numeric: Nilai Transaksi Rupiah)

### Sasaran Masalah
Departemen Keuangan membutuhkan rekapitulasi data:
1. Hanya kategori `"Elektronik"`.
2. Total omset per transaksi minimal `Rp 1.000.000`.
3. Tampilkan total akumulasi omset dan frekuensi transaksi per wilayah.
4. Urutkan wilayah dengan omset tertinggi ke terendah.
5. Ganti nama header agregasi agar representatif.

### Formula Solusi

```excel
=QUERY(A1:E1000; "SELECT C, COUNT(A), SUM(E) WHERE D = 'Elektronik' AND E >= 1000000 GROUP BY C ORDER BY SUM(E) DESC LABEL C 'Wilayah Operasional', COUNT(A) 'Volume Transaksi', SUM(E) 'Total Omset Bersih' FORMAT SUM(E) 'Rp#,##0'"; 1)
```

### Bedah Logika Formula
- `SELECT C, COUNT(A), SUM(E)`: Mengambil dimensi nama wilayah (Kolom C), menghitung jumlah baris unik ID (Kolom A), dan mengakumulasi total omset (Kolom E).
- `WHERE D = 'Elektronik' AND E >= 1000000`: Menerapkan filter deterministik. String wajib diapit tanda petik tunggal (`'...'`). Operator logika `AND` memastikan kedua parameter terpenuhi sebelum baris dialirkan ke tahap kalkulasi.
- `GROUP BY C`: Wajib disertakan karena kolom non-agregat (`C`) dipadukan dengan fungsi agregat (`COUNT` dan `SUM`). Data dikelompokkan berdasarkan keunikan teks di Kolom C.
- `ORDER BY SUM(E) DESC`: Mengurutkan tabel output akhir berdasarkan total akumulasi uang terbesar ke terkecil.
- `LABEL ...`: Menghapus header default API (`sum Total_Omset`) dan menggantinya dengan teks terdefinisi.
- `FORMAT SUM(E) 'Rp#,##0'`: Memaksa engine me-render angka numerik langsung dengan format mata uang tanpa mengubah format sel dasar pada sheet.
- Parameter `1`: Mengabaikan baris 1 sumber sebagai data murni, memperlakukannya sebagai representasi skema kolom.

---

## Studi Kasus 2: Matrix Cross-Tabulation Dinamis Berdasarkan Input Sel

Skenario: Membuat laporan ringkasan matriks cross-tabulation (dua dimensi) yang membagi transaksi per wilayah (vertikal) dan kategori produk (horizontal), dengan rentang tanggal fleksibel yang nilainya diisi pengguna di sel interface.

### Parameter Input Interaktif
- Sel `G1`: Tanggal Awal (Contoh: `2026-01-01`)
- Sel `G2`: Tanggal Akhir (Contoh: `2026-06-30`)

### Karakteristik Format Query String
Google Visualization API Query Language tidak dapat membaca langsung referensi sel spreadsheet di dalam tanda petik ganda. String harus diputus dan disambungkan menggunakan operator konkatenasi teks (`&`). Tipe data tanggal harus dikonversi ke format standar ISO teks `date 'YYYY-MM-DD'`.

### Penggunaan Array Literal `{...}` dan Notasi Kolom
Jika rentang sumber dibungkus dalam tanda kurung kurawal `{A1:E1000}`, referensi kolom berubah dari huruf (`A`, `B`, `C`) menjadi notasi indeks (`Col1`, `Col2`, `Col3`). Pola ini membuat formula kebal terhadap pergeseran posisi sheet dan kompatibel saat menggabungkan multi-tabel via `IMPORTRANGE`.

### Formula Solusi

```excel
=QUERY({A1:E1000}; "SELECT Col3, SUM(Col5) WHERE Col2 >= date '"&TEXT(G1; "yyyy-mm-dd")&"' AND Col2 <= date '"&TEXT(G2; "yyyy-mm-dd")&"' GROUP BY Col3 PIVOT Col4 LABEL Col3 'Wilayah / Kategori' FORMAT SUM(Col5) '#,##0'"; 1)
```

### Bedah Logika Lanjutan
- `{A1:E1000}`: Membuka engine array in-memory. Kolom A menjadi `Col1`, Kolom B menjadi `Col2`, dst.
- `date '"&TEXT(G1; "yyyy-mm-dd")&"'`: Mengurai objek tanggal serial Google Sheets dari sel `G1` menjadi representasi literal yang valid di API: `date '2026-01-01'`.
- `PIVOT Col4`: Mentransformasikan nilai unik di `Col4` (`Kategori`) menjadi header kolom horizontal baru. Secara otomatis menghasilkan cross-tabulation antara Kolom C (`Wilayah`) pada sumbu baris (Y) dan Kolom D (`Kategori`) pada sumbu kolom (X).
- Nilai di perpotongan baris dan kolom secara otomatis diagregasikan menggunakan fungsi `SUM(Col5)`.

---

## Pitfall, Kesalahan Fatal, dan Metode Debugging

### 1. Masalah Tipe Data Campuran (Mixed Data Types Silent Error)
**Gejala:** Nilai teks atau angka tertentu hilang menjadi sel kosong (*blank*) pada hasil `QUERY`, tanpa menampilkan pesan error.

**Penyebab Sistemik:** Google Visualization API Query Language menerapkan aturan *data type homogeneity*. Engine memeriksa mayoritas tipe data dalam 100 baris pertama suatu kolom. Jika 80% baris bertipe numerik dan 20% bertipe teks (misal: "N/A", "Pending"), engine menetapkan kolom tersebut sebagai *numeric*. Semua nilai teks otomatis diubah menjadi `NULL`.

**Solusi:**
Konversi kolom sumber menjadi teks seutuhnya sebelum diproses oleh `QUERY` menggunakan fungsi `TO_TEXT` atau menggabungkannya dengan string kosong via `ARRAYFORMULA`:

```excel
=QUERY(INDEX(TO_TEXT(A1:E1000)); "SELECT Col1, Col3 WHERE Col5 != ''"; 1)
```

### 2. Inkonsistensi Regional Separator (Koma `,` vs Titik Koma `;`)
Konfigurasi locale Google Sheets menentukan parsing sintaks:
- **Locale US / Internasional:** Menggunakan koma `,` sebagai pemisah argumen fungsi dan titik `.` untuk desimal.
  ```excel
  =QUERY(A1:E1000, "SELECT A, B", 1)
  ```
- **Locale Indonesia:** Menggunakan titik koma `;` sebagai pemisah argumen fungsi dan koma `,` untuk desimal.
  ```excel
  =QUERY(A1:E1000; "SELECT A, B"; 1)
  ```
*Catatan:* Di dalam tanda petik query string (`"SELECT A, B"`), pemisah antar kolom yang dipilih **selalu koma `,`**, tidak terpengaruh pengaturan locale spreadsheet.

### 3. Syntax Error Literal Tanggal
Menuliskan kondisi filter tanggal seperti `WHERE B >= '2026-01-01'` menghasilkan error parsing `#VALUE!`. Teks tersebut dibaca sebagai string biasa, bukan objek waktu.

Gunakan kata kunci `date`:
```text
WHERE B >= date '2026-01-01'
```
Untuk data berbasis timestamp jam/menit/detik, gunakan kata kunci `datetime`:
```text
WHERE B >= datetime '2026-01-01 14:30:00'
```

### 4. Pelanggaran Spill Range (`#REF!`)
Error `#REF!` bertuliskan *"Array result was not expanded because it would overwrite data in..."* terjadi jika sel di area pemekaran hasil `QUERY` terhalang oleh nilai atau teks lain.

**Solusi:** Kosongkan seluruh sel di bawah dan di kanan sel lokasi formula `QUERY` ditulis.

### 5. Masalah Klausa String Case-Sensitivity
Pencocokan teks dalam klausa `WHERE` bersifat *case-sensitive*. String `'Elektronik'` tidak akan menangkap record berstatus `'elektronik'` atau `'ELEKTRONIK'`.

Gunakan fungsi konversi teks bawaan API (`lower` atau `upper`):
```text
WHERE lower(D) = 'elektronik'
```
Atau gunakan operator regular expression `MATCHES`:
```text
WHERE D MATCHES '(?i)elektronik'
```

---

## Integrasi Lanjutan: Automasi Validasi Query via Google Apps Script

Bila struktur spreadsheet sangat kompleks atau diperbarui secara terjadwal oleh sistem backend eksternal, validasi integritas struktur formula `QUERY` dapat diamankan menggunakan Google Apps Script berikut:

```javascript
/**
 * Memvalidasi dan menyuntikkan formula QUERY secara otomatis
 * guna menghindari modifikasi manual yang merusak sintaks.
 */
function injectAutomatedQuery() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Data_Transaksi");
  const reportSheet = ss.getSheetByName("Laporan_Agregasi");
  
  if (!sourceSheet || !reportSheet) {
    throw new Error("Sheet sumber atau tujuan tidak ditemukan.");
  }

  // Tentukan formula terpusat
  const targetFormula = "=QUERY(Data_Transaksi!A1:E1000; \"SELECT C, SUM(E) WHERE D = 'Elektronik' GROUP BY C ORDER BY SUM(E) DESC LABEL C 'Wilayah', SUM(E) 'Total Transaksi'\"; 1)";
  
  // Letakkan formula pada target sel A1 laporan
  const targetCell = reportSheet.getRange("A1");
  targetCell.setFormula(targetFormula);
  
  Logger.log("Formula QUERY berhasil disinkronisasi pada: " + new Date().toISOString());
}
```

---

## Checklist Praktik Terbaik (Best Practices)

1. **Definisikan Rentang Kolom Secara Eksplisit:** Hindari penggunaan rentang terbuka tak terbatas (`A:E`) pada dataset besar karena memaksa spreadsheet memuat jutaan baris kosong ke memori. Batasi batas baris (`A1:E10000`).
2. **Kunci Argumen Header:** Selalu deklarasikan nilai integer parameter ketiga (`0` atau `1`). Jangan biarkan kosong agar engine tidak salah menerjemahkan baris data transaksi pertama sebagai label header.
3. **Gunakan Notasi Col1, Col2 jika Merujuk Array:** Wajib menggunakan format `ColX` saat data sumber berada dalam kurung kurawal `{}` atau berasal dari formula array lain seperti `IMPORTRANGE`. Perhatikan sensitivitas huruf kapital: `Col1` valid, `col1` menghasilkan error.
4. **Isolasi Tipe Data per Kolom:** Pastikan kolom numerik tidak mengandung catatan teks kaki (*footnote*) agar tidak memicu *nulling data* otomatis oleh Visualization API.
5. **Gunakan Klausa LIMIT Saat Debugging:** Uji kueri rumit pada dataset masif dengan menambahkan `LIMIT 10` di ujung instruksi untuk memvalidasi output sebelum memproses keseluruhan data.

---

## Ringkasan Eksekutif

Fungsi `QUERY` mengonsolidasikan beragam fungsi lookup, filter, sortir, dan matriks ke dalam satu mesin komputasi terstruktur. Menguasai urutan evaluasi klausul (`SELECT` hingga `FORMAT`), penanganan tipe data yang homogen, dan integrasi referensi dinamis menjamin pipeline pemrosesan data spreadsheet beroperasi cepat, terstruktur, minim formula bertingkat, dan tahan terhadap volume data besar.