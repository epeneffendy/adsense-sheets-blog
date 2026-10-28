---
title: "Membuat Mini Dashboard Visual di Sel dengan SPARKLINE Google Sheets"
description: "Desain tren grafik garis (line chart), grafik batang (bar), dan win-loss visual langsung di dalam satu sel menggunakan fungsi SPARKLINE."
category: "Google Sheets"
date: "2026-10-28"
readTime: "8 menit baca"
---

## Urgensi Visualisasi In-Cell pada Spreadsheet Skala Besar

Lembar kerja modern sering mengalami degradasi performa akibat kelebihan objek visual mengambang (*floating chart objects*). Saat dokumen spreadsheet memuat puluhan grafik standar, beban Document Object Model (DOM) browser melonjak drastis, memicu *input lag*, konsumsi RAM berlebih, dan keterlambatan rendering data dinamis. Masalah bertambah saat lembar kerja diekspor ke format PDF atau dicetak: tata letak floating chart bergeser, menutupi sel data, dan memutus keterbacaan laporan eksekutif.

Visualisasi di dalam sel (*in-cell visualization*) menyelesaikan kendala ini. Fungsi `SPARKLINE` di Google Sheets memproses vektor data langsung ke dalam raster engine sel. Formula ini tidak menciptakan objek mengambang independen, melainkan mengubah nilai numerik menjadi grafik mini terintegrasi langsung di grid sel.

Kebutuhan otomasi pelaporan finansial, logistik, dan metrik operasional menuntut visualisasi densitas tinggi (*high data-density displays*). Menggunakan `SPARKLINE`, puluhan metrik performa (KPI) dapat divisualisasikan dalam satu layar tanpa memerlukan scrollbar vertikal maupun horizontal.

---

## Anatomi Sintaks Fungsi & Spesifikasi Parameter Teknis

Sintaks dasar fungsi `SPARKLINE` membutuhkan satu argumen data wajib dan satu argumen opsional untuk konfigurasi atribut visual:

```excel
=SPARKLINE(data; [opsi])
```

- `data`: Rentang (*range*) atau array nilai numerik yang menjadi dasar kalkulasi visual (contoh: `A2:A13` atau `{10, 25, 40, -15, 30}`).
- `[opsi]`: Pasangan atribut array dua dimensi di dalam kurung kurawal `{}` yang menentukan jenis grafik, batas skala, dan palet warna.

### Struktur Array Opsi (Key-Value Pair)

Penulisan opsi `SPARKLINE` menggunakan format key-value. Penanda pemisah baris dan kolom bergantung pada konfigurasi lokal (Locale) spreadsheet:

- **Locale Indonesia / Eropa Kontinental**: Menggunakan titik koma (`;`) untuk memisahkan argumen fungsi, koma (`,`) untuk memisahkan pasangan kunci-nilai, dan titik koma (`;`) atau backslash (`\`) untuk baris opsi berikutnya:
  ```excel
  {"charttype", "line"; "color", "red"; "linewidth", 2}
  ```
- **Locale Amerika Serikat (US)**: Menggunakan koma (`,`) untuk memisahkan argumen fungsi, koma (`,`) untuk kunci-nilai, dan titik koma (`;`) untuk pemisah atribut:
  ```excel
  {"charttype", "line"; "color", "red"; "linewidth", 2}
  ```

### Tipe Grafik dan Parameter Spesifik

Fungsi `SPARKLINE` mendukung 4 tipe grafik: `"line"`, `"bar"`, `"column"`, dan `"winloss"`.

#### 1. Grafik Garis (`"line"`)
Jenis default jika `charttype` tidak didefinisikan secara eksplisit.

| Parameter | Tipe Data | Deskripsi Teknis |
| :--- | :--- | :--- |
| `"charttype"` | String | Nilai: `"line"`. |
| `"color"` | Hex / String | Menentukan warna garis (contoh: `"#1A73E8"` atau `"navy"`). |
| `"linewidth"` | Numerik | Ketebalan garis dalam satuan piksel (default: `1`). |
| `"ymin"` | Numerik | Batas minimum sumbu Y. Menentukan titik terbawah kanvas sel. |
| `"ymax"` | Numerik | Batas maksimum sumbu Y. Menghindari distorsi skala otomatis. |
| `"empty"` | String | Penanganan sel kosong: `"ignore"` (lewatkan) atau `"zero"` (anggap 0). |
| `"nan"` | String | Penanganan nilai non-numerik: `"convert"` (ubah ke 0) atau `"ignore"`. |
| `"rtl"` | Boolean | Menampilkan grafik dari kanan ke kiri (`TRUE` / `FALSE`). |

#### 2. Grafik Kolom Vertikal (`"column"`)
Menampilkan setiap titik data sebagai batang vertikal berurutan dari kiri ke kanan.

| Parameter | Tipe Data | Deskripsi Teknis |
| :--- | :--- | :--- |
| `"charttype"` | String | Nilai: `"column"`. |
| `"color"` | Hex / String | Warna default seluruh kolom data. |
| `"lowcolor"` | Hex / String | Warna khusus untuk titik nilai numerik terendah. |
| `"highcolor"`| Hex / String | Warna khusus untuk titik nilai numerik tertinggi. |
| `"firstcolor"`| Hex / String | Warna khusus untuk batang data urutan pertama. |
| `"lastcolor"` | Hex / String | Warna khusus untuk batang data urutan terakhir. |
| `"negcolor"` | Hex / String | Warna khusus untuk data bernilai negatif ($< 0$). |
| `"axis"` | Boolean | Menampilkan garis axis sumbu horizontal (`TRUE` / `FALSE`). |
| `"axiscolor"`| Hex / String | Menentukan warna garis sumbu horizontal jika `axis` bernilai `TRUE`. |
| `"ymin"`, `"ymax"` | Numerik | Menentukan limit skala absolut bawah dan atas. |

#### 3. Grafik Batang Horizontal Bertumpuk (`"bar"`)
Membagi lebar horizontal sel menjadi segmen-segmen proporsional. Ideal untuk *stacked progress bar*.

| Parameter | Tipe Data | Deskripsi Teknis |
| :--- | :--- | :--- |
| `"charttype"` | String | Nilai: `"bar"`. |
| `"max"` | Numerik | Nilai batas maksimum horizontal (100% lebar sel). |
| `"color1"` | Hex / String | Warna batang untuk nilai/segmen pertama. |
| `"color2"` | Hex / String | Warna batang untuk nilai/segmen kedua. |
| `"rtl"` | Boolean | Render batang horizontal dari kanan ke kiri. |

#### 4. Grafik Biner / Hasil Transaksi (`"winloss"`)
Mirip grafik kolom, tetapi menstandarkan tinggi batang: semua nilai positif dirender dengan tinggi seragam ke atas, nilai negatif dirender dengan tinggi seragam ke bawah, dan nilai nol tidak membentuk batang. Mendukung parameter pewarnaan yang identik dengan `"column"` (`color`, `negcolor`, `highcolor`, `lowcolor`, `axis`, `axiscolor`).

---

## Matriks Komparasi: SPARKLINE vs Floating Chart vs Data Bars

| Kriteria Evaluasi | In-Cell SPARKLINE | Floating Chart Standar | Conditional Formatting (Data Bars) |
| :--- | :--- | :--- | :--- |
| **Beban Memori / DOM** | Sangat Ringan (Dihitung internal) | Berat (Setiap chart adalah objek SVG terpisah) | Sangat Ringan (Native formatting CSS) |
| **Dukungan Baris Massal** | Ratusan baris tanpa lag | Drop performa parah jika $>10$ chart | Ratusan baris tanpa lag |
| **Kompensasi Auto-Resize** | Mengikuti dimensi baris/kolom | Posisi bergeser (*floating drift*) | Mengikuti dimensi baris/kolom |
| **Dukungan Tipe Tren** | Line, Column, Stacked Bar, Winloss | Semua varian grafik kompleks 2D/3D | Hanya Bar tunggal |
| **Keterbacaan Nilai Presisi** | Tidak ada tooltip angka langsung | Mendukung interaktivitas hover tooltip | Angka tampil langsung di dalam sel |
| **Skala Relatif Antar-Baris** | Harus dikunci via parameter `max`/`ymax` | Otomatis terkalibrasi jika 1 canvas | Menggunakan min/max otomatis atau kustom |
| **Kompatibilitas Formula** | Penuh (Dapat digabung fungsi matriks) | Terbatas pada sumber data range | Hanya merespons sel target tunggal |

---

## Studi Kasus 1: Monitoring Tren Metrik Finansial Bulanan

### Struktur Data Awal

Tabel kinerja finansial 12 bulan berada di rentang `A1:E13`:

| Sel | Kolom A (Bulan) | Kolom B (Pendapatan - Juta IDR) | Kolom C (Beban Operasional - Juta IDR) | Kolom D (Laba Bersih - Juta IDR) | Kolom E (Status Performa) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **A1:E1** | **Bulan** | **Pendapatan** | **Beban** | **Laba Bersih** | **Margin Status** |
| Baris 2 | Jan | 120 | 95 | 25 | Positif |
| Baris 3 | Feb | 135 | 110 | 25 | Positif |
| Baris 4 | Mar | 110 | 115 | -5 | Defisit |
| Baris 5 | Apr | 145 | 100 | 45 | Positif |
| Baris 6 | Mei | 160 | 120 | 40 | Positif |
| Baris 7 | Jun | 150 | 155 | -5 | Defisit |
| Baris 8 | Jul | 175 | 130 | 45 | Positif |
| Baris 9 | Agu | 190 | 140 | 50 | Positif |
| Baris 10 | Sep | 165 | 170 | -5 | Defisit |
| Baris 11 | Okt | 210 | 150 | 60 | Positif |
| Baris 12 | Nov | 225 | 160 | 65 | Positif |
| Baris 13 | Des | 240 | 165 | 75 | Positif |

Kita akan membuat dashboard ringkasan pada sel `G2:G5`.

### Formula 1: Visualisasi Tren Pendapatan Tahunan (Line Chart)

Tampilkan garis tren pendapatan di sel `G2` lengkap dengan penanda batas minimum dan maksimum otomatis agar grafik tidak flat:

```excel
=SPARKLINE(B2:B13; {"charttype", "line"; "color", "#1A73E8"; "linewidth", 2; "empty", "zero"; "ymin", MIN(B2:B13)*0.9; "ymax", MAX(B2:B13)*1.05})
```

#### Bedah Logika Formula:
1. `B2:B13`: Mengambil data deret waktu pendapatan Jan–Des.
2. `"charttype", "line"`: Menginstruksikan mesin rendering untuk membentuk garis spline kontinu.
3. `"color", "#1A73E8"`: Menggunakan standar warna korporat biru heksadesimal Google.
4. `"linewidth", 2`: Meningkatkan ketebalan garis dari 1 piksel ke 2 piksel agar tidak buram pada monitor resolusi tinggi.
5. `"ymin", MIN(B2:B13)*0.9`: Mengunci skala bawah pada 90% dari nilai pendapatan terkecil (110 * 0.9 = 99). Hal ini mencegah visualisasi terpotong di batas bawah sekaligus memperjelas fluktuasi data.
6. `"ymax", MAX(B2:B13)*1.05`: Menambahkan ruang atas (*padding*) 5% di atas angka tertinggi (240 * 1.05 = 252) agar kurva puncak tidak menyentuh garis kisi atas sel.

### Formula 2: Distribusi Laba Bersih Dinamis dengan Penanda Defisit (Column Chart)

Tampilkan distribusi laba bersih bulanan di sel `G3` dengan kondisi: laba positif berwarna abu-abu, laba tertinggi berwarna hijau, nilai terendah/defisit berwarna merah menyala, dan sertakan garis sumbu nol horizontal.

```excel
=SPARKLINE(D2:D13; {"charttype", "column"; "color", "#9AA0A6"; "highcolor", "#1E8E3E"; "negcolor", "#D93025"; "lowcolor", "#D93025"; "axis", TRUE; "axiscolor", "#202124"})
```

#### Bedah Logika Formula:
1. `D2:D13`: Mengambil data laba bersih yang memuat angka positif dan negatif.
2. `"charttype", "column"`: Mengubah proyeksi visual menjadi bar vertikal per bulan.
3. `"color", "#9AA0A6"`: Menetapkan warna dasar abu-abu netral untuk performa reguler.
4. `"highcolor", "#1E8E3E"`: Mengisolasi nilai maksimum (bulan Desember: 75) dan memberinya warna hijau tua.
5. `"negcolor", "#D93025"`: Setiap batang dengan nilai $<0$ (Maret, Juni, September: -5) otomatis diwarnai merah.
6. `"axis", TRUE` dan `"axiscolor", "#202124"`: Mengaktifkan garis sumbu dasar $Y=0$ berwarna abu-abu gelap, membagi kolom positif (mengarah ke atas) dan kolom negatif (mengarah ke bawah).

---

## Studi Kasus 2: Pelacakan Target Tim Sales & Win-Loss Streak Transaksi

### Skenario Progres Target Penjualan (Stacked Horizontal Bar)

Tabel berikut menunjukkan pencapaian penjualan 3 personel sales rep terhadap kuota individu masing-masing:

| Baris | Sales Rep (Kolom A) | Realisasi Capaian (Kolom B) | Target Kuota (Kolom C) | Sisa Target (Kolom D) | Progress In-Cell (Kolom E) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Row 2** | Budi Santoso | 85 | 100 | `=MAX(0; C2-B2)` | *Formula SPARKLINE* |
| **Row 3** | Siti Aminah | 120 | 100 | `=MAX(0; C3-B3)` | *Formula SPARKLINE* |
| **Row 4** | Hendra Wijaya | 45 | 100 | `=MAX(0; C4-B4)` | *Formula SPARKLINE* |

Kolom D dihitung menggunakan formula:
```excel
=MAX(0; C2-B2)
```

Untuk sel `E2` (Budi Santoso), buat grafik progres horizontal bertumpuk (*progress bar*): jika target belum tercapai, realisasi berwarna biru dan sisa target berwarna abu-abu terang. Batas horizontal mutlak adalah nilai Target (`C2`).

```excel
=SPARKLINE(B2:D2; {"charttype", "bar"; "color1", "#1A73E8"; "color2", "#E8EAED"; "max", C2})
```

#### Bedah Logika Formula:
1. `B2:D2`: Array dua sel berisi nilai realisasi (`85`) dan selisih sisa target (`15`).
2. `"charttype", "bar"`: Merender segmen batang horizontal tunggal.
3. `"color1", "#1A73E8"`: Menerapkan warna biru pada data pertama (`B2`).
4. `"color2", "#E8EAED"`: Menerapkan warna abu-abu lembut pada data kedua (`D2`), merepresentasikan ruang kosong menuju target.
5. `"max", C2`: Mengunci skala sel ke nilai target (`100`). Tanpa parameter `"max"`, formula akan menganggap total penjumlahan `B2+D2` sebagai 100%, menghilangkan efek perbandingan jika ada over-pencapaian.

Untuk kasus Siti Aminah yang mencapai 120% target, jika nilai `B3 > C3`, nilai `D3` adalah 0. Baris akan sepenuhnya berwarna biru solid tanpa merusak batas kanvas sel.

### Skenario Lanjutan: Win-Loss Streak Berbasis Filter Dinamis

Sebuah portofolio trading saham mencatat 15 riwayat eksekusi transaksi terakhir dalam rentang `H2:H16`. Nilai numerik merepresentasikan persentase gain/loss harian:

`{2.5; 1.2; -0.8; -1.5; 3.0; 0.0; 1.8; -2.2; 4.1; -0.5; -1.1; 0.9; 2.1; 1.4; -0.3}`

Buat visualisasi streak kemenangan (*Win/Loss Matrix*) pada sel tunggal `J2` yang hanya memfilter transaksi tidak bernilai nol, memberi warna hijau untuk profit, merah untuk rugi, serta menonjolkan profit transaksi pertama dan terakhir.

```excel
=SPARKLINE(FILTER(H2:H16; H2:H16<>0); {"charttype", "winloss"; "color", "#34A853"; "negcolor", "#EA4335"; "firstcolor", "#188038"; "lastcolor", "#F2994A"; "axis", TRUE; "axiscolor", "#5F6368"})
```

#### Bedah Logika Formula:
1. `FILTER(H2:H16; H2:H16<>0)`: Mengeliminasi nilai transaksi nol (`0.0`) dari data feeding visual, mencegah adanya gap kosong yang membingungkan pembaca.
2. `"charttype", "winloss"`: Mengubah angka bervariasi menjadi blok biner seragam (tinggi sama rata). Positif mengarah ke atas, negatif mengarah ke bawah.
3. `"firstcolor", "#188038"`: Memberikan penanda khusus hijau gelap pada transaksi pembuka periode.
4. `"lastcolor", "#F2994A"`: Menandai transaksi penutup dengan warna oranye untuk memudahkan pembaca melihat hasil transaksi paling mutakhir.

---

## Analisis Pitfall Teknis & Troubleshooting Runtime Error

Penerapan `SPARKLINE` pada skala enterprise sering memicu error rendering visual atau distorsi interpretasi data. Berikut panduan mitigasinya:

### 1. Masalah Separator Regional: Koma (,) vs Titik Koma (;)
Kesalahan paling umum yang memicu pesan `#ERROR!` (Formula Parse Error):

```excel
SALAH (Locale ID): =SPARKLINE(A1:A10, {"charttype", "line"})
BENAR (Locale ID): =SPARKLINE(A1:A10; {"charttype", "line"})
```

Di spreadsheet berbahasa Indonesia, argumen fungsi dipisahkan oleh titik koma (`;`). Jika menggunakan koma sebagai pemisah argumen fungsi, parser menganggap seluruh array kurung kurawal adalah bagian dari argumen pertama, memicu kegagalan kompilasi formula.

### 2. Error `#VALUE!` Akibat Array Key-Value Tidak Seimbang
`SPARKLINE` mewajibkan array opsi memiliki dimensi tepat 2 kolom: `[Kunci, Nilai]`.

```excel
SALAH: =SPARKLINE(A1:A10; {"charttype", "line", "color", "blue"})
BENAR: =SPARKLINE(A1:A10; {"charttype", "line"; "color", "blue"})
```

Pemisah antar-elemen dalam baris yang sama menggunakan koma (`,`), sedangkan transisi ke atribut baris berikutnya wajib menggunakan titik koma (`;`). Kesalahan pengetikan tanda baca ini merusak struktur matriks $N \times 2$ yang diharapkan Google Sheets, menghasilkan error `#VALUE! (Sparkline options must be key-value pairs)`.

### 3. Jebakan Distorsi Komparasi Antar-Baris (*The Independent Auto-Scale Fallacy*)
Saat formula `SPARKLINE` di-drag ke bawah sepanjang 50 baris, setiap sel secara default menghitung skala `min` dan `max` secara independen:

- Baris 2 memiliki rentang data `1` hingga `5`. Titik 5 berada di puncak atas sel.
- Baris 3 memiliki rentang data `1000` hingga `5000`. Titik 5000 juga berada di puncak atas sel.

Secara visual, kedua grafik terlihat identik, padahal magnitudo baris 3 adalah 1000 kali lebih besar.

**Solusi:** Kunci batas bawah dan atas menggunakan nilai absolut dari seluruh matriks tabel:

```excel
=SPARKLINE(B2:M2; {"charttype", "line"; "ymin", MIN($B$2:$M$50); "ymax", MAX($B$2:$M$50)})
```

Penggunaan tanda dollar absolut (`$`) memastikan seluruh baris tabel beroperasi di bawah skala referensi yang seragam.

### 4. Penanganan Nilai Teks, Kosong, dan Pembagian Nol (`#DIV/0!`)
Jika rentang sumber data mengandung hasil pembagian nol atau teks string kosong `""` hasil fungsi `IF`, grafik garis dapat terputus (*gap*) atau memunculkan pesan `#N/A`.

Tambahkan parameter penanganan data kotor:

```excel
=SPARKLINE(A1:A20; {"charttype", "line"; "empty", "zero"; "nan", "convert"})
```

- `"empty", "zero"`: Menafsirkan sel kosong bukan sebagai pemutus kurva, melainkan titik bernilai 0.
- `"nan", "convert"`: Mencegah kegagalan render bila terdapat teks, mengonversi string menjadi 0 secara visual tanpa merusak data asli.

---

## Integrasi Lanjutan: SPARKLINE Dinamis Menggunakan Google Apps Script

Bila konfigurasi opsi sparkline perlu diterapkan otomatis ke ratusan lembar kerja tanpa menyalin formula secara manual, gunakan Google Apps Script berikut untuk mengotomatisasi injeksi formula line chart dengan batas dinamis:

```javascript
/**
 * Mengotomatisasi penulisan formula SPARKLINE pada sheet aktif.
 * Mengunci skala Y-Min dan Y-Max secara global di seluruh baris.
 */
function generateDynamicSparklines() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;
  const numRows = 10;
  
  // Rentang data numerik: B2:M11 (12 bulan x 10 baris)
  const dataRange = sheet.getRange(startRow, 2, numRows, 12);
  const values = dataRange.getValues();
  
  // Ekstraksi nilai absolut min dan max dari seluruh matriks data
  let globalMin = Number.MAX_VALUE;
  let globalMax = Number.MIN_VALUE;
  
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      let val = values[r][c];
      if (typeof val === 'number') {
        if (val < globalMin) globalMin = val;
        if (val > globalMax) globalMax = val;
      }
    }
  }
  
  // Tulis formula SPARKLINE seragam ke Kolom N (indeks 14)
  for (let i = 0; i < numRows; i++) {
    let currentRow = startRow + i;
    let targetCell = sheet.getRange(currentRow, 14);
    
    // Format formula disesuaikan untuk regional titik-koma
    let formula = `=SPARKLINE(B${currentRow}:M${currentRow}; ` +
      `{"charttype", "line"; ` +
      `"color", "#1A73E8"; ` +
      `"linewidth", 2; ` +
      `"ymin", ${globalMin}; ` +
      `"ymax", ${globalMax}})`;
      
    targetCell.setFormula(formula);
  }
}
```

Script ini memastikan pembacaan seluruh rentang data terlebih dahulu, menghitung nilai terendah dan tertinggi mutlak dari data, lalu menghasilkan formula dengan skala yang sinkron di seluruh lembar kerja.

---

## Checklist Praktik Terbaik (Implementasi Produksi)

1. **Konsistensi Batas Sumbu:** Selalu sertakan `"ymin"` dan `"ymax"` (untuk `line`/`column`) atau `"max"` (untuk `bar`) jika Anda menampilkan grafik berdampingan antar-baris agar tidak menghasilkan ilusi optik visual.
2. **Definisikan Warna dengan Kode Heksadesimal:** Hindari string warna umum seperti `"red"` atau `"blue"`. Gunakan kode spesifik (seperti `"#1E8E3E"` untuk hijau atau `"#D93025"` untuk merah) guna menjaga konsistensi identitas merek pelaporan.
3. **Pembersihan Data Hulu:** Gunakan fungsi `FILTER` atau `INDEX` di dalam argumen data pertama untuk memangkas sel kosong di ujung array yang belum terisi.
4. **Optimasi Dimensi Sel:** Grafik in-cell sangat sensitif terhadap perubahan dimensi kanvas. Berikan lebar kolom minimal 120 piksel dan tinggi baris minimal 30 piksel agar grafik garis atau batang memiliki ruang rasio aspek (*aspect ratio*) yang memadai.
5. **Kunci Referensi Skala:** Gunakan notasi absolut `$` untuk parameter batas skala (misal: `MAX($B$2:$B$100)`) agar referensi tidak bergeser saat formula digandakan ke bawah.
6. **Dokumentasikan Pemisah Array:** Berikan catatan dokumentasi internal di spreadsheet jika dokumen diakses lintas negara, mengingat perbedaan separator regional (koma vs titik koma) kerap merusak formula saat diedit oleh pengguna dengan locale berbeda.
7. **Isolasi Nilai Negatif:** Pada grafik bertipe kolom atau winloss, manfaatkan selalu parameter `"negcolor"` agar pemangku kepentingan langsung mendeteksi anomali negatif tanpa harus membaca baris angka mentah satu per satu.

---

## Ringkasan Eksekutif Parameter SPARKLINE

| Tipe Grafik | Parameter Wajib Minimal | Parameter Kunci untuk Optimasi | Rekomendasi Kasus Penggunaan |
| :--- | :--- | :--- | :--- |
| **`"line"`** | `{"charttype", "line"}` | `"ymin"`, `"ymax"`, `"color"`, `"linewidth"` | Tren time-series finansial bulanan, log suhu, analitik web mingguan. |
| **`"column"`** | `{"charttype", "column"}` | `"highcolor"`, `"lowcolor"`, `"negcolor"`, `"axis"` | Distribusi laba-rugi, perbandingan volume bulanan, pencapaian kuartalan. |
| **`"bar"`** | `{"charttype", "bar"; "max", N}` | `"color1"`, `"color2"` | Progres target KPI, alokasi kapasitas server, completion rate proyek. |
| **`"winloss"`** | `{"charttype", "winloss"}` | `"negcolor"`, `"axis"`, `"lastcolor"` | Rekam jejak menang/kalah tim olahraga, streak profit transaksi pasar modal. |