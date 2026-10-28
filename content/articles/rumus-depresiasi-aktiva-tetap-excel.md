---
title: "Menghitung Penyusutan Aset Tetap di Excel: Metode Garis Lurus (SLN) dan Saldo Menurun (DB)"
description: "Implementasi formula finansial SLN, DB, dan DDB untuk menghitung jadwal amortisasi dan depresiasi aktiva tetap perusahaan."
category: "Excel"
date: "2026-10-28"
readTime: "8 menit baca"
---

## 1. Pendahuluan

Pencatatan aset tetap (*fixed assets*) menuntut presisi matematis dan kepatuhan terhadap regulasi akuntansi (PSAK 16 / IAS 16). Setiap aset tetap berwujud—kecuali tanah—kehilangan nilai ekonomis akibat keausan fisik, intensitas penggunaan, maupun keusangan teknologi. Depresiasi berfungsi mengalokasikan harga perolehan aset menjadi beban operasional secara berkala sepanjang estimasi masa manfaatnya (*useful life*).

Dalam praktik audit finansial, kesalahan alokasi depresiasi memicu distorsi material pada Laporan Laba Rugi (*Income Statement*) dan Neraca (*Balance Sheet*). Beban depresiasi yang dicatat terlalu rendah menyebabkan nilai laba bersih overvalued dan nilai buku aset overstated, memicu risiko temuan salah saji oleh auditor eksternal. Sebaliknya, depresiasi yang terlalu agresif menekan profitabilitas perusahaan secara artifisial.

Banyak staf keuangan masih menggunakan tabel kalkulasi manual atau rumus pembagian statis yang rentan terhadap *formula drift* dan kesalahan *copy-paste*. Penerapan formula finansial terstandarisasi di Microsoft Excel—khususnya fungsi `SLN`, `DB`, dan `DDB`—mengotomasi pembaruan jadwal penyusutan, mengeliminasi galat manusiawi saat terjadi penyesuaian asumsi nilai sisa (*salvage value*), dan memastikan model laporan keuangan tetap adaptif terhadap perubahan fiskal.

---

## 2. Anatomi Sintaks Fungsi Finansial Excel

Excel menyediakan beberapa fungsi finansial untuk menangani berbagai metode depresiasi. Masing-masing fungsi memiliki struktur argumen dan asumsi perhitungan matematis yang berbeda.

```
       SLN Formula                   DB Formula (Declining Balance)
 ┌──────────────────────┐     ┌──────────────────────────────────────────┐
 │ =SLN(cost, sal, life)│     │ =DB(cost, sal, life, period, [month])   │
 └──────────────────────┘     └──────────────────────────────────────────┘
             │                                     │
             ▼                                     ▼
 Beban Tetap Rata per Tahun            Beban Lebih Besar di Tahun Awal
 (Sesuai Karakter Aset Pasif)          (Sesuai Aset Berbasis Teknologi/Mesin)
```

### 2.1 Fungsi SLN (Straight-Line)

Fungsi `SLN` menghitung beban penyusutan garis lurus untuk satu periode. Nilai beban yang dihasilkan identik untuk setiap periode amortisasi.

```excel
=SLN(cost, salvage, life)
```

*   `cost` (Wajib): Nilai perolehan awal aset, mencakup harga beli dan seluruh biaya langsung terkait persiapan aset hingga siap beroperasi (misal: bea masuk, pengiriman, biaya instalasi teknis).
*   `salvage` (Wajib): Estimasi nilai sisa atau nilai jual kembali aset pada akhir masa manfaatnya.
*   `life` (Wajib): Jumlah periode masa manfaat aset (biasanya dinyatakan dalam tahun atau bulan).

Rumus matematika di balik fungsi ini:

$$\text{Depresiasi SLN} = \frac{\text{Cost} - \text{Salvage}}{\text{Life}}$$

### 2.2 Fungsi DB (Fixed-Declining Balance)

Fungsi `DB` menghitung penyusutan menggunakan metode saldo menurun dengan tingkat bunga/tarif tetap (*fixed-rate*). Beban depresiasi bernilai paling tinggi pada periode awal dan menurun secara proporsional seiring menurunnya nilai buku (*book value*).

```excel
=DB(cost, salvage, life, period, [month])
```

*   `cost` (Wajib): Nilai perolehan aset.
*   `salvage` (Wajib): Estimasi nilai residu pada akhir masa manfaat.
*   `life` (Wajib): Total periode masa manfaat aset.
*   `period` (Wajib): Periode spesifik yang dihitung beban penyusutannya. Nilai harus menggunakan satuan waktu yang sama dengan `life`.
*   `month` (Opsional): Jumlah bulan pemanfaatan aset pada tahun pertama buku. Nilai default adalah 12 jika parameter ini dikosongkan.

Excel menghitung tingkat suku penyusutan (*depreciation rate*) internal dengan rumus:

$$\text{Rate} = 1 - \left(\frac{\text{Salvage}}{\text{Cost}}\right)^{\frac{1}{\text{Life}}}$$

Nilai `Rate` tersebut dibulatkan oleh mesin Excel hingga 3 tempat desimal (*three decimal places*).

Untuk periode pertama ($period = 1$), jika nilai $month < 12$:

$$\text{Beban Periode 1} = \text{Cost} \times \text{Rate} \times \frac{\text{Month}}{12}$$

Untuk periode selanjutnya ($2 \le period \le life$):

$$\text{Beban Periode } t = (\text{Cost} - \sum \text{Depresiasi Sebelumnya}) \times \text{Rate}$$

Untuk periode terakhir ($period = life + 1$, apabila aset diakuisisi di pertengahan tahun):

$$\text{Beban Periode Akhir} = (\text{Cost} - \sum \text{Depresiasi Sebelumnya}) \times \text{Rate} \times \frac{12 - \text{Month}}{12}$$

### 2.3 Fungsi DDB (Double-Declining Balance)

Fungsi `DDB` menghitung saldo menurun ganda dengan mengalikan tarif garis lurus dengan sebuah faktor akselerasi (default adalah 2).

```excel
=DDB(cost, salvage, life, period, [factor])
```

*   `cost`, `salvage`, `life`, `period`: Identik dengan fungsi `DB`.
*   `factor` (Opsional): Kecepatan penurunan saldo. Jika tidak diisi, Excel mengasumsikan nilai `2` (metode saldo menurun ganda).

Rumus dasar penentuan tarif:

$$\text{Tarif DDB} = \frac{\text{Factor}}{\text{Life}}$$

Fungsi `DDB` secara otomatis menghentikan depresiasi atau menurunkan nilai beban tahun berjalan apabila akumulasi penyusutan akan melampaui batas $(\text{Cost} - \text{Salvage})$. Nilai buku tidak akan pernah turun di bawah nilai `salvage`.

---

## 3. Matriks Perbandingan Karakteristik Metode

| Parameter | Garis Lurus (`SLN`) | Saldo Menurun (`DB`) | Saldo Menurun Ganda (`DDB`) | Kalkulasi Manual Non-Formula |
| :--- | :--- | :--- | :--- | :--- |
| **Metodologi Beban** | Konstan / Flat | Menurun bertahap berdasarkan tarif tetap | Menurun drastis pada tahun-tahun awal | Statis / Sering inkonsisten |
| **Profil Risiko Audit** | Sangat Rendah | Rendah - Moderat | Rendah - Moderat | Tinggi (rawan kesalahan rumus) |
| **Dampak Pajak Awal** | Beban pajak merata | Menekan beban pajak tahun pertama | Menekan beban pajak paling signifikan di awal | Tidak dapat diprediksi |
| **Presisi Nilai Residu** | Tepat sesuai input | Deviasi kecil akibat pembulatan tarif 3 desimal | Dijamin tidak tembus di bawah nilai residu | Memerlukan intervensi manual |
| **Relevansi Aset** | Bangunan, furnitur kantor, aset pasif | Mesin pabrik, perangkat IT, kendaraan operasional | Peralatan komputasi performa tinggi, server | Tidak disarankan untuk tata kelola enterprise |

---

## 4. Studi Kasus 1: Jadwal Amortisasi Mesin Cetak Industri (SLN)

### Profil Kasus
Perusahaan manufaktur membeli mesin cetak industri pada awal tahun buku:
*   Harga Pembelian & Pemasangan (`Cost`): Rp 450.000.000
*   Nilai Residu Diestimasi (`Salvage`): Rp 50.000.000
*   Masa Manfaat Ekonomis (`Life`): 5 Tahun

### Struktur Lembar Kerja Excel
Susun data input pada rentang sel `A1:B3`:

| Sel | Label Parameter | Nilai | Format Sel |
| :--- | :--- | :--- | :--- |
| **A1** | Nilai Perolehan (Cost) | `450000000` | `Currency (Rp)` |
| **A2** | Nilai Sisa (Salvage) | `50000000` | `Currency (Rp)` |
| **A3** | Masa Manfaat (Tahun) | `5` | `General` |

Tabel jadwal penyusutan disiapkan pada rentang `A6:E11`:

| Baris | A (Periode/Tahun) | B (Nilai Buku Awal) | C (Beban Depresiasi) | D (Akumulasi Depresiasi) | E (Nilai Buku Akhir) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **6** | Periode | Saldo Awal | Beban Tahun Berjalan | Akumulasi Beban | Saldo Akhir |
| **7** | 1 | `=$B$1` | `=SLN($B$1; $B$2; $B$3)` | `=C7` | `=B7-C7` |
| **8** | 2 | `=E7` | `=SLN($B$1; $B$2; $B$3)` | `=D7+C8` | `=B8-C8` |
| **9** | 3 | `=E8` | `=SLN($B$1; $B$2; $B$3)` | `=D8+C9` | `=B9-C9` |
| **10**| 4 | `=E9` | `=SLN($B$1; $B$2; $B$3)` | `=D9+C10` | `=B10-C10` |
| **11**| 5 | `=E10` | `=SLN($B$1; $B$2; $B$3)` | `=D10+C11` | `=B11-C11` |

### Implementasi Formula Dinamis
Untuk menghindari penulisan berulang, masukkan formula berikut pada baris 7, lalu salin (*drag down*) hingga baris 11:

```excel
Kolom B (B7): =IF(A7=1; $B$1; E6)
Kolom C (C7): =SLN($B$1; $B$2; $B$3)
Kolom D (D7): =SUM($C$7:C7)
Kolom E (E7): =B7-C7
```

*Catatan: Pola `=SUM($C$7:C7)` mengunci titik pangkal ekspansi sehingga menghasilkan *running total* yang dinamis tanpa memicu sirkular referensi.*

### Output Hasil Kalkulasi
Setelah evaluasi formula, tabel terisi nilai numerik berikut:

| Tahun (A) | Saldo Awal (B) | Beban Depresiasi (C) | Akumulasi Depresiasi (D) | Saldo Akhir (E) |
| :--- | :--- | :--- | :--- | :--- |
| **1** | Rp 450.000.000 | Rp 80.000.000 | Rp 80.000.000 | Rp 370.000.000 |
| **2** | Rp 370.000.000 | Rp 80.000.000 | Rp 160.000.000 | Rp 290.000.000 |
| **3** | Rp 290.000.000 | Rp 80.000.000 | Rp 240.000.000 | Rp 210.000.000 |
| **4** | Rp 210.000.000 | Rp 80.000.000 | Rp 320.000.000 | Rp 130.000.000 |
| **5** | Rp 130.000.000 | Rp 80.000.000 | Rp 400.000.000 | **Rp 50.000.000** |

Perhatikan bahwa Nilai Buku Akhir pada tahun ke-5 (Sel `E11`) bernilai persis Rp 50.000.000, setara dengan parameter `Salvage` di sel `B2`.

---

## 5. Studi Kasus 2: Komputasi Infrastruktur Server dengan Akuisisi Parsial (DB vs DDB)

### Profil Kasus Lanjutan
Perusahaan teknologi membeli unit Rack Server Data Center:
*   Tanggal Akuisisi: 1 September 2026 (Tahun fiskal berakhir 31 Desember. Durasi operasional tahun pertama = 4 bulan).
*   Nilai Perolehan (`Cost`): Rp 180.000.000
*   Nilai Sisa (`Salvage`): Rp 15.000.000
*   Masa Manfaat (`Life`): 4 Tahun
*   Bulan Berjalan Tahun Pertama (`Month`): 4

Karena aset dioperasikan di pertengahan tahun, jadwal depresiasi mencakup $4 + 1 = 5$ periode fiskal.

### Skema Setup Input Parameter
Letakkan metadata pada range `B1:B5`:
*   `B1`: `180000000` (Cost)
*   `B2`: `15000000` (Salvage)
*   `B3`: `4` (Life)
*   `B4`: `4` (Month)

```
        Alur Garis Waktu Aset (Akuisisi 1 September)
  Th 1: 4 Bulan │     Th 2: 12 Bulan     │     Th 3: 12 Bulan     │     Th 4: 12 Bulan     │ Th 5: 8 Bulan
├───────────────┼────────────────────────┼────────────────────────┼────────────────────────┼───────────────┤
  Sep      Des   Jan                Des   Jan                Des   Jan                Des   Jan         Agt
```

### Penulisan Formula Komparasi Multi-Kolom
Susun jadwal amortisasi dinamis dari baris 8 hingga baris 13:

```excel
Kolom A (Periode): 1, 2, 3, 4, 5

Kolom B (Beban DB):
=DB($B$1; $B$2; $B$3; A8; $B$4)

Kolom C (Nilai Buku Akhir DB):
=$B$1-SUM($B$8:B8)

Kolom D (Beban DDB):
=DDB($B$1; $B$2; $B$3; A8; 2)

Kolom E (Nilai Buku Akhir DDB):
=$B$1-SUM($D$8:D8)
```

*Perhatian: Parameter `[month]` tersedia langsung di fungsi `DB`. Fungsi `DDB` standar tidak menyediakan parameter bulan parsial secara internal, sehingga penggunaannya untuk tahun berjalan parsial memerlukan penyesuaian durasi atau beralih ke fungsi `VDB` (Variable Declining Balance).*

Berikut adalah formula penanganan tahun parsial presisi tinggi pada `DDB` menggunakan bantuan evaluasi logika:

```excel
=IF(A8=1; DDB($B$1; $B$2; $B$3; 1) * ($B$4/12); IF(A8<= $B$3; DDB($B$1 - ($B$1 * (2/$B$3) * ($B$4/12)); $B$2; $B$3; A8 - 1 + ($B$4/12)); 0))
```

Agar jadwal komparasi tetap transparan dan mudah diaudit, perhatikan tabel evaluasi numerik fungsi `DB` bawaan Excel:

### Hasil Eksekusi Fungsi DB dengan Parameter Month = 4

Perhitungan suku tarif DB internal:
$$\text{Rate} = 1 - (15.000.000 / 180.000.000)^{1/4} = 1 - (0,08333)^{0,25} = 1 - 0,53728 = 0,4627 \approx 0,463$$

| Periode (A) | Durasi Operasional | Beban DB (B) | Akumulasi DB | Saldo Buku Akhir DB (C) |
| :--- | :--- | :--- | :--- | :--- |
| **1** | 4 Bulan | Rp 27.780.000 | Rp 27.780.000 | Rp 152.220.000 |
| **2** | 12 Bulan | Rp 70.477.860 | Rp 98.257.860 | Rp 81.742.140 |
| **3** | 12 Bulan | Rp 37.846.611 | Rp 136.104.471 | Rp 43.895.529 |
| **4** | 12 Bulan | Rp 20.323.630 | Rp 156.428.101 | Rp 23.571.899 |
| **5** | 8 Bulan | Rp 7.275.986 | Rp 163.704.087 | **Rp 16.295.913** |

### Analisis Deviasi Pembulatan DB
Nilai buku akhir periode 5 adalah Rp 16.295.913, bukan tepat Rp 15.000.000. Deviasi sebesar Rp 1.295.913 ini merupakan sifat bawaan (*inherent behavior*) dari algoritma fungsi `DB` Excel. Mesin Excel membulatkan nilai `Rate` menjadi tiga desimal (`0,463`), yang secara matematis menghasilkan selisih akumulasi pada akhir masa pakai aktiva.

Untuk menutup deviasi ini pada periode pelaporan final agar neraca seimbang (*reconciled*), modifikasi formula beban pada sel periode akhir (Tahun 5):

```excel
=IF(A12=$B$3+1; C11-$B$2; DB($B$1; $B$2; $B$3; A12; $B$4))
```

Formula di atas memaksa pemotongan beban tahun terakhir tepat sebesar sisa saldo buku dikurangi nilai residu target ($23.571.899 - 15.000.000 = 8.571.899$).

---

## 6. Automasi Alternatif: Google Apps Script

Bagi organisasi yang mengintegrasikan ekosistem spreadsheet daring untuk sinkronisasi inventaris aset secara otomatis, fungsi JavaScript berikut mereproduksi logika matematis `SLN` dan `DB` tanpa ketergantungan pada library pihak ketiga:

```javascript
/**
 * Menghitung penyusutan garis lurus (Straight-Line Depreciation).
 * @param {number} cost - Harga perolehan aset.
 * @param {number} salvage - Nilai residu aset.
 * @param {number} life - Masa manfaat aset (dalam tahun/periode).
 * @return {number} Beban penyusutan per periode.
 */
function calculateSLN(cost, salvage, life) {
  if (life <= 0) {
    throw new Error("Parameter life harus bernilai lebih besar dari 0.");
  }
  if (cost < salvage) {
    throw new Error("Nilai perolehan (cost) tidak boleh lebih rendah dari nilai residu (salvage).");
  }
  return (cost - salvage) / life;
}

/**
 * Menghitung penyusutan metode saldo menurun (Fixed-Declining Balance).
 * Meniru algoritma pembulatan 3-desimal Microsoft Excel.
 * @param {number} cost - Harga perolehan aset.
 * @param {number} salvage - Nilai sisa aset.
 * @param {number} life - Masa manfaat aset dalam tahun.
 * @param {number} period - Periode yang dihitung.
 * @param {number} month - Jumlah bulan di tahun pertama (1-12). Default 12.
 * @return {number} Beban penyusutan pada periode tersebut.
 */
function calculateDB(cost, salvage, life, period, month) {
  month = month || 12;

  if (cost <= 0 || life <= 0 || period <= 0) {
    throw new Error("Parameter cost, life, dan period harus bernilai positif.");
  }
  if (period > life + 1) {
    return 0;
  }

  // Hitung rate depresiasi dan bulatkan ke 3 angka desimal sesuai spesifikasi Excel
  let rawRate = 1 - Math.pow(salvage / cost, 1 / life);
  let rate = Math.round(rawRate * 1000) / 1000;

  // Beban periode 1
  let initialDepreciation = cost * rate * (month / 12);
  if (period === 1) {
    return initialDepreciation;
  }

  let accumulatedDepreciation = initialDepreciation;
  let currentBookValue = cost - accumulatedDepreciation;
  let currentDepreciation = 0;

  // Iterasi perhitungan hingga periode target
  for (let p = 2; p <= period; p++) {
    if (p === life + 1) {
      // Periode akhir penyesuaian bulan parsial
      currentDepreciation = currentBookValue * rate * ((12 - month) / 12);
    } else {
      currentDepreciation = currentBookValue * rate;
    }
    accumulatedDepreciation += currentDepreciation;
    currentBookValue = cost - accumulatedDepreciation;
  }

  return currentDepreciation;
}
```

---

## 7. Jebakan Teknis & Penanganan Error (*Troubleshooting*)

```
             Tipe Error Formula Penyusutan di Excel
                               │
       ┌───────────────────────┴───────────────────────┐
       ▼                                               ▼
   #NUM! Error                                    #VALUE! Error
   • Salvage >= Cost                              • Angka berformat Teks
   • Life <= 0                                    • Separator salah ("," vs ";")
   • Period > Life (+ 1)                          • Nilai non-numerik di argumen
```

### 7.1 Error `#NUM!` (Invalid Numeric Arguments)
*Penyebab:*
1.  Nilai `salvage` bernilai negatif atau melampaui nilai `cost`. Formula logaritma internal pada fungsi `DB` tidak dapat mengevaluasi bilangan real negatif.
2.  Argumen `life` diisi $\le 0$. Nilai pembagi nol memicu *runtime exception*.
3.  Nilai `period` pada fungsi `DB` melebihi `life + 1` (atau melebihi `life` ketika `month = 12`).

*Solusi Defensif:*
Bungkus formula dengan validasi logika `IF`:

```excel
=IF(OR(B1<=0; B3<=0); "Parameter Tidak Valid"; SLN(B1; B2; B3))
```

### 7.2 Error `#VALUE!` (Type Mismatch)
*Penyebab:*
Salah satu sel input berisi karakter string, spasi tak terlihat, atau separator ribuan yang tidak dikenali sebagai angka akibat inkonsistensi regional settings (misal: memasukkan format teks `"Rp 450.000.000"` secara manual tanpa mengonfigurasi fitur *Cell Formatting*).

*Solusi Teknis:*
Gunakan formula pembersih data numerik otomatis:

```excel
=SLN(VALUE(TRIM(CLEAN(B1))); VALUE(TRIM(CLEAN(B2))); B3)
```

### 7.3 Masalah Regional Settings: Tanda Koma (,) vs Titik Koma (;)
Format pemisah argumen ditentukan oleh konfigurasi sistem operasi:
*   **Standar US/Internasional:** Memakai tanda koma `,` untuk memisahkan argumen rumus:
    `=SLN(cost, salvage, life)`
*   **Standar Indonesia/Eropa Continental:** Memakai tanda titik koma `;` karena tanda koma dialokasikan sebagai penanda desimal:
    `=SLN(cost; salvage; life)`

*Tindakan Preventif:*
Periksa petunjuk konteks rumus (*tooltips*) saat mengetik fungsi di formula bar Excel. Jika *tooltips* menampilkan `[cost, salvage, life]`, gunakan koma. Jika menampilkan `[cost; salvage; life]`, gunakan titik koma.

### 7.4 Kerusakan Jadwal Akibat Kesalahan Referensi Sel (Absolut vs Relatif)
Pengisian tabel amortisasi menggunakan fitur *AutoFill* sering rusak karena kegagalan penguncian sel referensi parameter perolehan.

*   `B1` (Relatif): Bergeser ke `B2`, `B3`, dst. saat disalin ke bawah. Mengakibatkan nilai parameter tertukar.
*   `$B$1` (Absolut Sempurna): Selalu merujuk pada sel `B1` di mana pun rumus ditempatkan. Formula penyusutan wajib mengunci referensi `cost`, `salvage`, dan `life` menggunakan tombol `F4` (`$B$1; $B$2; $B$3`).
*   `A8` (Relatif Baris): Biarkan nomor baris periode tanpa simbol `$` agar nomor periode bertambah seiring penyalinan baris ke bawah.

---

## 8. Ringkasan Eksekutif & Praktik Terbaik Audit

Sebelum mendistribusikan model penyusutan aset ke unit bisnis atau auditor, lakukan validasi kelayakan menggunakan daftar periksa berikut:

```
[ ] Seluruh parameter input (Cost, Salvage, Life) terisolasi di sel khusus (Data Input Block).
[ ] Tidak ada nilai konstan (hardcoded numbers) di dalam formula amortisasi.
[ ] Formula referensi parameter telah terkunci menggunakan simbol absolut ($).
[ ] Nilai buku akhir (Ending Book Value) periode final sama persis dengan Nilai Residu target.
[ ] Format sel seluruh nominal finansial menggunakan format akuntansi seragam.
[ ] Tipe metode penyusutan telah divalidasi dan disetujui sesuai kebijakan audit perusahaan.
```

1.  **Garis Lurus (`SLN`)** adalah standar baku untuk aset yang memberikan utilitas seragam tanpa depresiasi fungsi yang cepat. Metode ini menjamin stabilitas laporan beban di buku besar.
2.  **Saldo Menurun (`DB` / `DDB`)** harus diterapkan pada aset teknologi, armada transportasi, dan permesinan berat yang efisiensi kerjanya menurun tajam seiring waktu atau membutuhkan biaya pemeliharaan (*maintenance cost*) yang kian membengkak di masa tua.
3.  **Rekonsiliasi Deviasi:** Saat menerapkan fungsi `DB` untuk kebutuhan pelaporan formal, perhitungkan selisih pembulatan 3-desimal bawaan Excel pada periode akhir agar rekonsiliasi neraca tidak menghasilkan selisih material.