---
title: "Cara Membuat Dropdown Bertingkat di Excel dengan INDIRECT"
description: "Langkah menyusun dependent drop-down list (Kategori > Sub-Kategori) di Excel menggunakan Data Validation dan formula INDIRECT."
category: "Excel"
date: "2026-10-28"
readTime: "8 menit baca"
---

# Panduan Implementasi Dependent Dropdown List di Microsoft Excel Menggunakan Fungsi INDIRECT

Penyusunan basis data pada Microsoft Excel sering mengalami anomali akibat inkonsistensi input manual (*human error*). Kesalahan pengetikan, penggunaan sinonim tidak terstandarisasi, dan pemilihan data anak (*sub-category*) yang tidak sesuai dengan data induk (*parent category*) menurunkan integritas data saat proses transformasi data (ETL) dan visualisasi *dashboard*.

Solusi arsitektural untuk mencegah inkonsistensi ini adalah penerapan *Dependent Dropdown List* (daftar pilihan bertingkat). Mekanisme ini mengunci opsi dropdown kedua agar hanya menampilkan subset data yang valid berdasarkan nilai terpilih pada dropdown pertama. Pendekatan standar industri memanfaatkan kombinasi fitur **Data Validation**, penetapan **Defined Names (Named Range)**, dan evaluasi referensi dinamis melalui fungsi **INDIRECT**.

---

## 1. Anatomi Sintaks dan Komponen Teknis

Fungsi inti dalam pembentukan dropdown dinamis adalah `INDIRECT`. Fungsi ini mengonversi representasi string teks menjadi referensi sel atau rentang (*range*) yang valid secara komputasional.

### Sintaks Dasar Formula

```excel
=INDIRECT(ref_text, [a1])
```

### Parameter Komputasi

1. **`ref_text` (Wajib):**
   Teks yang merepresentasikan referensi sel, nama rentang (*named range*), atau alamat sel dalam bentuk teks. Jika `ref_text` tidak merujuk pada referensi sel yang valid atau nama rentang yang terdaftar di *Name Manager*, formula mengembalikan galat `#REF!`.
2. **`[a1]` (Opsional):**
   Nilai logika (*boolean*) penentu format referensi:
   * `TRUE` atau diabaikan: `ref_text` diinterpretasikan sebagai referensi gaya A1.
   * `FALSE`: `ref_text` diinterpretasikan sebagai referensi gaya R1C1. Untuk implementasi Data Validation, parameter ini dibiarkan default (`TRUE`).

### Mekanisme Kerja Internal

Ketika sel A2 berisi string `"Elektronik"`, formula `=INDIRECT(A2)` tidak mengambil nilai teks `"Elektronik"`, melainkan mencari objek bernama `Elektronik` dalam memori kerja Excel. Objek tersebut dapat berupa sel tunggal atau kumpulan sel (*Named Range*). Excel kemudian mengembalikan rentang nilai yang dicakup oleh nama tersebut ke mesin *Data Validation*.

### Penanganan Spasi dengan Fungsi Pendukung

Aturan penamaan *Named Range* Excel melarang penggunaan karakter spasi. Jika kategori induk berisi string multi-kata (misal: "Alat Tulis"), formula harus mentranslasikan spasi menjadi karakter garis bawah (*underscore*) menggunakan fungsi `SUBSTITUTE`:

```excel
=INDIRECT(SUBSTITUTE(A2, " ", "_"))
```

---

## 2. Perbandingan Metode Validasi Data Dependent

| Parameter | Validasi Manual Tunggal | Nested IF (IF Bersarang) | INDIRECT + Named Range | FILTER + Spill Array (Excel 365) |
| :--- | :--- | :--- | :--- | :--- |
| **Skalabilitas Data** | Sangat Rendah | Sangat Rendah (Limit 64 tingkat) | Tinggi | Sangat Tinggi |
| **Kompleksitas Rumus** | Tidak ada | Sangat Rumit jika kategori > 5 | Statis & Ringkas (`=INDIRECT()`) | Ringkas (`=FILTER()`) |
| **Beban Pemeliharaan** | Tinggi | Tinggi (Edit formula manual) | Rendah (Update *Name Manager*) | Rendah (Otomatis menyesuaikan tabel) |
| **Kompatibilitas** | Semua Versi Excel | Semua Versi Excel | Excel 2003 – Excel 365 | Hanya Excel 365 & Excel 2021 |
| **Kinerja Memori** | Statis (Ringan) | Statis (Ringan) | Volatile (Dihitung ulang tiap perubahan) | Semi-Dinamis (Ringan) |

Metode `INDIRECT` merupakan titik tengah optimal bagi lingkungan perusahaan yang membutuhkan kompatibilitas lintas versi (Excel 2010/2013/2016/2019/365) tanpa memerlukan pembaruan kode makro VBA.

---

## 3. Studi Kasus 1: Dropdown Bertingkat 2-Tier (Kategori > Sub-Kategori)

### Skenario Masalah
Divisi Pengadaan Barang (*Procurement*) membutuhkan standarisasi input inventaris aset kantor. Pengguna hanya diizinkan memilih Sub-Kategori yang benar-benar terdaftar di bawah Kategori terkait.

### Skema Data Master

Data referensi ditempatkan pada lembar kerja terpisah, diberi nama sheet `Data_Master`.

Rentang sel `$A$1:$C$5` dikonfigurasi sebagai berikut:

| Sel | Kolom A (Hardware) | Kolom B (Software) | Kolom C (Jaringan) |
| :--- | :--- | :--- | :--- |
| **Baris 1** | **Hardware** | **Software** | **Jaringan** |
| **Baris 2** | Laptop | Windows OS | Router |
| **Baris 3** | Monitor | MS Office | Switch |
| **Baris 4** | Keyboard | Antivirus | Access Point |
| **Baris 5** | Docking Station | Adobe CC | Kabel LAN |

### Lembar Kerja Transaksi

Data transaksi dicatat pada sheet `Input_Transaksi`:

| Sel | Kolom A (Kategori Terpilih) | Kolom B (Sub-Kategori Terpilih) |
| :--- | :--- | :--- |
| **Baris 1** | **Kategori Aset** | **Sub-Kategori Aset** |
| **Baris 2** | *[Dropdown Level 1]* | *[Dropdown Level 2]* |

### Prosedur Implementasi Langkah demi Langkah

#### Tahap 1: Pembentukan Named Range Secara Massal
1. Blok rentang `$A$1:$C$5` pada sheet `Data_Master`.
2. Buka ribbon tab **Formulas** > grup **Defined Names** > klik **Create from Selection** (Pintasan: `Ctrl + Shift + F3`).
3. Centang hanya kotak **Top row**, hilangkan centang lainnya.
4. Klik **OK**. Excel secara otomatis membuat tiga *Named Range* independen:
   * `Hardware` merujuk ke `=Data_Master!$A$2:$A$5`
   * `Software` merujuk ke `=Data_Master!$B$2:$B$5`
   * `Jaringan` merujuk ke `=Data_Master!$C$2:$C$5`
5. Buat satu *Named Range* tambahan untuk daftar kategori induk:
   * Buka ribbon **Formulas** > **Name Manager** > **New**.
   * Isi kolom *Name* dengan: `Kategori_Utama`.
   * Isi kolom *Refers to* dengan: `=Data_Master!$A$1:$C$1`.
   * Klik **OK**.

#### Tahap 2: Pembuatan Dropdown Level 1 (Kategori)
1. Buka sheet `Input_Transaksi`.
2. Klik sel `A2` (atau blok rentang `A2:A100` untuk massal).
3. Buka tab **Data** > grup **Data Tools** > klik **Data Validation**.
4. Pada tab *Settings*, ubah dropdown *Allow* menjadi **List**.
5. Pada kotak *Source*, masukkan formula:
   ```excel
   =Kategori_Utama
   ```
6. Pastikan opsi **In-cell dropdown** dicentang. Klik **OK**.

#### Tahap 3: Pembuatan Dropdown Level 2 (Sub-Kategori Dependent)
1. Pada sheet `Input_Transaksi`, pilih sel `B2` (atau blok `B2:B100`).
2. Buka menu **Data Validation**.
3. Pilih opsi **List** pada kolom *Allow*.
4. Pada kotak isian *Source*, masukkan formula berikut:
   ```excel
   =INDIRECT($A2)
   ```
   *Catatan kritis:* Tanda pengunci kolom absolut `$` pada `$A2` penting agar saat validasi disalin ke baris bawahnya (`B3`, `B4`, dst.), formula otomatis merujuk ke baris yang sesuai (`$A3`, `$A4`), namun tetap mengunci referensi pada kolom `A`.
5. Klik **OK**. Jika muncul peringatan: *"The Source currently evaluates to an error. Do you wish to continue?"*, klik **Yes**. Peringatan ini muncul normal karena sel `A2` saat itu masih dalam keadaan kosong.

---

## 4. Studi Kasus 2: Penanganan Spasi dan 3-Tier Multi-Level Validation

### Skenario Lanjutan
Entitas bisnis memerlukan tiga tingkatan klasifikasi: **Divisi** > **Departemen** > **Unit Kerja**. Nama kategori memiliki spasi, seperti "Sumber Daya Manusia" atau "Teknologi Informasi".

### Matriks Data Master

Struktur penamaan disiapkan pada sheet `Master_Hierarki` menggunakan konvensi garis bawah (*underscore*) sebagai substitusi spasi:

#### Tier 1 (Tingkat Divisi)
Rentang `$A$1:$B$1`:
* `A1`: `Teknologi_Informasi`
* `B1`: `Operasional_Bisnis`

Range diberi nama: `Daftar_Divisi`.

#### Tier 2 (Tingkat Departemen)
Di bawah kolom `Teknologi_Informasi` (`$A$2:$A$3`):
* `Infrastruktur`
* `Pengembangan_Sistem`

Rentang `$A$2:$A$3` diberi nama: `Teknologi_Informasi`.

Di bawah kolom `Operasional_Bisnis` (`$B$2:$B$3`):
* `Logistik`
* `Pengadaan`

Rentang `$B$2:$B$3` diberi nama: `Operasional_Bisnis`.

#### Tier 3 (Tingkat Unit Kerja)
Unit kerja dinamai dengan menggabungkan nama parent (Tier 2) agar terisolasi dari duplikasi nama departemen:

Rentang `$D$2:$D$4` (Daftar unit di bawah `Infrastruktur`):
* `Cloud_Admin`
* `Network_Engineer`
* `Helpdesk`

Beri nama range `$D$2:$D$4`: `Infrastruktur`.

Rentang `$E$2:$E$4` (Daftar unit di bawah `Pengembangan_Sistem`):
* `Backend_Dev`
* `Frontend_Dev`
* `Database_Admin`

Beri nama range `$E$2:$E$4`: `Pengembangan_Sistem`.

Rentang `$F$2:$F$3` (Daftar unit di bawah `Logistik`):
* `Armada`
* `Pergudangan`

Beri nama range `$F$2:$F$3`: `Logistik`.

Rentang `$G$2:$G$3` (Daftar unit di bawah `Pengadaan`):
* `Vendor_Management`
* `Purchasing`

Beri nama range `$G$2:$G$3`: `Pengadaan`.

### Implementasi Formula Multi-Tier

Struktur target pada tabel pencatatan:
* Sel `A2`: Pilihan Divisi (Dapat menggunakan string dengan spasi, misal: "Teknologi Informasi").
* Sel `B2`: Pilihan Departemen.
* Sel `C2`: Pilihan Unit Kerja.

#### Formula Validasi Sel A2 (Tier 1)
Pada opsi *List*, masukkan sumber formula:
```excel
=Daftar_Divisi
```

#### Formula Validasi Sel B2 (Tier 2 - Mengatasi Spasi)
Karena nilai pada `A2` berisi spasi ("Teknologi Informasi"), fungsi `INDIRECT` akan gagal jika langsung memanggil teks tersebut. Gunakan fungsi `SUBSTITUTE`:
```excel
=INDIRECT(SUBSTITUTE($A2, " ", "_"))
```
Logika evaluasi:
1. `$A2` bernilai `"Teknologi Informasi"`.
2. `SUBSTITUTE("Teknologi Informasi", " ", "_")` mengubah teks menjadi `"Teknologi_Informasi"`.
3. `INDIRECT("Teknologi_Informasi")` mengekstrak rentang yang ditunjuk nama tersebut (`$A$2:$A$3`), memunculkan opsi: `Infrastruktur` dan `Pengembangan_Sistem`.

#### Formula Validasi Sel C2 (Tier 3 - Relasi Kompleks)
Untuk dropdown tingkat ketiga, sumber opsi bergantung langsung pada nilai sel `B2`:
```excel
=INDIRECT(SUBSTITUTE($B2, " ", "_"))
```
Jika sel `B2` bernilai `"Pengembangan_Sistem"`, `INDIRECT` mengevaluasi nama tersebut dan mengembalikan daftar: `Backend_Dev`, `Frontend_Dev`, dan `Database_Admin`.

---

## 5. Analisis Kesalahan Umum (*Troubleshooting & Pitfalls*)

Kegagalan implementasi formula `INDIRECT` pada Data Validation umumnya bersumber dari kesalahan format referensi dan konfigurasi regional.

### 1. Galat `#REF!`
* **Penyebab:** Nilai teks di dalam sel referensi tidak terdaftar di dalam *Name Manager*, atau terdapat kesalahan pengetikan karakter tunggal (*typo*).
* **Solusi Teknis:** Buka **Name Manager** (`Ctrl + F3`), verifikasi kecocokan karakter antara isi sel dengan entri nama range. Penamaan *case-insensitive*, tetapi karakter khusus dan spasi harus identik.

### 2. Galat Sintaks Pemisah Argumen (Koma vs Titik Koma)
* **Penyebab:** Perbedaan *Regional Settings* pada sistem operasi Windows.
* **Format US/Internasional:** Menggunakan tanda koma (`,`):
  ```excel
  =INDIRECT(SUBSTITUTE($A2, " ", "_"))
  ```
* **Format Indonesia/Eropa Continental:** Menggunakan tanda titik koma (`;`):
  ```excel
  =INDIRECT(SUBSTITUTE($A2; " "; "_"))
  ```

### 3. Kesalahan Penguncian Sel Dinamis (*Dollar Sign Reference*)
* **Masalah:** Formula disalin ke seluruh kolom menggunakan penguncian absolut penuh (`$A$2`). Akibatnya, baris ke-10 tetap mengambil referensi kategori pada baris ke-2.
* **Perbaikan:** Gunakan format mixed reference: `$A2`. Karakter `$` mengunci kolom A agar tidak bergeser saat divalidasi ke arah horizontal, namun nomor baris `2` dibiarkan relatif agar bergeser saat divalidasi ke arah vertikal (`$A3`, `$A4`, `$A5`).

### 4. Masalah *Invalid Selection Persistence*
* **Kerentanan Desain:** Ketika pengguna memilih Kategori `"Hardware"` lalu Sub-Kategori `"Monitor"`, kemudian pengguna mengubah Kategori induk menjadi `"Software"`, nilai `"Monitor"` pada kolom Sub-Kategori tidak otomatis terhapus secara default oleh Excel. Hal ini menghasilkan anomali data (Kategori: Software, Sub-kategori: Monitor).
* **Mitigasi Proaktif:** Tambahkan conditional formatting untuk mendeteksi anomali. Masukkan formula berikut pada *Conditional Formatting* sel `B2`:
  ```excel
  =ISERROR(MATCH($B2, INDIRECT(SUBSTITUTE($A2, " ", "_")), 0))
  ```
  Tetapkan format visual warna merah jika formula menghasilkan nilai `TRUE` (kondisi di mana nilai sel anak tidak berada dalam domain sel induk).

### 5. Penamaan Terlarang (*Illegal Name Identifier*)
* Excel melarang pendefinisian *Named Range* yang menyerupai referensi sel standar, misalnya `C1`, `R1`, `FY2026`, atau dimulai dengan angka (`1_Kategori`).
* **Solusi Teknis:** Tambahkan prefiks konstan pada seluruh penamaan data master. Jika data awal bernilai `"2026"`, transformasikan menjadi `_2026` atau `Data_2026`. Pada validasi, bungkus formula dengan prefiks:
  ```excel
  =INDIRECT("Data_" & $A2)
  ```

---

## 6. Sifat Volatilitas dan Batasan Skalabilitas

Fungsi `INDIRECT` tergolong sebagai **Volatile Function**. 

### Karakteristik Eksekusi
Formula standar dalam Excel hanya dieksekusi ulang saat sel presedennya (*precedent cell*) mengalami perubahan nilai. Sebaliknya, fungsi volatil dipaksa melakukan kalkulasi ulang (*recalculation*) setiap kali terjadi perubahan pada sel manapun di seluruh *workbook*, termasuk saat eksekusi fungsi auto-save atau refresh tabel.

### Dampak Kinerja
* **Skala Kecil-Menengah (< 5.000 baris validasi):** Tidak ada dampak performa signifikan yang terdeteksi pada CPU modern.
* **Skala Masif Enterprise (> 50.000 baris):** Penggunaan formula `=INDIRECT()` pada puluhan ribu baris validasi secara simultan menyebabkan penurunan performa lembar kerja (*lagging* komputasi).

### Rekomendasi Arsitektural untuk Skala Masif
Jika dataset melampaui puluhan ribu baris dengan tingkat pembaruan transaksi yang tinggi, pertimbangkan migrasi teknik menggunakan:
1. Kombinasi fungsi non-volatil: `INDEX` dan `MATCH` dengan rentang data yang sudah diurutkan (*sorted range*).
2. Pemanfaatan formula *Dynamic Array* Excel 365 (`FILTER`, `UNIQUE`, `XLOOKUP`).
3. Validasi di level database (SQL Constraint) sebelum data masuk ke lingkungan spreadsheet.

---

## 7. Checklist Praktik Terbaik (Best Practices)

Sebelum mendistribusikan spreadsheet berbasis *dependent dropdown* kepada pengguna akhir, jalankan proses audit berikut:

* [ ] **Normalisasi Karakter:** Semua spasi dalam data master kategori telah digantikan dengan karakter garis bawah (`_`).
* [ ] **Validasi Name Manager:** Semua daftar anak terdaftar dengan nama yang persis sama dengan elemen nilai induknya di *Name Manager*.
* [ ] **Penguncian Referensi Tepat:** Formula validasi dependent menggunakan referensi sel semi-absolut (contoh: `$A2`, bukan `$A$2` atau `A2`).
* [ ] **Pengaturan Regional:** Karakter pemisah argumen formula telah disesuaikan dengan konfigurasi sistem (tanda titik koma `;` untuk regional Indonesia, tanda koma `,` untuk regional US).
* [ ] **Proteksi Lembar Kerja:** Tab `Data_Master` disembunyikan (*hidden*) atau dikunci (*protect sheet*) untuk mencegah manipulasi data referensi oleh pengguna non-otoritas.
* [ ] **Audit Nilai Kosong:** Opsi *"Ignore blank"* pada pengaturan Data Validation tetap dicentang untuk mencegah kemunculan galat awal sebelum data induk dipilih.

---

## 8. Ringkasan Eksekutif

Penerapan *Dependent Dropdown List* menggunakan fungsi `INDIRECT` merupakan mekanisme kontrol input data pada level antarmuka Microsoft Excel. Formula mengurai dependensi hierarkis melalui resolusi nama rentang secara dinamis pada saat eksekusi (*run-time resolution*).

Keberhasilan implementasi metode ini bergantung pada konsistensi struktur penamaan rentang pada data master dan manipulasi representasi teks kategori menggunakan fungsi pembantu seperti `SUBSTITUTE`. Penerapan metode ini membatasi masukan pengguna secara ketat, menjaga normalisasi struktur data, dan meminimalkan pekerjaan pembersihan data (*data cleansing*) pada siklus analisis data berikutnya.