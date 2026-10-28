---
title: "Ekstraksi Pola Teks Rumit Menggunakan REGEXEXTRACT dan REGEXREPLACE"
description: "Ekstrak nomor telepon, kode pos, email, dan bersihkan teks acak di Google Sheets dengan regular expression berkekuatan tinggi."
category: "Google Sheets"
date: "2026-10-28"
readTime: "8 menit baca"
---

## 1. Pendahuluan

Dataset operasional sering simpan teks tidak terstruktur. Ekspor CRM, tiket helpdesk, form logistik, hasil web scraping campur aduk teks acak. Nomor kontak, email, kode pos, ID transaksi masuk satu sel tanpa pemisah standar.

Pendekatan konvensional pakai kombinasi `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH`, `SUBSTITUTE`. Hasil: formula rapuh. Pola geser satu spasi, formula hasilkan `#VALUE!`. Teks input ubah panjang, parameter angka hardcode rusak. Formula bersarang puluhan baris buat lembar sebar berat, komputasi lambat.

Google Sheets tanam engine regex RE2 Google. Sediakan tiga fungsi inti: `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXMATCH`. Regex definisikan logika parsing berbasis pola, bukan koordinat indeks statis. Otomasi parsing teks hemat waktu olah data, cegah human error, siap eksekusi skala ratusan ribu baris.

---

## 2. Anatomi Sintaks Fungsi

### 2.1 REGEXEXTRACT

Ambil fragmen teks pertama cocok dengan ekspresi reguler.

```excel
=REGEXEXTRACT(teks; ekspresi_reguler)
```

- `teks`: String sumber atau referensi sel tunggal (contoh: `A2`). Wajib tipe teks. Nilai numerik wajib konversi via `TO_TEXT` atau rangkai string kosong.
- `ekspresi_reguler`: Pola regex format string. Wajib apit tanda petik dua. 

Perilaku capturing group:
- Pola tanpa kurung tangkap `()`: Fungsi kembalikan seluruh string cocok pertama.
- Pola pakai satu kurung tangkap `(pola)`: Fungsi kembalikan teks dalam kurung tangkap.
- Pola pakai banyak kurung tangkap `(pola1).*(pola2)`: Hasil tumpah horizontal (*spill array*) ke sel sebelah kanan.

### 2.2 REGEXREPLACE

Ganti substring cocok dengan string baru.

```excel
=REGEXREPLACE(teks; ekspresi_reguler; teks_pengganti)
```

- `teks`: Teks input target modifikasi.
- `ekspresi_reguler`: Pola pencarian substring hapus atau substitusi.
- `teks_pengganti`: Teks baru. Terima teks kosong `""` untuk hapus karakter. Dukung backreference `$1`, `$2` untuk panggil isi capturing group.

### 2.3 Batasan Engine RE2 Google Sheets

Mesin RE2 prioritaskan waktu eksekusi linear. Fitur regex berat dipangkas:
- Tidak dukung *lookaround assertions* penuh. Lookahead `(?=...)` dan lookbehind `(?<=...)` picu error parsing di Sheets. Gunakan capturing group alternatif.
- Tidak dukung backreference di pola pencarian (contoh: `\1` cari duplikat dilarang). Backreference `$1` hanya sah di argumen `teks_pengganti` fungsi `REGEXREPLACE`.
- Karakter *case-sensitive* bawaan. Butuh flag `(?i)` di awal pola jika cari huruf besar-kecil sama.

---

## 3. Perbandingan Metode: Konvensional vs REGEX

| Dimensi | Metode Konvensional (`MID`, `FIND`, `SUBSTITUTE`) | Metode REGEX (`REGEXEXTRACT`, `REGEXREPLACE`) |
| :--- | :--- | :--- |
| **Beban Formula** | Sangat panjang. Rantai 4-6 formula bersarang. | Sangat pendek. 1 fungsi ringkas tangani pola dinamis. |
| **Ketahanan Format** | Gagal saat posisi karakter geser 1 indeks. | Stabil selama pola teks (digit, huruf, simbol) cocok. |
| **Ekstraksi Multi-Kolom** | Buat rumus manual per kolom. | 1 rumus dengan grup `()` otomatis isi kolom samping. |
| **Pembersihan Simbol Acak** | Butuh `SUBSTITUTE` bertumpuk hapus karakter satu per satu. | 1 pola kelas karakter `[^0-9]` sapu bersih non-digit. |
| **Waktu Maintenance** | Sulit audit logika setelah beberapa bulan. | Pola RE2 standar industri, mudah diuji di regex tester. |

---

## 4. Studi Kasus 1: Pembersihan Data Pelanggan Campuran

### 4.1 Data Mentah (Tabel Sumber)

Kolom A simpan catatan CRM mentah. Data campur nama, nomor telepon format liar, email, kode pos.

| Baris | Kolom A (String Mentah CRM) |
| :--- | :--- |
| **1** | `Raw_Data_Customer` |
| **2** | `Budi Santoso - Telp: +62 812-3456-7890, Email: budi.santoso@corp.id, Pos: 10110 Jakarta Pusat` |
| **3** | `SITI AMINAH [08198765432] siti_aminah99@gmail.com - KODEPOS: 40115 (Bandung)` |
| **4** | `Order #9921: Joko Widodo (joko.w@startup.co.id) WA: 0857-1122-3344 - 60241` |
| **5** | `DEWI LESTARI; dewiles@domain.org; 081122334455; POS 50132` |
| **6** | `Rian Hidayat \| rian.h@instansi.go.id \| Pos: 12950 \| No HP: +6281399887766` |

Target parsing:
- Kolom B: Nama Pelanggan Bersih
- Kolom C: Nomor Telepon Format Standar E.164 (`628...`)
- Kolom D: Alamat Email
- Kolom E: Kode Pos 5 Digit

---

### 4.2 Solusi Formula Kolom B: Ekstraksi Nama

Karakter nama selalu muncul sebelum pemisah awal (strip, kurung siku, titik dua, titik koma, pipa).

```excel
=REGEXEXTRACT(A2; "^(?:Order\s#\d+:\s)?([A-Za-z\s]+?)(?:\s[-:;\[|]|$)")
```

Bedah logika pola:
- `^`: Jangkar awal string.
- `(?:Order\s#\d+:\s)?`: Grup non-tangkap opsional. Buang prefix order tiket jika ada.
- `([A-Za-z\s]+?)`: Capturing group utama. Ambil alfabet dan spasi secara lazy. Berhenti segera saat ketemu karakter batas.
- `(?:\s[-:;\[|]|$)`: Grup non-tangkap penanda akhir nama. Cocokkan spasi diikuti simbol delimiter `-`, `:`, `;`, `[`, `|`, atau ujung baris `$`.

---

### 4.3 Solusi Formula Kolom C: Standardisasi Nomor Telepon

Dua tahap: ambil potongan string nomor telepon, lalu buang seluruh karakter non-digit dan samakan prefix ke `62`.

```excel
=REGEXREPLACE(REGEXREPLACE(REGEXEXTRACT(A2; "(?:(?:\+62)|0)[0-9\s\-]{8,16}"); "[^0-9]"; ""); "^0"; "62")
```

Bedah logika bertingkat:
1. `REGEXEXTRACT(A2; "(?:(?:\+62)|0)[0-9\s\-]{8,16}")`:
   - `(?:(?:\+62)|0)`: Deteksi awal nomor format internasional `+62` atau domestik `0`.
   - `[0-9\s\-]{8,16}`: Tangkap digit angka, spasi, atau tanda strip sepanjang 8 hingga 16 karakter.
2. `REGEXREPLACE(...; "[^0-9]"; "")`:
   - `[^0-9]`: Negasi digit. Hapus semua tanda minus, spasi, tanda tambah. Sisa angka murni.
3. `REGEXREPLACE(...; "^0"; "62")`:
   - `^0`: Deteksi digit 0 di awal string. Ganti jadi `62`. Format E.164 tercapai.

---

### 4.4 Solusi Formula Kolom D: Ekstraksi Email

Alamat email miliki format universal: user, `@`, domain, TLD.

```excel
=REGEXEXTRACT(A2; "[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
```

Bedah logika pola:
- `[a-zA-Z0-9._%+\-]+`: Username email. Izinkan karakter alfanumerik, titik, garis bawah, persen, plus, minus minimal 1 karakter.
- `@`: Karakter literal `@`.
- `[a-zA-Z0-9.\-]+`: Domain host.
- `\.`: Titik literal pemisah TLD. Wajib escape pakai *backslash*.
- `[a-zA-Z]{2,}`: Top Level Domain minimal 2 huruf (`id`, `com`, `org`, `co.id`).

---

### 4.5 Solusi Formula Kolom E: Ekstraksi Kode Pos

Kode pos Indonesia miliki panjang tepat 5 digit angka. Hindari tabrakan dengan nomor order atau nomor telepon.

```excel
=REGEXEXTRACT(A2; "\b\d{5}\b")
```

Bedah logika pola:
- `\b`: Word boundary. Batas kata non-alfanumerik. Cegah kecocokan 5 digit di dalam potongan nomor telepon 12 digit.
- `\d{5}`: Tepat 5 digit angka `0-9`.
- `\b`: Word boundary penutup. Pastikan angka berhenti di 5 digit.

---

### 4.6 Hasil Akhir Ekstraksi Data

| Baris | Kolom B (Nama) | Kolom C (Telepon E.164) | Kolom D (Email) | Kolom E (Kode Pos) |
| :--- | :--- | :--- | :--- | :--- |
| **2** | `Budi Santoso` | `6281234567890` | `budi.santoso@corp.id` | `10110` |
| **3** | `SITI AMINAH` | `628198765432` | `siti_aminah99@gmail.com` | `40115` |
| **4** | `Joko Widodo` | `6285711223344` | `joko.w@startup.co.id` | `60241` |
| **5** | `DEWI LESTARI` | `6281122334455` | `dewiles@domain.org` | `50132` |
| **6** | `Rian Hidayat` | `6281399887766` | `rian.h@instansi.go.id` | `12950` |

---

## 5. Studi Kasus 2: Skenario Kompleks & Pola Dinamis

### 5.1 Ekstraksi Log API Masuk ke Banyak Kolom Sekaligus

Log sistem masuk ke sel tunggal `A2`:

```text
TIMESTAMP:[2026-10-28 14:22:01] | STATUS:[SUCCESS] | TXID:[TXN-98412] | AMOUNT:[IDR 1,500,000] | GATEWAY:[XENDIT]
```

Kebutuhan: Pecah 5 parameter log ke sel `B2:F2` tanpa tulis 5 formula terpisah. Gunakan sifat array horizontal `REGEXEXTRACT` dengan capturing group berantai:

```excel
=REGEXEXTRACT(A2; "^TIMESTAMP:\[(.*?)\]\s\|\sSTATUS:\[(.*?)\]\s\|\sTXID:\[(.*?)\]\s\|\sAMOUNT:\[(?:IDR\s)?(.*?)\]\s\|\sGATEWAY:\[(.*?)\]$")
```

Bedah cara kerja:
- Setiap pasang kurung `(.*?)` definisikan 1 capturing group.
- Ada 5 grup tangkap: `TIMESTAMP`, `STATUS`, `TXID`, `AMOUNT`, `GATEWAY`.
- Google Sheets otomatis distribusikan hasil grup 1 ke `B2`, grup 2 ke `C2`, grup 3 ke `D2`, grup 4 ke `E2`, grup 5 ke `F2`.
- `(?:IDR\s)?`: Non-capturing group. Buang label mata uang agar kolom `AMOUNT` hanya berisi nominal angka bersih.

---

### 5.2 Dynamic Array Formula: Eksekusi Otomatis Seluruh Baris

Jangan salin rumus manual dari baris 2 ke baris 1000. Gabungkan `REGEXREPLACE` dengan `ARRAYFORMULA`.

Normalisasi nomor telepon kolom `A2:A` otomatis:

```excel
=ARRAYFORMULA(
  IF(A2:A = ""; "";
    REGEXREPLACE(
      REGEXREPLACE(
        REGEXREPLACE(A2:A; "[^0-9]"; "");
        "^62"; "0"
      );
      "^0"; "62"
    )
  )
)
```

Alur logika:
1. `IF(A2:A = ""; ""; ...)`: Cek baris kosong. Lewati komputasi jika sel kosong. Cegah lembar sebar lempar error.
2. `REGEXREPLACE(A2:A; "[^0-9]"; "")`: Bersihkan semua karakter selain angka untuk seluruh baris aktif.
3. `REGEXREPLACE(...; "^62"; "0")`: Normalisasi sementara nomor awalan `62` jadi `0`.
4. `REGEXREPLACE(...; "^0"; "62")`: Pastikan seluruh nomor berawalan `62` seragam.

---

### 5.3 Modifikasi Struktur Teks dengan Backreference `$1`, `$2`

Skenario: Ubah format tanggal input log `YYYY-MM-DD` jadi format lokal `DD/MM/YYYY` langsung di dalam string kalimat.

Contoh string sel `A2`:
`Laporan audit selesai tanggal 2026-10-28 oleh tim QC.`

Formula:

```excel
=REGEXREPLACE(A2; "(\d{4})-(\d{2})-(\d{2})"; "$3/$2/$1")
```

Bedah logika backreference:
- `(\d{4})`: Grup $1 (Tahun = 2026).
- `(\d{2})`: Grup $2 (Bulan = 10).
- `(\d{2})`: Grup $3 (Hari = 28).
- `"$3/$2/$1"`: Susun ulang posisi nilai tangkapan. Hasil teks: `Laporan audit selesai tanggal 28/10/2026 oleh tim QC.`

---

## 6. Pitfall & Kesalahan Teknis

### 6.1 Error `#N/A` (Pola Tidak Ditemukan)
Penyebab: `REGEXEXTRACT` lempar error `#N/A` saat teks target tidak miliki substring cocok dengan ekspresi reguler.
Solusi: Amankan dengan pembungkus `IFERROR`. Berikan string kosong atau nilai default.

```excel
=IFERROR(REGEXEXTRACT(A2; "\b\d{5}\b"); "KODE_POS_TIDAK_VALID")
```

### 6.2 Error `#VALUE!` (Ketidakcocokan Tipe Data)
Penyebab: Sel referensi berisi data numerik murni, nilai tanggal, atau boolean. Fungsi REGEX di Google Sheets tolak input non-teks.
Solusi: Paksa konversi tipe data jadi string via fungsi `TO_TEXT` atau operator ampersand `&`.

```excel
=REGEXEXTRACT(TO_TEXT(A2); "\d+")
```
atau:
```excel
=REGEXEXTRACT(A2 & ""; "\d+")
```

### 6.3 Konflik Separator Regional (Koma `,` vs Titik Koma `;`)
Penyebab: Setelan lokal (*Locale*) Google Sheets pengaruhi pemisah argumen fungsi.
- Locale Amerika Serikat / Internasional: Pemisah fungsi pakai koma `,`.
- Locale Indonesia: Pemisah fungsi pakai titik koma `;`, pemisah desimal pakai koma `,`.

Jika formula di tutorial tidak jalan, lakukan penyesuaian:
```excel
=REGEXEXTRACT(A2, "\d+")   <- Sintaks US
=REGEXEXTRACT(A2; "\d+")   <- Sintaks ID
```

### 6.4 Lupa Escape Karakter Reserved Regex
Penyebab: Karakter regex khusus dipakai langsung tanpa tanda escape `\`. Simbol bentrok: `. * + ? ^ $ { } ( ) [ ] | \`.
Contoh kasus salah (ingin cari domain `.com`):
`=REGEXEXTRACT(A2; ".com")`
Arti salah: Titik `.` cocokkan karakter acak apa saja sebelum `com` (contoh `xcom`, `1com` ikut cocok).

Solusi benar: Tambahkan escape `\.`.
```excel
=REGEXEXTRACT(A2; "\.com")
```

### 6.5 Jebakan Greedy `.*` vs Lazy `.*?`
Pola rakus (*greedy*) `.*` ambil karakter sebanyak mungkin sampai batas akhir ditemukan. Pola malas (*lazy*) `.*?` berhenti pada batas terdekat.

Contoh teks:
`<b>Data 1</b> dan <b>Data 2</b>`

Uji Ekstraksi:
- Pola Greedy: `=REGEXEXTRACT(A2; "<b>(.*)</b>")`
  Hasil: `Data 1</b> dan <b>Data 2` (Salah, ambil dari tag pembuka pertama sampai tag penutup terakhir).
- Pola Lazy: `=REGEXEXTRACT(A2; "<b>(.*?)</b>")`
  Hasil: `Data 1` (Benar, berhenti di tag penutup pertama).

---

## 7. Solusi Cadangan: Google Apps Script (Custom UDF)

Saat pemrosesan data libatkan ratusan ribu baris, formula kompleks lembar sebar picu batas memori Google Sheets. Gunakan Apps Script via eksekusi JavaScript engine V8 untuk komputasi background.

Buka menu **Extensions > Apps Script**, masukkan skrip:

```javascript
/**
 * Ekstrak data regex secara cepat untuk rentang array besar.
 * @param {string[][]} data Range sel input teks.
 * @param {string} regexPattern Pola regular expression.
 * @param {number} groupIndex Indeks group capture (0 untuk full match).
 * @return {string[][]} Output array satu dimensi.
 * @customfunction
 */
function REGEX_PARSER_FAST(data, regexPattern, groupIndex) {
  const targetGroup = (typeof groupIndex === "number") ? groupIndex : 0;
  const regex = new RegExp(regexPattern);
  
  return data.map(row => {
    const cellValue = String(row[0] || "");
    if (!cellValue) return [""];
    
    const match = cellValue.match(regex);
    if (!match) return [""];
    
    return [match[targetGroup] !== undefined ? match[targetGroup] : ""];
  });
}
```

Panggil formula kustom di sel lembar sebar:

```excel
=REGEX_PARSER_FAST(A2:A1000; "([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})"; 1)
```

Kelebihan Apps Script: Evaluasi range besar dalam satu payload memori internal array, hindari limit kalkulasi formula sel standar Google Sheets.

---

## 8. Checklist Praktik Terbaik & Ringkasan

### 8.1 Checklist Pra-Deploy Formula
- [ ] Validasi pola regex di tester pihak ketiga (pastikan setelan mesin RE2 aktif).
- [ ] Pastikan tidak ada sintaks *lookahead* / *lookbehind* yang ditolak engine Google Sheets.
- [ ] Periksa titik koma `;` atau koma `,` sesuai locale lembar sebar yang aktif.
- [ ] Bungkus formula ekstrak pakai `IFERROR` untuk tangani string tidak cocok.
- [ ] Pakai `TO_TEXT` jika sel input berpotensi berisi angka murni.
- [ ] Ganti quantifier `.*` jadi `.*?` saat parsing teks berulang di baris yang sama.
- [ ] Gunakan `ARRAYFORMULA` di sel header atas untuk otomatisasi baris baru, hindari copy-paste formula manual.

### 8.2 Ringkasan Sintaks Acuan Cepat

| Kebutuhan Olah Teks | Rumus Acuan Google Sheets |
| :--- | :--- |
| **Ekstrak Angka Saja** | `=REGEXEXTRACT(A2; "\d+")` |
| **Hapus Semua Simbol / Tanda Baca** | `=REGEXREPLACE(A2; "[^a-zA-Z0-9\s]"; "")` |
| **Ekstrak Huruf Saja** | `=REGEXEXTRACT(A2; "[A-Za-z]+")` |
| **Pangkas Spasi Ganda Jadi Spasi Tunggal** | `=REGEXREPLACE(A2; "\s+"; " ")` |
| **Ambil Teks di Dalam Tanda Kurung** | `=REGEXEXTRACT(A2; "\((.*?)\)")` |
| **Cek Keberadaan Pola (True/False)** | `=REGEXMATCH(A2; "^[0-9]+$")` |