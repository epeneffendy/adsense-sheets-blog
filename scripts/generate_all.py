import os
import re
import json
import yaml
import time
import requests

CONFIG_PATH = os.path.expanduser('~/.hermes/config.yaml')
TARGET_DIR = '/home/effendy/project/adsense-sheets-blog/content/articles'

with open(CONFIG_PATH, 'r') as f:
    cfg = yaml.safe_load(f)

API_KEY = cfg.get('model', {}).get('api_key', '')
BASE_URL = cfg.get('model', {}).get('base_url', '')

# 20 Topik Pelengkap untuk mencapai target 27 artikel > 1.000 kata
ALL_TOPICS = [
    # --- EXCEL BATCH (10 Topik) ---
    {
        "filename": "panduan-power-query-excel-pemula.md",
        "title": "Panduan Praktis Power Query Excel untuk Otomasi Pembersihan Data",
        "category": "Excel",
        "description": "Cara menggunakan Power Query untuk menggabungkan banyak file, unpivot kolom, dan membersihkan data kotor tanpa makro VBA."
    },
    {
        "filename": "rumus-textsplit-textjoin-excel.md",
        "title": "Membedah Rumus Manipulasi Teks Modern Excel: TEXTSPLIT dan TEXTJOIN",
        "category": "Excel",
        "description": "Pisahkan dan gabungkan teks multi-karakter dengan pemisah dinamis di Excel menggunakan TEXTSPLIT, TEXTJOIN, dan TEXTBEFORE."
    },
    {
        "filename": "cara-membuat-drop-down-bertingkat-excel.md",
        "title": "Cara Membuat Dropdown Bertingkat di Excel dengan INDIRECT",
        "category": "Excel",
        "description": "Langkah menyusun dependent drop-down list (Kategori > Sub-Kategori) di Excel menggunakan Data Validation dan formula INDIRECT."
    },
    {
        "filename": "panduan-pivot-table-grouping-excel.md",
        "title": "Teknik Grouping Tanggal dan Angka pada Pivot Table Excel",
        "category": "Excel",
        "description": "Kelompokkan transaksi bulanan, kuartalan, dan interval nominal di Pivot Table Excel secara instan untuk laporan manajemen."
    },
    {
        "filename": "rumus-datedif-hitung-umur-kerja-excel.md",
        "title": "Menghitung Masa Kerja dan Selisih Tanggal dengan Rumus DATEDIF Excel",
        "category": "Excel",
        "description": "Panduan menghitung tahun, bulan, dan hari masa kerja karyawan secara akurat menggunakan fungsi tersembunyi DATEDIF."
    },
    {
        "filename": "panduan-vstack-hstack-excel.md",
        "title": "Menggabungkan Banyak Tabel Secara Dinamis dengan VSTACK dan HSTACK",
        "category": "Excel",
        "description": "Satukan laporan cabang secara vertikal dan horizontal tanpa copy-paste manual menggunakan array dinamis VSTACK dan HSTACK."
    },
    {
        "filename": "conditional-formatting-dengan-formula-excel.md",
        "title": "Penerapan Conditional Formatting Berbasis Formula Kustom di Excel",
        "category": "Excel",
        "description": "Warnai baris otomatis berdasarkan tanggal jatuh tempo, status persetujuan, dan nilai di atas rata-rata dengan aturan formula kustom."
    },
    {
        "filename": "rekonsiliasi-bank-otomatis-excel.md",
        "title": "Teknik Rekonsiliasi Bank Otomatis di Excel untuk Tim Akuntansi",
        "category": "Excel",
        "description": "Template dan logika formula pencocokan mutasi bank dan buku besar kas menggunakan perpaduan SUMIFS, COUNTIF, dan conditional rules."
    },
    {
        "filename": "rumus-depresiasi-aktiva-tetap-excel.md",
        "title": "Menghitung Penyusutan Aset Tetap di Excel: Metode Garis Lurus (SLN) dan Saldo Menurun (DB)",
        "category": "Excel",
        "description": "Implementasi formula finansial SLN, DB, dan DDB untuk menghitung jadwal amortisasi dan depresiasi aktiva tetap perusahaan."
    },
    {
        "filename": "audit-formula-dan-trace-precedents-excel.md",
        "title": "Cara Melacak dan Mengaudit Error Formula Rumit di Excel",
        "category": "Excel",
        "description": "Gunakan fitur Evaluate Formula, Trace Precedents, dan Trace Dependents untuk mendeteksi akar penyebab #VALUE!, #REF!, dan circular reference."
    },

    # --- GOOGLE SHEETS BATCH (10 Topik) ---
    {
        "filename": "panduan-regexextract-regexreplace-sheets.md",
        "title": "Ekstraksi Pola Teks Rumit Menggunakan REGEXEXTRACT dan REGEXREPLACE",
        "category": "Google Sheets",
        "description": "Ekstrak nomor telepon, kode pos, email, dan bersihkan teks acak di Google Sheets dengan regular expression berkekuatan tinggi."
    },
    {
        "filename": "panduan-filter-sortn-google-sheets.md",
        "title": "Menampilkan Data Dinamis dengan Fungsi FILTER dan SORTN Google Sheets",
        "category": "Google Sheets",
        "description": "Filter data multi-kriteria dan tampilkan ranking Top 5 atau Top 10 transaksi secara otomatis tanpa perlu mengubah tabel sumber."
    },
    {
        "filename": "membuat-grafik-mini-sparkline-sheets.md",
        "title": "Membuat Mini Dashboard Visual di Sel dengan SPARKLINE Google Sheets",
        "category": "Google Sheets",
        "description": "Desain tren grafik garis (line chart), grafik batang (bar), dan win-loss visual langsung di dalam satu sel menggunakan fungsi SPARKLINE."
    },
    {
        "filename": "otomasi-google-apps-script-kirim-email.md",
        "title": "Otomasi Pengiriman Email Pengingat dari Google Sheets via Apps Script",
        "category": "Google Sheets",
        "description": "Tutorial menulis kode JavaScript Google Apps Script (GAS) untuk mengirim notifikasi tagihan atau peringatan otomatis via Gmail."
    },
    {
        "filename": "proteksi-sel-dan-manajemen-akses-sheets.md",
        "title": "Cara Mengunci Sel dan Mengatur Izin Akses Kolom di Google Sheets",
        "category": "Google Sheets",
        "description": "Lindungi formula penting dari ketidaksengajaan edit tim kerja dengan Protected Sheets and Ranges serta permission bertingkat."
    },
    {
        "filename": "otomasi-checkbox-checklist-sheets.md",
        "title": "Otomasi Progress Bar dan Perhitungan Status dengan Checkbox di Sheets",
        "category": "Google Sheets",
        "description": "Buat tracker to-do list interaktif, strike-through teks otomatis, dan progress bar persentase selesai menggunakan kotak centang."
    },
    {
        "filename": "panduan-unique-transpose-google-sheets.md",
        "title": "Ekstraksi Nilai Unik dan Rotasi Matriks dengan UNIQUE dan TRANSPOSE",
        "category": "Google Sheets",
        "description": "Hapus duplikasi data secara real-time dan balik orientasi baris menjadi kolom secara dinamis di Google Sheets."
    },
    {
        "filename": "sinkronisasi-google-form-ke-sheets.md",
        "title": "Manajemen Database Respon Google Form di Google Sheets Secara Efisien",
        "category": "Google Sheets",
        "description": "Aturan arsitektur memisahkan data mentah form response dengan sheet analisis kalkulasi untuk mencegah corrupt data entri."
    },
    {
        "filename": "trik-shortcut-keyboard-google-sheets.md",
        "title": "Kumpulan Shortcut Keyboard Google Sheets untuk Melipatgandakan Kecepatan Kerja",
        "category": "Google Sheets",
        "description": "Daftar tombol pintas esensial navigasi cepat, insert baris, seleksi data masif, dan format angka tanpa perlu menyentuh mouse."
    },
    {
        "filename": "perbandingan-excel-vs-google-sheets.md",
        "title": "Excel vs Google Sheets: Analisis Kebutuhan Bisnis, Kolaborasi, dan Kapasitas Data",
        "category": "Spreadsheet",
        "description": "Perbandingan mendalam kapabilitas pemrosesan baris, fitur kolaborasi live, integrasi cloud, dan keamanan data antara Excel dan Sheets."
    }
]

def stream_chat_completion(prompt):
    headers = {
        'Authorization': f'Bearer {API_KEY}',
        'Content-Type': 'application/json'
    }
    payload = {
        "model": "Kama-Agent",
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3
    }
    
    resp = requests.post(f"{BASE_URL}/chat/completions", headers=headers, json=payload, stream=True, timeout=180)
    if resp.status_code != 200:
        raise Exception(f"API Error {resp.status_code}: {resp.text}")
        
    full_text = []
    for line in resp.iter_lines():
        if not line:
            continue
        line_str = line.decode('utf-8')
        if line_str.startswith('data: '):
            data_content = line_str[6:].strip()
            if data_content == '[DONE]':
                break
            try:
                chunk = json.loads(data_content)
                delta = chunk.get('choices', [{}])[0].get('delta', {})
                content = delta.get('content', '')
                if content:
                    full_text.append(content)
            except Exception:
                pass
                
    return "".join(full_text)

def generate_article(topic):
    prompt = f"""
Anda adalah Senior Data Analyst dan Praktisi Spreadsheet profesional.
Tuliskan satu artikel panduan teknis yang sangat komprehensif, edukatif, dan praktis untuk web edukasi 'RumusPro'.

Topik:
- Judul: {topic['title']}
- Kategori: {topic['category']}
- Deskripsi: {topic['description']}

Ketentuan Wajib (Standar Google AdSense & SEO Tingkat Tinggi):
1. Panjang artikel: Wajib minimal 1.200 kata (bahasa Indonesia formal, terstruktur, mendalam, bebas basa-basi).
2. Format Output: Markdown murni lengkap dengan Frontmatter YAML di bagian paling atas:
---
title: "{topic['title']}"
description: "{topic['description']}"
category: "{topic['category']}"
date: "2026-10-28"
readTime: "8 menit baca"
---
3. Struktur Artikel:
   - Pendahuluan mendalam (Latar belakang masalah, skenario dunia kerja nyata, urgensi otomasi).
   - Anatomi Sintaks Fungsi & Penjelasan Setiap Argumen / Parameter teknis.
   - Tabel Perbandingan atau Kelebihan dibanding metode konvensional.
   - Studi Kasus Nyata 1 (Problem spesifik, Contoh Data Kolom A-E, Solusi Formula, Bedah Logika).
   - Studi Kasus Nyata 2 / Skenario Lanjutan (Kasus kompleks, kombinasi formula majemuk).
   - Pitfall / Kesalahan Umum (#N/A, #VALUE!, tanda pemisah koma vs titik koma, referensi absolut $).
   - Checklist Praktik Terbaik & Ringkasan Eksekutif.
4. Semua blok kode formula wajib dibungkus fenced code block ```excel atau ```javascript (untuk Google Apps Script).
5. Dilarang menggunakan placeholder atau memotong artikel dengan kata '...lanjutkan sendiri'. Tuliskan rumus lengkap dan penjelasan setiap sel.

Tulis sekarang:
"""
    print(f"Generating {topic['filename']} ...")
    content = stream_chat_completion(prompt).strip()
    
    if content.startswith("```markdown"):
        content = re.sub(r"^```markdown\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    elif content.startswith("```"):
        content = re.sub(r"^```[a-zA-Z]*\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        
    target_path = os.path.join(TARGET_DIR, topic['filename'])
    with open(target_path, 'w', encoding='utf-8') as f:
        f.write(content)
        
    word_count = len(content.split())
    print(f"Selesai: {topic['filename']} ({word_count} kata)")

if __name__ == '__main__':
    for t in ALL_TOPICS:
        target_path = os.path.join(TARGET_DIR, t['filename'])
        if os.path.exists(target_path) and os.path.getsize(target_path) > 1000:
            print(f"Skip {t['filename']} (sudah ada).")
            continue
        try:
            generate_article(t)
            time.sleep(2)
        except Exception as e:
            print(f"Error {t['filename']}: {e}")
