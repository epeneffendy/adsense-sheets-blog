import os
import re
import json
import yaml
import time
import requests

CONFIG_PATH = os.path.expanduser('~/.hermes/config.yaml')
TARGET_DIR = '/home/effendy/project/adsense-sheets-blog/content/articles'

# Baca kredensial LLM dari konfigurasi lokal Hermes
with open(CONFIG_PATH, 'r') as f:
    cfg = yaml.safe_load(f)

API_KEY = cfg.get('model', {}).get('api_key', '')
BASE_URL = cfg.get('model', {}).get('base_url', '')

# Model reasoning cepat
MODEL_NAME = 'th/deepseek-v4.1-flash:free'

TOPICS = [
    {
        "filename": "panduan-rumus-index-match-excel.md",
        "title": "Panduan Lengkap INDEX MATCH Excel: Pencarian Multi-Kriteria Dua Arah",
        "category": "Excel",
        "description": "Pelajari cara menggabungkan fungsi INDEX dan MATCH di Excel untuk pencarian data horizontal, vertikal, dan dua arah dengan kriteria ganda."
    },
    {
        "filename": "menguasai-sumifs-countifs-excel.md",
        "title": "Menguasai SUMIFS dan COUNTIFS Excel: Agregasi Data dengan Kriteria Majemuk",
        "category": "Excel",
        "description": "Panduan kalkulasi total dan frekuensi data berdasarkan rentang tanggal, status transaksi, dan teks dinamis menggunakan SUMIFS dan COUNTIFS."
    },
    {
        "filename": "panduan-query-function-google-sheets.md",
        "title": "Panduan Lengkap Fungsi QUERY di Google Sheets: Olah Data Layaknya SQL",
        "category": "Google Sheets",
        "description": "Gunakan perintah SELECT, WHERE, GROUP BY, PIVOT, dan ORDER BY langsung di dalam Google Sheets menggunakan fungsi sakti QUERY."
    },
    {
        "filename": "panduan-importrange-google-sheets.md",
        "title": "Cara Menghubungkan Antar Spreadsheet dengan IMPORTRANGE Google Sheets",
        "category": "Google Sheets",
        "description": "Langkah praktis menghubungkan sheet terpisah, membatasi izin akses, dan mengoptimalkan performa sinkronisasi database IMPORTRANGE."
    },
    {
        "filename": "otomasi-google-finance-sheets.md",
        "title": "Otomasi Portofolio Saham dan Kurs Mata Uang Menggunakan GOOGLEFINANCE",
        "category": "Google Sheets",
        "description": "Cara menarik kurs USD/IDR realtime, harga saham IDX, serta riwayat pergerakan aset keuangan langsung ke lembar kerja Google Sheets."
    }
]

def generate_article(topic):
    prompt = f"""
Anda adalah Senior Data Analyst dan Praktisi Spreadsheet profesional.
Tuliskan satu artikel panduan teknis yang sangat komprehensif, edukatif, dan praktis untuk web edukasi 'RumusPro'.

Topik:
- Judul: {topic['title']}
- Kategori: {topic['category']}
- Deskripsi: {topic['description']}

Ketentuan Wajib (Standar Google AdSense & SEO Tingkat Tinggi):
1. Panjang artikel: Wajib minimal 1.000 kata (bahasa Indonesia formal, terstruktur, bebas basa-basi).
2. Format Output: Markdown murni lengkap dengan Frontmatter YAML di bagian paling atas:
---
title: "{topic['title']}"
description: "{topic['description']}"
category: "{topic['category']}"
date: "2026-10-28"
readTime: "8 menit baca"
---
3. Struktur Artikel:
   - Pendahuluan mendalam (Latar belakang masalah, urgensi efisiensi kerja kantor).
   - Anatomi Sintaks Fungsi & Penjelasan Setiap Argumen / Parameter.
   - Tabel Perbandingan atau Kelebihan dibanding metode konvensional.
   - Studi Kasus Nyata 1 (Problem, Struktur Data Kolom A-E, Solusi Formula, Bedah Logika).
   - Studi Kasus Nyata 2 / Skenario Lanjutan (Kasus kompleks, kombinasi formula lain).
   - Pitfall / Kesalahan Umum (#N/A, #VALUE!, kesalahan tanda pemisah koma vs titik koma, referensi relatif/absolut $).
   - Checklist Praktik Terbaik & Ringkasan Eksekutif.
4. Semua blok kode formula wajib dibungkus fenced code block ```excel atau ```javascript (untuk Google Apps Script).
5. Dilarang menggunakan placeholder atau memotong artikel dengan kata '...lanjutkan sendiri'. Tuliskan rumus lengkap dan penjelasan setiap sel.

Tulis sekarang:
"""
    headers = {
        'Authorization': f'Bearer {API_KEY}',
        'Content-Type': 'application/json'
    }
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.4
    }
    
    print(f"Generating: {topic['filename']} ...")
    resp = requests.post(f"{BASE_URL}/chat/completions", headers=headers, json=payload, timeout=120)
    if resp.status_code != 200:
        raise Exception(f"API Error {resp.status_code}: {resp.text}")
    
    content = resp.json()['choices'][0]['message']['content'].strip()
    # Bersihkan pembungkus markdown jika LLM membungkus seluruh respon dalam ```markdown ... ```
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
    for t in TOPICS:
        try:
            generate_article(t)
            time.sleep(1)
        except Exception as e:
            print(f"Error {t['filename']}: {e}")
