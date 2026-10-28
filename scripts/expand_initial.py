import os
import re
import json
import yaml
import requests

CONFIG_PATH = os.path.expanduser('~/.hermes/config.yaml')
TARGET_DIR = '/home/effendy/project/adsense-sheets-blog/content/articles'

with open(CONFIG_PATH, 'r') as f:
    cfg = yaml.safe_load(f)

API_KEY = cfg.get('model', {}).get('api_key', '')
BASE_URL = cfg.get('model', {}).get('base_url', '')

TOPICS = [
    {
        "filename": "panduan-rumus-xlookup-excel.md",
        "title": "Panduan Lengkap Rumus XLOOKUP Excel: Pengganti VLOOKUP dan HLOOKUP",
        "category": "Excel",
        "description": "Pelajari cara kerja fungsi XLOOKUP di Microsoft Excel, keunggulan dibanding VLOOKUP tradisional, dan cara pencarian horizontal serta backward lookup tanpa ribet."
    },
    {
        "filename": "menguasai-arrayformula-google-sheets.md",
        "title": "Menguasai ARRAYFORMULA Google Sheets untuk Otomasi Kalkulasi Otomatis",
        "category": "Google Sheets",
        "description": "Panduan praktis menggunakan ARRAYFORMULA di Google Sheets: kalkulasi ratusan baris secara otomatis dari 1 sel tanpa perlu drag-and-drop manual."
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

for topic in TOPICS:
    prompt = f"""
Tuliskan artikel mendalam standar industri untuk 'RumusPro' dengan panjang minimal 1.500 kata.
Judul: {topic['title']}
Kategori: {topic['category']}
Deskripsi: {topic['description']}

Wajib:
1. Markdown murni diawali frontmatter YAML (title, description, category, date, readTime).
2. Struktur: Pendahuluan latar belakang, Anatomi sintaks fungsi per parameter, Tabel kelebihan vs formula lama, Studi Kasus Nyata 1 (problem, dataset, formula, penjelasan detail), Studi Kasus Nyata 2 (kondisi ganda / reverse lookup), Kesalahan umum (#N/A, spill error #SPILL!), dan Best practice checklist.
3. Code block fenced ```excel.
"""
    print(f"Re-generating {topic['filename']} ...")
    content = stream_chat_completion(prompt).strip()
    if content.startswith("```markdown"):
        content = re.sub(r"^```markdown\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    elif content.startswith("```"):
        content = re.sub(r"^```[a-zA-Z]*\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        
    with open(os.path.join(TARGET_DIR, topic['filename']), 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Selesai: {topic['filename']} ({len(content.split())} kata)")
