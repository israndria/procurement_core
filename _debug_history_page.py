import cloudscraper, re

scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows'})
headers = {'Referer': 'https://spse.inaproc.id/kalselprov/lelang/10116871000/pengumuman'}

r = scraper.get('https://spse.inaproc.id/kalselprov/jadwal/10116871000/history', headers=headers, timeout=15)

# Save full HTML
with open('_history_debug.html', 'w', encoding='utf-8') as f:
    f.write(r.text)

# Find ALL script content (inline)
scripts = re.findall(r'<script[^>]*>(.*?)</script>', r.text, re.DOTALL)
print(f'Inline scripts: {len(scripts)}', flush=True)
for i, s in enumerate(scripts):
    if 'jadwal' in s.lower() or 'history' in s.lower() or 'data' in s.lower() or 'ajax' in s.lower() or 'url' in s.lower():
        print(f'\n--- Script #{i+1} ({len(s)} chars) ---', flush=True)
        print(s[:500], flush=True)
        print('...', flush=True)
