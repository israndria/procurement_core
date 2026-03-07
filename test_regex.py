import bs4
import re
import json

html_path = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\LPSE - (LPSE) Pengumuman Delegasi Pokja.html"
with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
    soup = bs4.BeautifulSoup(f, 'html.parser')

# Get text blocks separated by newline or space
raw_text = soup.get_text(separator=' \n ', strip=True)

data = {
    "Kode Tender": "",
    "Nama Tender": "",
    "Nilai Pagu": "",
    "Nilai HPS": "",
    "Kode RUP": "",
    "MAK": ""
}

# The page layout might be literally just "MAK:123 Kode Tender:123"
# Since they are strong/span/td, get_text(' ') will result in "MAK : 123 Kode Tender : 123"
text = soup.get_text(separator=' ', strip=True)

patterns = {
    "MAK": r"MAK\s*:\s*([\d\.]+)",
    "Kode Tender": r"Kode Tender\s*:\s*(\d+)",
    "Nama Tender": r"Nama Tender\s*:\s*(.+?)Kode RUP",
    "Kode RUP": r"Kode RUP\s*:\s*(\d+)",
    "Nilai Pagu": r"Nilai Pagu\s*:\s*(Rp\.\s*[\d\.,]+)",
    "Nilai HPS": r"Nilai HPS\s*:\s*(Rp\.\s*[\d\.,]+)",
}

for key, regex in patterns.items():
    match = re.search(regex, text, re.IGNORECASE)
    if match:
        data[key] = match.group(1).strip()

print(json.dumps(data, indent=2))
