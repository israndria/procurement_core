"""
Import Data LPSE HTML + PDF -> JSON Perantara
=============================================
Python membaca HTML dan PDF, menulis ke file JSON sementara (_import_lpse.json).
VBA yang akan membaca JSON dan mengisi cell Excel nya sendiri (tanpa konflik COM).

Cell mapping:
  HTML -> E3(MAK), E5(Kode Tender), E6(Nama Tender), E8(Kode RUP), E10(Pagu), E11(HPS)
  PDF  -> E12(Nomor Surat Dinas), E13(Nomor PP/BPBJ), E14(Kode Pokja), F17(Nama Dinas)
"""
import sys
import os
import re
import json
import bs4
import ctypes


def read_html_data(html_path):
    """Membaca file HTML dan mengekstrak 6 field data utama menggunakan regex."""
    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        soup = bs4.BeautifulSoup(f, 'html.parser')

    data = {
        "Kode Tender": "",
        "Nama Tender": "",
        "Nilai Pagu": "",
        "Nilai HPS": "",
        "Kode RUP": "",
        "MAK": ""
    }

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

    return data


def read_pdf_data(pdf_path):
    """Membaca file PDF Lembar Disposisi Pokja dan mengekstrak 4 field data."""
    try:
        import pdfplumber
    except ImportError:
        print("pdfplumber tidak terinstall. Install dengan: pip install pdfplumber")
        return {}

    data = {
        "Nomor Surat Dinas": "",   # -> E12  contoh: 000.3.2/756-BUDKAN/DISKAN/2025
        "Nomor PP": "",            # -> E13  contoh: 800.1.11.1/248/075-PP/BPBJ/2025
        "Kode Pokja": "",          # -> E14  contoh: 075
        "Nama Dinas": "",          # -> F17  contoh: DINAS PERIKANAN
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Gabungkan teks dari semua halaman
            full_text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    full_text += t + "\n"
        
        # Regex patterns berdasarkan format Lembar Disposisi Pokja Kabupaten Tapin
        patterns = {
            # "Surat Dari DINAS PERIKANAN Diterima Tanggal"
            "Nama Dinas": r"Surat Dari\s*:?\s*(.+?)\s+Diterima Tanggal",
            # "Nomor Surat : 000.3.2/756-BUDKAN/DISKAN/2025"
            "Nomor Surat Dinas": r"Nomor Surat\s*:\s*([\w\.\-/]+)",
            # "Nomor : 800.1.11.1/248/075-PP/BPBJ/2025"
            "Nomor PP": r"Nomor\s*:\s*([\d\.]+/[\w/\-]+/\d{4})\b",
            # "Diteruskan Kepada Pokja Pemilihan : 075"
            "Kode Pokja": r"Diteruskan Kepada Pokja Pemilihan\s*:\s*(\d+)",
        }
        
        for key, regex in patterns.items():
            match = re.search(regex, full_text, re.IGNORECASE)
            if match:
                data[key] = match.group(1).strip()
                
    except Exception as e:
        print(f"Error membaca PDF: {e}")

    return data


def get_html_from_folder(folder):
    """Mencari file html/htm di folder."""
    for f in os.listdir(folder):
        if f.lower().endswith(('.html', '.htm')):
            return os.path.join(folder, f)
    return None


def get_pdf_from_folder(folder):
    """Mencari file pdf LPSE di folder (mengabaikan hasil cetak sistem)."""
    ignore_prefixes = ("undangan_", "ba_reviu", "ba_pembuktian", "revaluasi_", "isi_reviu", "test_")
    
    pdfs = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
    
    # 1. Prioritaskan file yang memiliki keyword "pokja" di namanya
    for f in pdfs:
        lower_name = f.lower()
        if "pokja" in lower_name and not lower_name.startswith(ignore_prefixes):
            return os.path.join(folder, f)
            
    # 2. Jika tidak ada keyword pokja, ambil PDF apapun yang bukan buatan sistem kita
    for f in pdfs:
        lower_name = f.lower()
        if not lower_name.startswith(ignore_prefixes):
            return os.path.join(folder, f)
            
    return None


def show_error(msg):
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, "Import Data Error", 0x10)
    except:
        print(f"ERROR: {msg}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
        
    excel_path = sys.argv[1]
    folder = os.path.dirname(os.path.abspath(excel_path))
    
    all_data = {}
    
    # === Baca HTML ===
    html_path = get_html_from_folder(folder)
    if html_path:
        print(f"Membaca HTML: {html_path}")
        all_data.update(read_html_data(html_path))
    else:
        print("Tidak ada file HTML ditemukan di folder ini.")
    
    # === Baca PDF ===
    pdf_path = get_pdf_from_folder(folder)
    if pdf_path:
        print(f"Membaca PDF: {pdf_path}")
        all_data.update(read_pdf_data(pdf_path))
    else:
        print("Tidak ada file PDF ditemukan di folder ini.")
    
    # Cek apakah ada data yang berhasil dibaca
    found_count = sum(1 for v in all_data.values() if v)
    if found_count == 0:
        show_error("Tidak ada data yang berhasil dibaca dari HTML maupun PDF di folder ini.")
        sys.exit(1)
    
    # Tulis ke JSON (VBA yang baca dan isi ke cell Excel)
    json_path = os.path.join(folder, "_import_lpse.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False)
    
    print(f"JSON ditulis: {json_path}")
    print(f"Data: {all_data}")
