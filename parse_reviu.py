"""
parse_reviu.py — Parse Dokumen Tender (Single PDF or Directory) → isi database_reviu + database_dokpil
====================================================================================================
Optimasi untuk Pokja 030:
1. Mendukung scan folder untuk file spesifik (Daftar Peralatan, Personil, Kuantitas, dll).
2. Perbaikan regex dan table extractor untuk struktur tabel yang terpecah (multi-line cells).
3. Penanganan "Resiko Tertinggi" pada RK3K yang lebih presisi.
"""

import sys
import os
import json
import re
from bs4 import BeautifulSoup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─── Load env Supabase ───────────────────────────────────────────────────────
def _load_env():
    env_file = os.path.join(BASE_DIR, "secret_supabase.env")
    if os.path.exists(env_file):
        with open(env_file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

_load_env()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

# ─── Fetch cara_pembayaran dari Supabase ─────────────────────────────────────
def fetch_cara_pembayaran():
    try:
        import httpx
        r = httpx.get(
            f"{SUPABASE_URL}/rest/v1/cara_pembayaran?select=id,keyword,teks",
            headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
            timeout=10
        )
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return [
        {
            "id": "monthly_certificate",
            "keyword": ["monthly", "monthly certificate", "bulanan"],
            "teks": "(Monthly Certificate) Pembayaran dilakukan dengan cara didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara bulanan"
        },
        {
            "id": "termin",
            "keyword": ["termin", "pembayaran termin", "angsuran"],
            "teks": "(Termin) Pembayaran dilakukan dengan cara pembayarannya didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara cara angsuran"
        }
    ]

def deteksi_cara_pembayaran(daftar_cp, bidang=""):
    bidang_lower = (bidang or "").lower().strip()
    KEYWORD_BINAMARGA = ["bina marga", "jalan", "jembatan", "perkerasan", "rigid", "flexible", "hotmix", "aspal", "overlay"]
    is_binamarga = any(kw in bidang_lower for kw in KEYWORD_BINAMARGA)
    id_pilih = "monthly_certificate" if is_binamarga else "termin"
    for cp in daftar_cp:
        if cp["id"] == id_pilih:
            return cp["id"], cp["teks"]
    return daftar_cp[0]["id"], daftar_cp[0]["teks"]

# ─── Helpers ────────────────────────────────────────────────────────────
def bersihkan(s):
    if not s: return ""
    return " ".join(str(s).split()).strip()

def normalisasi_pengalaman(s):
    if not s: return 0
    s = str(s).lower()
    if "-" in s or "0" in s or not any(c.isdigit() for c in s): return 0
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 0

def extract_rekap_from_xlsx(path):
    try:
        import pandas as pd
        xl = pd.ExcelFile(path)
        sheet_name = ""
        for s in xl.sheet_names:
            if any(kw in s.upper() for kw in ["REKAP", "BOQ", "DAFTAR", "KUANTITAS"]):
                sheet_name = s
                break
        if not sheet_name: sheet_name = xl.sheet_names[0]
        
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        items = []
        for index, row in df.iterrows():
            # Gabungkan semua kolom dalam baris ini untuk mencari kata "DIVISI"
            full_row_text = " ".join([str(v) for v in row if v and not pd.isna(v)])
            
            # Cari pola "DIVISI [X] [DESKRIPSI]"
            m_div = re.search(r"(DIVISI\s+[\d\w]+)\s*(.*)", full_row_text, re.IGNORECASE)
            if m_div:
                div_num = m_div.group(1).upper()
                div_desc = bersihkan(m_div.group(2))
                
                # Filter jika ternyata ini header (misal: DIVISI URAIAN)
                if any(kw in div_desc.upper() for kw in ["URAIAN", "HARGA", "BOBOT", "KETERANGAN"]):
                    continue
                
                # Bersihkan sisa-sisa karakter sampah di akhir
                div_desc = re.sub(r"[\d.,\s-]+$", "", div_desc).strip()
                
                full_val = f"{div_num} {div_desc}".strip()
                if len(full_val) > 8 and full_val not in items:
                    items.append(full_val)
            
            # Fallback untuk pola Romawi
            elif re.match(r"^[ \t]*([IVXLC]+\.?)[ \t]+([A-Z][A-Z \/,()&]{5,100})", full_row_text):
                items.append(bersihkan(full_row_text))
                
        return items
    except Exception: return []

def extract_text_from_docx(path):
    try:
        from docx import Document
        doc = Document(path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        return "\n".join(full_text)
    except Exception: return ""

def extract_tables_from_docx(path):
    try:
        from docx import Document
        doc = Document(path)
        tables = []
        for table in doc.tables:
            rows = []
            for row in table.rows:
                rows.append([cell.text for cell in row.cells])
            tables.append(rows)
        return tables
    except Exception: return []

def find_files_by_keywords(directory, keywords):
    """Mencari file dengan prioritas format: xlsx > docx > pdf."""
    found = []
    if not os.path.isdir(directory): return found
    
    all_matching = []
    for f in os.listdir(directory):
        ext = f.lower().split(".")[-1]
        if ext not in ["pdf", "docx", "xlsx", "xlsm"]: continue
        if any(kw.lower() in f.lower() for kw in keywords):
            priority = 0
            if ext in ["xlsx", "xlsm"]: priority = 3
            elif ext == "docx": priority = 2
            elif ext == "pdf": priority = 1
            all_matching.append((priority, os.path.join(directory, f)))
            
    # Sort berdasarkan priority DESC
    all_matching.sort(key=lambda x: x[0], reverse=True)
    return [x[1] for x in all_matching]

# ─── Logic Parsers Per Bagian ───────────────────────────────────────────────

def _parse_fung_bangunan(pdf_or_path, hasil):
    """E2: Fungsi Bangunan"""
    fname = os.path.basename(pdf_or_path if isinstance(pdf_or_path, str) else pdf_or_path.stream.name if hasattr(pdf_or_path, "stream") else "").lower()
    if "gambar" in fname: return False # Skip file gambar/stempel

    if isinstance(pdf_or_path, str):
        if pdf_or_path.lower().endswith(".docx"):
            teks_list = [extract_text_from_docx(pdf_or_path)]
        else: return False
    else:
        teks_list = [p.extract_text() or "" for p in pdf_or_path.pages]

    FORBIDDEN_KEYWORDS = ["digambar oleh", "diperiksa oleh", "no. gambar", "judul gambar", "skala :", "proyek :", "koordinator", "rencana kerja"]
    
    for teks in teks_list:
        # Pola Pekerjaan : [Nama Pekerjaan]
        m_pek = re.search(r"(?:Nama\s+)?Pekerjaan\s*[:/]\s*([^\n]{10,200})", teks, re.IGNORECASE)
        if m_pek:
            val = bersihkan(m_pek.group(1))
            if not any(kw in val.lower() for kw in FORBIDDEN_KEYWORDS) and 10 < len(val) < 180:
                hasil["reviu"]["E2"]["nilai"] = val
                hasil["reviu"]["E2"]["status"] = "terisi"
                return True
                
        # Pola Tujuan / Maksud
        m = re.search(r"(?:Tujuan|Maksud)\s*[:/]?\s*\n*(.*?)(?=\n\s*\d+\.|\Z)", teks, re.DOTALL | re.IGNORECASE)
        if m:
            val = bersihkan(m.group(1))
            if 20 < len(val) < 250:
                hasil["reviu"]["E2"]["nilai"] = val
                hasil["reviu"]["E2"]["status"] = "terisi"
                return True
    return False

def _parse_peralatan(pdf_or_path, hasil):
    """E9-E26: Peralatan"""
    alat_list = []
    if isinstance(pdf_or_path, str):
        if pdf_or_path.lower().endswith(".docx"):
            tables = extract_tables_from_docx(pdf_or_path)
        else: return False
    else:
        tables = []
        for page in pdf_or_path.pages:
            t = page.extract_tables()
            if t: tables.extend(t)

    for table in tables:
        if not table or len(table) < 2: continue
        
        # Deteksi index kolom
        idx_nama = -1; idx_kap = -1; idx_jml = -1
        header = [str(c or "").lower() for c in table[0]]
        for i, c in enumerate(header):
            if any(kw in c for kw in ["jenis", "nama", "peralatan"]): idx_nama = i
            if any(kw in c for kw in ["kapasitas", "ukuran"]): idx_kap = i
            if any(kw in c for kw in ["jumlah", "qty"]): idx_jml = i
        
        if idx_nama == -1: continue # Bukan tabel peralatan
        
        for row in table[1:]:
            if len(row) <= idx_nama: continue
            nama_raw = str(row[idx_nama] or "").strip()
            if not nama_raw or nama_raw.isdigit() or len(nama_raw) < 3: continue
            if any(kw in nama_raw.lower() for kw in ["no", "jenis", "peralatan", "nama"]): continue
            
            # Cek jika ada multiple item dalam satu baris (biasanya terpisah \n di PDF atau baris baru di DOCX)
            namas = [v.strip() for v in nama_raw.split("\n") if len(v.strip()) > 2]
            kaps = [v.strip() for v in str(row[idx_kap] or "").split("\n")] if idx_kap != -1 else [""] * len(namas)
            jmls = [v.strip() for v in str(row[idx_jml] or "").split("\n")] if idx_jml != -1 else ["1"] * len(namas)
            
            for i, n_item in enumerate(namas):
                n = bersihkan(n_item)
                if not n or n.isdigit() or any(kw in n.lower() for kw in ["no", "jenis", "peralatan", "nama"]): continue
                k = bersihkan(kaps[i]) if i < len(kaps) else (bersihkan(" ".join(kaps)) if i == 0 else "")
                j = bersihkan(jmls[i]) if i < len(jmls) else (bersihkan(" ".join(jmls)) if i == 0 else "1 Unit")
                if not any(c.isdigit() for c in j): j = "1 Unit"
                elif "unit" not in j.lower(): j += " Unit"
                alat_list.append({"nama": n, "kapasitas": k, "jumlah": j})
    
    if not alat_list: return False

    # Isi ke hasil
    ALAT_CELLS  = ["E9","E10","E11","E12","E13","E14"]
    KAP_CELLS   = ["E15","E16","E17","E18","E19","E20"]
    JML_CELLS   = ["E21","E22","E23","E24","E25","E26"]
    for i, alat in enumerate(alat_list[:6]):
        hasil["reviu"][ALAT_CELLS[i]]["nilai"] = alat["nama"]
        hasil["reviu"][ALAT_CELLS[i]]["status"] = "terisi"
        hasil["reviu"][KAP_CELLS[i]]["nilai"] = alat["kapasitas"]
        hasil["reviu"][KAP_CELLS[i]]["status"] = "terisi"
        hasil["reviu"][JML_CELLS[i]]["nilai"] = alat["jumlah"]
        hasil["reviu"][JML_CELLS[i]]["status"] = "terisi"
    for i in range(len(alat_list), 6):
        hasil["reviu"][ALAT_CELLS[i]]["status"] = "tidak_ada"
        hasil["reviu"][KAP_CELLS[i]]["status"] = "tidak_ada"
        hasil["reviu"][JML_CELLS[i]]["status"] = "tidak_ada"
    return True

def _parse_personil(pdf_or_path, hasil):
    """E27-E32: Personil"""
    personil_list = []
    if isinstance(pdf_or_path, str):
        if pdf_or_path.lower().endswith(".docx"):
            tables = extract_tables_from_docx(pdf_or_path)
        else: return False
    else:
        tables = []
        for page in pdf_or_path.pages:
            t = page.extract_tables()
            if t: tables.extend(t)

    for table in tables:
        if not table or len(table) < 2: continue
        idx_jab = -1; idx_exp = -1; idx_sert = -1
        header = [str(c or "").lower() for c in table[0]]
        for i, c in enumerate(header):
            if "jabatan" in c: idx_jab = i
            if "pengalaman" in c: idx_exp = i
            if any(kw in c for kw in ["keterangan", "sertifikat", "skk", "skt", "bukti", "syarat"]): idx_sert = i
        
        if idx_jab == -1: continue
        
        for row in table[1:]:
            if len(row) <= idx_jab: continue
            jab_raw = str(row[idx_jab] or "").strip()
            if not jab_raw or len(jab_raw) < 3 or any(kw in jab_raw.lower() for kw in ["no", "jabatan"]): continue
            
            jabs = [j.strip() for j in jab_raw.split("\n") if len(j.strip()) > 3]
            exps_raw = str(row[idx_exp] or "").split("\n") if idx_exp != -1 else ["0"]
            srts_raw = str(row[idx_sert] or "").split("\n") if idx_sert != -1 else [""]
            
            for i, jb in enumerate(jabs):
                ex_val = exps_raw[i] if i < len(exps_raw) else exps_raw[-1]
                ex = normalisasi_pengalaman(ex_val)
                step = len(srts_raw) // len(jabs) if len(jabs) > 0 else 1
                st = " ".join(srts_raw[i*step : (i+1)*step]) if step > 0 else " ".join(srts_raw)
                personil_list.append({"jabatan": bersihkan(jb), "pengalaman": ex, "sertifikat": bersihkan(st)})
    
    if not personil_list: return False

    p_teknis = None; p_k3 = None
    for p in personil_list:
        if any(kw in p["jabatan"].lower() for kw in ["k3", "keselamatan", "hse"]):
            if not p_k3: p_k3 = p
        else:
            if not p_teknis: p_teknis = p
            
    if p_teknis:
        hasil["reviu"]["E27"]["nilai"] = p_teknis["jabatan"]
        hasil["reviu"]["E27"]["status"] = "terisi"
        hasil["reviu"]["E28"]["nilai"] = p_teknis["pengalaman"]
        hasil["reviu"]["E28"]["status"] = "terisi"
        hasil["reviu"]["E29"]["nilai"] = p_teknis["sertifikat"]
        hasil["reviu"]["E29"]["status"] = "terisi"
    if p_k3:
        hasil["reviu"]["E30"]["nilai"] = p_k3["jabatan"]
        hasil["reviu"]["E30"]["status"] = "terisi"
        hasil["reviu"]["E31"]["nilai"] = p_k3["pengalaman"]
        hasil["reviu"]["E31"]["status"] = "terisi"
        hasil["reviu"]["E32"]["nilai"] = p_k3["sertifikat"]
        hasil["reviu"]["E32"]["status"] = "terisi"
    return True

def _parse_rk3k(pdf_or_path, hasil):
    """E33-E34: RK3K"""
    bahaya_list = []
    risiko_tertinggi_bahaya = ""
    risiko_tertinggi_uraian = ""
    
    if isinstance(pdf_or_path, str):
        if pdf_or_path.lower().endswith(".docx"):
            tables = extract_tables_from_docx(pdf_or_path)
            teks_full = extract_text_from_docx(pdf_or_path)
        else: return False
    else:
        tables = []
        for page in pdf_or_path.pages:
            t = page.extract_tables()
            if t: tables.extend(t)
        teks_full = "\n".join([p.extract_text() or "" for p in pdf_or_path.pages])

    for table in tables:
        if not table or len(table) < 2: continue
        idx_uraian = 1; idx_detail = 2; idx_bahaya = 3; idx_risiko = -1
        header = [str(c or "").lower() for c in table[0]]
        for ii, col in enumerate(header):
            if "identifikasi" in col and "bahaya" in col: idx_bahaya = ii
            if "uraian" in col and "pekerjaan" in col: idx_uraian = ii
            if "detail" in col or "item" in col: idx_detail = ii
            if any(kw in col for kw in ["keterangan", "risiko", "tingkat", "level"]): idx_risiko = ii
        if idx_detail == -1 and idx_uraian != -1 and len(header) > idx_uraian + 1: idx_detail = idx_uraian + 1
        
        for row in table[1:]:
            if len(row) > idx_bahaya:
                b_val = bersihkan(row[idx_bahaya])
                if b_val and len(b_val) > 5 and not b_val.isdigit() and "identifikasi" not in b_val.lower():
                    bahaya_list.append(b_val)
                    is_highest = False
                    if idx_risiko != -1 and len(row) > idx_risiko:
                        r_val = str(row[idx_risiko] or "").lower()
                        high_keywords = ["tertinggi", "besar", "sangat besar", "tinggi", "ekstrem", "kritis", "high", "extreme", "major", "fatal", "sangat fatal", "meninggal", "cacat"]
                        if any(kw in r_val for kw in high_keywords): is_highest = True
                    if is_highest:
                        u_val = bersihkan(row[idx_uraian]) if idx_uraian != -1 and len(row) > idx_uraian else ""
                        d_val = bersihkan(row[idx_detail]) if idx_detail != -1 and len(row) > idx_detail else ""
                        risiko_tertinggi_uraian = f"{u_val} ({d_val})" if u_val and d_val else (u_val or d_val)
                        risiko_tertinggi_bahaya = b_val

    if not risiko_tertinggi_bahaya:
        m_rt = re.search(r"(?:Resiko|Tingkat\s+Risiko)\s+Tertinggi\s*[:\s]*([A-Z][^\n]{10,})", teks_full, re.IGNORECASE)
        if m_rt: risiko_tertinggi_bahaya = bersihkan(m_rt.group(1))

    if risiko_tertinggi_uraian:
        hasil["reviu"]["E33"]["nilai"] = risiko_tertinggi_uraian
        hasil["reviu"]["E33"]["status"] = "terisi"
    elif bahaya_list:
        hasil["reviu"]["E33"]["nilai"] = "\n".join(list(dict.fromkeys(bahaya_list))[:5])
        hasil["reviu"]["E33"]["status"] = "terisi"
    if risiko_tertinggi_bahaya:
        hasil["reviu"]["E34"]["nilai"] = risiko_tertinggi_bahaya
        hasil["reviu"]["E34"]["status"] = "terisi"
    return len(bahaya_list) > 0

def _parse_dokpil_uraian(pdf_or_path, hasil):
    """E6-E15 Dokpil: Uraian Pekerjaan (Divisi)"""
    divisi_list = []
    if isinstance(pdf_or_path, str):
        if pdf_or_path.lower().endswith((".xlsx", ".xlsm")):
            divisi_list = extract_rekap_from_xlsx(pdf_or_path)
        elif pdf_or_path.lower().endswith(".docx"):
            teks_list = [extract_text_from_docx(pdf_or_path)]
        else: return False
    else:
        teks_list = [p.extract_text() or "" for p in pdf_or_path.pages]

    if not divisi_list:
        for teks in teks_list:
            matches = re.findall(r"^[ \t]*([IVXLC]+\.?)[ \t]+([A-Z][A-Z \/,()&]{5,100})", teks, re.MULTILINE)
            for m in matches:
                rom, desc = m
                desc = bersihkan(desc)
                if desc and desc not in divisi_list and len(desc) > 3:
                    if any(kw in desc.lower() for kw in ["pekerjaan", "divisi", "umum", "galian", "tanah", "aspal", "beton", "persiapan"]):
                        divisi_list.append(desc)
            if len(divisi_list) >= 2: break
        
    DOKPIL_CELLS = ["E6","E7","E8","E9","E10","E11","E12","E13","E14","E15"]
    for i, div in enumerate(divisi_list[:10]):
        hasil["dokpil"][DOKPIL_CELLS[i]]["nilai"] = div
        hasil["dokpil"][DOKPIL_CELLS[i]]["status"] = "terisi"
    for i in range(len(divisi_list), 10):
        hasil["dokpil"][DOKPIL_CELLS[i]]["status"] = "tidak_ada"
    return len(divisi_list) > 0

# ─── Parser Utama ────────────────────────────────────────────────────────────

def parse_pdf_enhanced(path_or_dir, bidang=""):
    try:
        import pdfplumber
    except ImportError: return None, "pdfplumber tidak tersedia"

    hasil = {
        "input_data": {
            "E16": {"nilai": "", "status": "kosong", "label": "Kegiatan/Sub Kegiatan"},
            "E32": {"nilai": "", "status": "kosong", "label": "Lokasi Pekerjaan"},
            "E33": {"nilai": "", "status": "kosong", "label": "Sumber Dana"},
        },
        "reviu": {
            "E2":  {"nilai": "", "status": "kosong", "label": "Fungsi Bangunan"},
            "E6":  {"nilai": "", "status": "kosong", "label": "SBU KBLI 2020"},
            "E7":  {"nilai": "", "status": "kosong", "label": "SBU KBLI 2015"},
            "E9":  {"nilai": "", "status": "kosong", "label": "Alat 1"},
            "E10": {"nilai": "", "status": "kosong", "label": "Alat 2"},
            "E11": {"nilai": "", "status": "kosong", "label": "Alat 3"},
            "E12": {"nilai": "", "status": "kosong", "label": "Alat 4"},
            "E13": {"nilai": "", "status": "kosong", "label": "Alat 5"},
            "E14": {"nilai": "", "status": "kosong", "label": "Alat 6"},
            "E15": {"nilai": "", "status": "kosong", "label": "Kapasitas 1"},
            "E16": {"nilai": "", "status": "kosong", "label": "Kapasitas 2"},
            "E17": {"nilai": "", "status": "kosong", "label": "Kapasitas 3"},
            "E18": {"nilai": "", "status": "kosong", "label": "Kapasitas 4"},
            "E19": {"nilai": "", "status": "kosong", "label": "Kapasitas 5"},
            "E20": {"nilai": "", "status": "kosong", "label": "Kapasitas 6"},
            "E21": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 1"},
            "E22": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 2"},
            "E23": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 3"},
            "E24": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 4"},
            "E25": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 5"},
            "E26": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 6"},
            "E27": {"nilai": "", "status": "kosong", "label": "Jabatan Teknis"},
            "E28": {"nilai": 0,  "status": "kosong", "label": "Pengalaman Teknis (Tahun)"},
            "E29": {"nilai": "", "status": "kosong", "label": "Nama SKK/SKT"},
            "E30": {"nilai": "", "status": "kosong", "label": "Jabatan K3"},
            "E31": {"nilai": 0,  "status": "kosong", "label": "Pengalaman K3 (Tahun)"},
            "E32": {"nilai": "", "status": "kosong", "label": "Sertifikat K3"},
            "E33": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan RK3"},
            "E34": {"nilai": "", "status": "kosong", "label": "Risiko Tertinggi RK3"},
        },
        "dokpil": {
            "E6":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 1"},
            "E7":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 2"},
            "E8":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 3"},
            "E9":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 4"},
            "E10": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 5"},
            "E11": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 6"},
            "E12": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 7"},
            "E13": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 8"},
            "E14": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 9"},
            "E15": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 10"},
            "E16": {"nilai": "", "status": "kosong", "label": "Cara Pembayaran", "id_cp": ""},
        },
        "error": ""
    }

    base_dir = path_or_dir if os.path.isdir(path_or_dir) else os.path.dirname(path_or_dir)
    file_map = {
        "alat": find_files_by_keywords(base_dir, ["Peralatan", "Alat"]),
        "personil": find_files_by_keywords(base_dir, ["Personil", "Tenaga"]),
        "rk3k": find_files_by_keywords(base_dir, ["RK3K", "Keselamatan"]),
        "dokpil": find_files_by_keywords(base_dir, ["Kuantitas", "BoQ", "RAB"]),
        "uraian": find_files_by_keywords(base_dir, ["Uraian Singkat", "KAK", "Spek"]),
        "merged": [path_or_dir] if os.path.isfile(path_or_dir) else find_files_by_keywords(base_dir, ["Draft_Pokja"])
    }

    # Helper untuk memproses file berdasarkan ekstensi
    def process_file(file_list, func, hasil):
        for f in file_list:
            ext = f.lower().split(".")[-1]
            if ext == "pdf":
                with pdfplumber.open(f) as pdf:
                    if func(pdf, hasil): return True
            elif ext in ["docx", "xlsx", "xlsm"]:
                if func(f, hasil): return True
        return False

    process_file(file_map["uraian"] + file_map["merged"], _parse_fung_bangunan, hasil)
    process_file(file_map["alat"] + file_map["merged"], _parse_peralatan, hasil)
    process_file(file_map["personil"] + file_map["merged"], _parse_personil, hasil)
    process_file(file_map["rk3k"] + file_map["merged"], _parse_rk3k, hasil)
    process_file(file_map["dokpil"] + file_map["merged"], _parse_dokpil_uraian, hasil)

    # 3. Common Fields (Input Data) dari SEMUA file (biasanya ada di RK3K atau KAK)
    all_files = file_map["uraian"] + file_map["rk3k"] + file_map["personil"] + file_map["merged"]
    for f in all_files:
        ext = f.lower().split(".")[-1]
        if ext == "pdf":
            with pdfplumber.open(f) as pdf:
                txt = "\n".join([p.extract_text() or "" for p in pdf.pages[:10]])
        elif ext == "docx":
            txt = extract_text_from_docx(f)
        else: continue
        
        # E16: Kegiatan
        if not hasil["input_data"]["E16"]["nilai"]:
            m_keg = re.search(r"(?:Kegiatan|Sub\s+Kegiatan|Nama\s+Paket)\s*[:/]?\s*([^\n]{5,150})", txt, re.IGNORECASE)
            if m_keg: 
                val = bersihkan(m_keg.group(1))
                if len(val) > 5 and not any(kw in val.lower() for kw in ["tujuan", "latar"]):
                    hasil["input_data"]["E16"]["nilai"] = val
                    hasil["input_data"]["E16"]["status"] = "terisi"
        
        # E32: Lokasi
        if not hasil["input_data"]["E32"]["nilai"]:
            m_lok = re.search(r"Lokasi\s*(?:Pekerjaan)?\s*[:/]?\s*([^\n]{5,100})", txt, re.IGNORECASE)
            if m_lok:
                hasil["input_data"]["E32"]["nilai"] = bersihkan(m_lok.group(1))
                hasil["input_data"]["E32"]["status"] = "terisi"
        
        # E33: Sumber Dana
        if not hasil["input_data"]["E33"]["nilai"]:
            m_sd = re.search(r"(?:Sumber\s+Dana|Sumber\s+Anggaran|Anggaran)\s*[:/]?\s*([^\n]{5,60})", txt, re.IGNORECASE)
            if m_sd:
                hasil["input_data"]["E33"]["nilai"] = bersihkan(m_sd.group(1))
                hasil["input_data"]["E33"]["status"] = "terisi"
            else:
                m_sd2 = re.search(r"(APBD|APBN)\s+(?:Tahun\s+)?(\d{4})", txt, re.IGNORECASE)
                if m_sd2:
                    hasil["input_data"]["E33"]["nilai"] = f"{m_sd2.group(1)} {m_sd2.group(2)}"
                    hasil["input_data"]["E33"]["status"] = "terisi"

    # 4. Cara Pembayaran
    daftar_cp = fetch_cara_pembayaran()
    id_cp, teks_cp = deteksi_cara_pembayaran(daftar_cp, bidang)
    hasil["dokpil"]["E16"]["nilai"] = teks_cp
    hasil["dokpil"]["E16"]["id_cp"] = id_cp
    hasil["dokpil"]["E16"]["status"] = "terisi"

    return hasil, ""

def main():
    bidang = ""
    if len(sys.argv) >= 3 and sys.argv[1] == "--argfile":
        argfile = sys.argv[2]
        with open(argfile, encoding="utf-8") as f:
            lines = [l.rstrip("\n").rstrip("\r") for l in f.readlines()]
        path_pdf = lines[0] if len(lines) > 0 else ""
        folder_output = lines[1] if len(lines) > 1 else ""
        bidang = lines[2] if len(lines) > 2 else ""
    elif len(sys.argv) >= 3:
        path_pdf = sys.argv[1]
        folder_output = sys.argv[2]
        bidang = sys.argv[3] if len(sys.argv) > 3 else ""
    else:
        print("Usage: python parse_reviu.py <path_pdf_or_dir> <folder_output> [bidang]")
        sys.exit(1)

    hasil, err = parse_pdf_enhanced(path_pdf, bidang)
    if err:
        hasil = {"error": err, "reviu": {}, "dokpil": {}, "input_data": {}}

    out_path = os.path.join(folder_output, "_parse_reviu.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(hasil, f, ensure_ascii=False, indent=2)
    print(f"OK: {out_path}")

if __name__ == "__main__":
    main()
