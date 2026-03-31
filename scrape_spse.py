"""
scrape_spse.py — Standalone scraper SPSE → Supabase
Dipanggil oleh GitHub Actions (.github/workflows/scrape-spse.yml)
atau bisa dijalankan manual: python scrape_spse.py

Env vars yang dibutuhkan:
  SUPABASE_URL, SUPABASE_KEY  — dari GitHub Secret atau secret_supabase.env
  SCRAPE_KODE_LPSE            — kode LPSE tertentu, kosong = semua
  SCRAPE_TAHUN                — tahun anggaran, default 2026
  SCRAPE_KATEGORI             — Tender / Non Tender / Pencatatan, default Tender
"""

import urllib.request
import urllib.parse
import http.cookiejar
import json
import re
import logging
import os
import sys
import time
import pandas as pd
from io import StringIO
from datetime import datetime
from dotenv import load_dotenv
from supabase import create_client

# --- Setup ---
_dir = os.path.dirname(os.path.abspath(__file__))
_env = os.path.join(_dir, "secret_supabase.env")
if os.path.exists(_env):
    load_dotenv(_env)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(_dir, "scraper.log"), encoding="utf-8"),
    ]
)
log = logging.getLogger("scrape_spse")

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
if not SUPABASE_URL or not SUPABASE_KEY:
    log.error("SUPABASE_URL / SUPABASE_KEY tidak ditemukan!")
    sys.exit(1)

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- Konfigurasi ---
CONFIG_KATEGORI = {
    "Tender":      {"tabel": "tender",      "endpoint": "lelang"},
    "Non Tender":  {"tabel": "non_tender",  "endpoint": "nontender"},
    "Pencatatan":  {"tabel": "pencatatan",  "endpoint": "pencatatan"},
}

DAFTAR_LPSE = [
    {"nama": "Kabupaten Tapin",        "kode": "tapinkab"},
    {"nama": "Kota Banjarmasin",        "kode": "banjarmasinkota"},
    {"nama": "Kota Banjarbaru",         "kode": "banjarbarukota"},
    {"nama": "Kabupaten Tanah Laut",    "kode": "tanahlautkab"},
    {"nama": "Kabupaten Barito Kuala",  "kode": "baritokualakab"},
    {"nama": "Kabupaten Banjar",        "kode": "banjarkab"},
    {"nama": "Kabupaten Tanah Bumbu",   "kode": "tanahbumbukab"},
    {"nama": "Kabupaten Kotabaru",      "kode": "kotabarukab"},
    {"nama": "Kabupaten HSS",           "kode": "hulusungaiselatankab"},
    {"nama": "Kabupaten HST",           "kode": "hstkab"},
    {"nama": "Kabupaten HSU",           "kode": "hsu"},
    {"nama": "Kabupaten Balangan",      "kode": "balangankab"},
    {"nama": "Kabupaten Tabalong",      "kode": "tabalongkab"},
    {"nama": "Provinsi Kalsel",         "kode": "kalselprov"},
]

BASE_URL = "https://spse.inaproc.id"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"

def strip_html(text):
    """Hapus tag HTML dari string (misal: badge Seleksi Ulang/Gagal)."""
    return re.sub(r"<[^>]+>", "", str(text)).strip()

# --- HTTP helpers ---
def buat_opener():
    cj = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

def fetch_html(opener, url, referer=None, data=None):
    headers = {
        "User-Agent": UA,
        "Referer": referer or url,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0",
    }
    body = None
    if data:
        body = urllib.parse.urlencode(data).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        headers["X-Requested-With"] = "XMLHttpRequest"
        headers["Sec-Fetch-Dest"] = "empty"
        headers["Sec-Fetch-Mode"] = "cors"
    req = urllib.request.Request(url, data=body, headers=headers)
    with opener.open(req, timeout=20) as r:
        return r.read().decode("utf-8", errors="replace")

def get_session(opener, kode_lpse, endpoint):
    url = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    html = fetch_html(opener, url)
    time.sleep(2)  # jeda setelah GET halaman utama
    m = re.search(r"authenticityToken = '([a-f0-9]+)'", html)
    return m.group(1) if m else ""

def get_list_paket(opener, token, kode_lpse, endpoint, tahun):
    url = f"{BASE_URL}/{kode_lpse}/dt/{endpoint}?tahun={tahun}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    semua = []
    start = 0
    PAGE = 100
    MAX_RETRY = 3
    while True:
        data = {
            "draw": str(start // PAGE + 1),
            "start": str(start),
            "length": str(PAGE),
            "authenticityToken": token,
        }
        for attempt in range(MAX_RETRY):
            try:
                raw = fetch_html(opener, url, referer=referer, data=data)
                parsed = json.loads(raw)
                rows = parsed.get("data", [])
                break
            except Exception as e:
                if attempt < MAX_RETRY - 1:
                    log.warning(f"  Retry {attempt+1}/{MAX_RETRY} get_list start={start}: {e}")
                    time.sleep(5 * (attempt + 1))  # backoff: 5s, 10s, 15s
                else:
                    log.error(f"  get_list_paket gagal start={start}: {e}")
                    return semua
        if not rows:
            break
        semua.extend(rows)
        start += PAGE
        time.sleep(1)  # jeda antar halaman
        if len(rows) < PAGE:
            break
    return semua

def parse_tabel1(html):
    try:
        tables = pd.read_html(StringIO(html))
        if len(tables) > 1 and len(tables[1]) > 0:
            row = tables[1].iloc[0]
            return {str(k): str(v) for k, v in row.items()}
    except:
        pass
    return {}

def get_detail_paket(opener, kode_lpse, endpoint, kode_tender):
    base_lelang   = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}"
    base_evaluasi = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    detail = {}

    # Pengumuman
    try:
        html = fetch_html(opener, f"{base_lelang}/pengumumanlelang", referer=referer)
        t = pd.read_html(StringIO(html))[0]
        kv = {}
        for _, row in t.iterrows():
            vals = [str(v).strip() for v in row.values]
            kv[vals[0]] = vals[1]
            if len(vals) > 2 and vals[2] and vals[2] != "nan":
                kv[vals[2]] = vals[3] if len(vals) > 3 else "-"
        detail["jenis_pengadaan"]  = kv.get("Jenis Pengadaan", "-")
        detail["satuan_kerja"]     = kv.get("Satuan Kerja", "-")
        detail["nilai_pagu"]       = kv.get("Nilai Pagu Paket", "0")
        detail["tahun_anggaran"]   = kv.get("Tahun Anggaran", "-")
        detail["nilai_hps"]        = kv.get("Nilai HPS Paket", "0")
        detail["metode_pengadaan"] = kv.get("Metode Pengadaan", "-")
    except Exception as e:
        log.warning(f"  pengumumanlelang {kode_tender}: {e}")

    # Pemenang
    detail["nama_pemenang"] = "Belum Ada Pemenang"
    try:
        d = parse_tabel1(fetch_html(opener, f"{base_evaluasi}/pemenang", referer=referer))
        if d:
            detail["nama_pemenang"] = d.get("Nama Pemenang", "Belum Ada Pemenang")
    except Exception as e:
        log.warning(f"  pemenang {kode_tender}: {e}")

    # Pemenang Berkontrak
    detail["pemenang_berkontrak"] = "Belum Ada Kontrak"
    detail["alamat"]              = "-"
    detail["harga_kontrak"]       = "0"
    try:
        d = parse_tabel1(fetch_html(opener, f"{base_evaluasi}/pemenangberkontrak", referer=referer))
        if d:
            detail["pemenang_berkontrak"] = d.get("Nama Pemenang", "Belum Ada Kontrak")
            detail["alamat"]              = d.get("Alamat", "-")
            detail["harga_kontrak"]       = d.get("Harga Kontrak", "0")
    except Exception as e:
        log.warning(f"  pemenangberkontrak {kode_tender}: {e}")

    # Jadwal kontrak
    detail["kontrak_mulai"]   = "-"
    detail["kontrak_selesai"] = "-"
    try:
        html = fetch_html(opener, f"{base_lelang}/jadwal", referer=referer)
        for _, row in pd.read_html(StringIO(html))[0].iterrows():
            vals = [str(v) for v in row.values]
            if any("Penandatanganan Kontrak" in v for v in vals):
                detail["kontrak_mulai"]   = vals[2] if len(vals) > 2 else "-"
                detail["kontrak_selesai"] = vals[3] if len(vals) > 3 else "-"
                break
    except Exception as e:
        log.warning(f"  jadwal {kode_tender}: {e}")

    return detail

def scrape_satu_lpse(kode_lpse, nama_lpse, endpoint, tabel, tahun):
    log.info(f"=== Mulai: {nama_lpse} ({endpoint}, tahun {tahun}) ===")
    try:
        opener = buat_opener()
        token  = get_session(opener, kode_lpse, endpoint)
        rows   = get_list_paket(opener, token, kode_lpse, endpoint, tahun)
        log.info(f"  {len(rows)} paket ditemukan")
    except Exception as e:
        log.error(f"  Gagal ambil list: {e}")
        return 0, 0

    ok = 0
    err = 0
    for row in rows:
        try:
            kode_tender = str(row[0]).strip()
            nama_paket  = strip_html(row[1])
            instansi    = strip_html(row[2])
            tahapan     = strip_html(row[3])
            if not nama_paket or nama_paket == "nan":
                continue

            detail = get_detail_paket(opener, kode_lpse, endpoint, kode_tender)
            data = {
                "kode_tender":         kode_tender,
                "nama_paket":          nama_paket,
                "instansi":            instansi,
                "tahapan":             tahapan,
                "link_detail":         f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/pengumumanlelang",
                "jenis_pengadaan":     detail.get("jenis_pengadaan", "-"),
                "satuan_kerja":        detail.get("satuan_kerja", "-"),
                "nilai_pagu":          detail.get("nilai_pagu", "0"),
                "tahun_anggaran":      detail.get("tahun_anggaran", "-"),
                "nilai_hps":           detail.get("nilai_hps", "0"),
                "metode_pengadaan":    detail.get("metode_pengadaan", "-"),
                "nama_pemenang":       detail.get("nama_pemenang", "Belum Ada Pemenang"),
                "pemenang_berkontrak": detail.get("pemenang_berkontrak", "Belum Ada Kontrak"),
                "alamat":              detail.get("alamat", "-"),
                "harga_kontrak":       detail.get("harga_kontrak", "0"),
                "kontrak_mulai":       detail.get("kontrak_mulai", "-"),
                "kontrak_selesai":     detail.get("kontrak_selesai", "-"),
                "diambil_pada":        datetime.now().isoformat(),
            }
            supabase.table(tabel).upsert(data).execute()
            ok += 1
            if ok % 10 == 0:
                log.info(f"  {ok} paket diproses...")
        except Exception as e:
            err += 1
            log.warning(f"  Error {row[0] if row else '?'}: {str(e)[:150]}")

    log.info(f"  Selesai: {ok} OK, {err} error")
    return ok, err

# --- Main ---
def main():
    tahun     = os.environ.get("SCRAPE_TAHUN", "2026")
    kategori  = os.environ.get("SCRAPE_KATEGORI", "Tender")
    filter_kode = os.environ.get("SCRAPE_KODE_LPSE", "").strip()

    if kategori not in CONFIG_KATEGORI:
        log.error(f"Kategori tidak valid: {kategori}. Pilih: {list(CONFIG_KATEGORI)}")
        sys.exit(1)

    config   = CONFIG_KATEGORI[kategori]
    endpoint = config["endpoint"]
    tabel    = config["tabel"]

    targets = DAFTAR_LPSE
    if filter_kode:
        targets = [l for l in DAFTAR_LPSE if l["kode"] == filter_kode]
        if not targets:
            log.error(f"Kode LPSE tidak ditemukan: {filter_kode}")
            sys.exit(1)

    log.info(f"Scrape {kategori} tahun {tahun} -> tabel '{tabel}' | {len(targets)} LPSE")

    total_ok = total_err = 0
    for i, lpse in enumerate(targets):
        ok, err = scrape_satu_lpse(lpse["kode"], lpse["nama"], endpoint, tabel, tahun)
        total_ok  += ok
        total_err += err
        if i < len(targets) - 1:
            log.info("  Jeda 5 detik sebelum LPSE berikutnya...")
            time.sleep(5)

    log.info(f"=== SELESAI: total {total_ok} OK, {total_err} error ===")
    if total_ok == 0 and total_err > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
