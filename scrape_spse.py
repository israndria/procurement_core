"""
scrape_spse.py — Standalone scraper SPSE → Supabase
Dipanggil oleh Task Scheduler atau bisa dijalankan manual: python scrape_spse.py

Env vars yang dibutuhkan:
  SUPABASE_URL, SUPABASE_KEY  — dari secret_supabase.env
  SCRAPE_KODE_LPSE            — kode LPSE tertentu, kosong = semua
  SCRAPE_TAHUN                — tahun anggaran, default 2026
  SCRAPE_KATEGORI             — Tender / Non Tender / Pencatatan, default Tender
"""

import json
import re
import logging
import os
import sys
import time
import requests
import pandas as pd
import httpx
from io import StringIO
from datetime import datetime
from dotenv import load_dotenv
from supabase import create_client
from supabase.lib.client_options import SyncClientOptions

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

# --- Konfigurasi ---
CONFIG_KATEGORI = {
    "Tender":      {"tabel": "tender",     "endpoint": "lelang",     "endpoint_dt": "lelang", "suffix": "pengumumanlelang"},
    "Non Tender":  {"tabel": "non_tender", "endpoint": "nontender",  "endpoint_dt": "pl",     "suffix": "pengumumanpl"},
    "Pencatatan":  {"tabel": "pencatatan", "endpoint": "pencatatan", "endpoint_dt": "nonspk", "suffix": "pengumuman"},
}

DAFTAR_LPSE = [
    {"nama": "Kabupaten Tapin",        "kode": "tapinkab"},
    {"nama": "Kota Banjarmasin",       "kode": "banjarmasinkota"},
    {"nama": "Kota Banjarbaru",        "kode": "banjarbarukota"},
    {"nama": "Kabupaten Tanah Laut",   "kode": "tanahlautkab"},
    {"nama": "Kabupaten Barito Kuala", "kode": "baritokualakab"},
    {"nama": "Kabupaten Banjar",       "kode": "banjarkab"},
    {"nama": "Kabupaten Tanah Bumbu",  "kode": "tanahbumbukab"},
    {"nama": "Kabupaten Kotabaru",     "kode": "kotabarukab"},
    {"nama": "Kabupaten HSS",          "kode": "hulusungaiselatankab"},
    {"nama": "Kabupaten HST",          "kode": "hstkab"},
    {"nama": "Kabupaten HSU",          "kode": "hsu"},
    {"nama": "Kabupaten Balangan",     "kode": "balangankab"},
    {"nama": "Kabupaten Tabalong",     "kode": "tabalongkab"},
    {"nama": "Provinsi Kalsel",        "kode": "kalselprov"},
]

BASE_URL = "https://spse.inaproc.id"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/147.0.0.0 Safari/537.36"


def _sb():
    return create_client(SUPABASE_URL, SUPABASE_KEY,
                         options=SyncClientOptions(httpx_client=httpx.Client()))


def strip_html(text):
    return re.sub(r"<[^>]+>", "", str(text)).strip()


def parse_rupiah(text):
    """Kembalikan string asli nilai rupiah, bersihkan dari noise."""
    t = str(text).strip()
    if t in ("nan", "NaN", "None", "-", ""):
        return "0"
    return t


# --- HTTP helpers (requests.Session — cookie otomatis) ---
def buat_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    })
    return s


def fetch_html(session, url, referer=None, data=None):
    headers = {"Referer": referer or url}
    if data:
        headers["Accept"] = "application/json, text/javascript, */*; q=0.01"
        headers["X-Requested-With"] = "XMLHttpRequest"
        r = session.post(url, data=data, headers=headers, timeout=20)
    else:
        headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        r = session.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    return r.text


def get_session(session, kode_lpse, endpoint):
    url = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    html = fetch_html(session, url)
    time.sleep(2)
    m = re.search(r"authenticityToken = '([a-f0-9]+)'", html)
    return m.group(1) if m else ""


def get_list_paket(session, token, kode_lpse, endpoint, endpoint_dt, tahun):
    url = f"{BASE_URL}/{kode_lpse}/dt/{endpoint_dt}?tahun={tahun}"
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
                raw = fetch_html(session, url, referer=referer, data=data)
                parsed = json.loads(raw)
                rows = parsed.get("data", [])
                break
            except Exception as e:
                if attempt < MAX_RETRY - 1:
                    log.warning(f"  Retry {attempt+1}/{MAX_RETRY} get_list start={start}: {e}")
                    time.sleep(5 * (attempt + 1))
                else:
                    log.error(f"  get_list_paket gagal start={start}: {e}")
                    return semua
        if not rows:
            break
        semua.extend(rows)
        start += PAGE
        time.sleep(1)
        if len(rows) < PAGE:
            break
    return semua


def parse_tabel_kv(html, tabel_idx=0):
    """Parse tabel key-value (col0=key, col1=value) → dict."""
    try:
        tables = pd.read_html(StringIO(html))
        if len(tables) <= tabel_idx:
            return {}
        t = tables[tabel_idx]
        kv = {}
        for _, row in t.iterrows():
            vals = [str(v).strip() for v in row.values]
            if len(vals) >= 2 and vals[0] not in ("nan", ""):
                kv[vals[0]] = vals[1] if vals[1] != "nan" else ""
            if len(vals) >= 4 and vals[2] not in ("nan", ""):
                kv[vals[2]] = vals[3] if vals[3] != "nan" else ""
        return kv
    except:
        return {}


def parse_tabel_rows(html, tabel_idx=1):
    """Parse tabel rows (header row + data) → list of dict."""
    try:
        tables = pd.read_html(StringIO(html))
        if len(tables) <= tabel_idx:
            return []
        t = tables[tabel_idx]
        return t.to_dict(orient="records")
    except:
        return []


def get_detail_tender(session, kode_lpse, kode_tender):
    """
    Ambil semua detail untuk 1 paket tender:
    - pengumumanlelang: info dasar + kode RUP + lokasi + bobot + peserta
    - peserta: semua peserta + harga penawaran
    - evaluasi/hasil: skor evaluasi per peserta
    - evaluasi/pemenang: pemenang + NPWP + harga negosiasi
    - evaluasi/pemenangberkontrak: harga kontrak + PDN + UMK
    - jadwal: tanggal kontrak
    """
    base   = f"{BASE_URL}/{kode_lpse}/lelang/{kode_tender}"
    base_e = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    ref    = f"{BASE_URL}/{kode_lpse}/lelang"
    detail = {}

    # 1. pengumumanlelang
    try:
        html = fetch_html(session, f"{base}/pengumumanlelang", referer=ref)
        kv = parse_tabel_kv(html, 0)
        detail["jenis_pengadaan"]          = kv.get("Jenis Pengadaan", "-")
        detail["satuan_kerja"]             = kv.get("Satuan Kerja", "-")
        detail["nilai_pagu"]               = kv.get("Nilai Pagu Paket", "0")
        detail["tahun_anggaran"]           = kv.get("Tahun Anggaran", "-")
        detail["nilai_hps"]                = kv.get("Nilai HPS Paket", "0")
        detail["metode_pengadaan"]         = kv.get("Metode Pengadaan", "-")
        detail["jenis_kontrak"]            = kv.get("Jenis Kontrak", "-")
        detail["lokasi_pekerjaan"]         = kv.get("Lokasi Pekerjaan", "-")
        detail["alasan_ulang"]             = kv.get("Alasan di ulang", "")

        # Bobot teknis/biaya
        try:
            detail["bobot_teknis"] = float(kv.get("Bobot Teknis", "") or 0) or None
            detail["bobot_biaya"]  = float(kv.get("Bobot Biaya", "") or 0) or None
        except:
            detail["bobot_teknis"] = None
            detail["bobot_biaya"]  = None

        # Jumlah peserta
        m = re.search(r"(\d+)\s*peserta", kv.get("Peserta Tender", ""))
        detail["jumlah_peserta"] = int(m.group(1)) if m else None

        # Kode RUP + Sumber Dana dari tabel[1]
        try:
            tables = pd.read_html(StringIO(html))
            if len(tables) > 1:
                rup_row = tables[1].iloc[0]
                detail["kode_rup"]    = str(rup_row.get("Kode RUP", "") or "").strip()
                detail["sumber_dana"] = str(rup_row.get("Sumber Dana", "") or "").strip()
        except:
            detail["kode_rup"]    = ""
            detail["sumber_dana"] = ""

        # Tanggal pembuatan
        detail["tanggal_pembuatan"] = kv.get("Tanggal Pembuatan", "")

    except Exception as e:
        log.warning(f"  pengumumanlelang {kode_tender}: {e}")

    # 2. /peserta (daftar semua peserta)
    peserta_list = []
    try:
        html_p = fetch_html(session, f"{base}/peserta", referer=ref)
        rows_p = parse_tabel_rows(html_p, 0)
        for row in rows_p:
            peserta_list.append({
                "urutan":           int(row.get("No", 0) or 0),
                "nama_peserta":     str(row.get("Nama Peserta", "") or "").strip(),
                "npwp":             str(row.get("NPWP", "") or "").strip(),
                "harga_penawaran":  parse_rupiah(row.get("Harga Penawaran", "")),
                "harga_terkoreksi": parse_rupiah(row.get("Harga Terkoreksi", "")),
            })
    except Exception as e:
        log.warning(f"  peserta {kode_tender}: {e}")

    # 3. /evaluasi/hasil (skor per peserta)
    skor_map = {}  # nama_peserta → skor
    try:
        html_h = fetch_html(session, f"{base_e}/hasil", referer=ref)
        rows_h = parse_tabel_rows(html_h, 0)
        for row in rows_h:
            nama = str(row.get("Nama Peserta", "") or "").strip()
            skor_map[nama] = {
                "skor_administrasi": _to_float(row.get("SK")),
                "skor_teknis":       _to_float(row.get("ST")),
                "skor_harga":        _to_float(row.get("SH")),
                "skor_akhir":        _to_float(row.get("SA")),
                "harga_negosiasi":   parse_rupiah(row.get("HN", "")),
                "alasan_gugur":      str(row.get("Alasan", "") or "").strip(),
            }
    except Exception as e:
        log.warning(f"  evaluasi/hasil {kode_tender}: {e}")

    # 4. /evaluasi/pemenang
    detail["nama_pemenang"]             = "Belum Ada Pemenang"
    detail["npwp_pemenang"]             = ""
    detail["harga_penawaran_pemenang"]  = "0"
    detail["harga_terkoreksi_pemenang"] = "0"
    detail["harga_negosiasi"]           = "0"
    nama_pemenang = ""
    try:
        html_w = fetch_html(session, f"{base_e}/pemenang", referer=ref)
        rows_w = parse_tabel_rows(html_w, 1)
        if rows_w:
            w = rows_w[0]
            nama_pemenang                       = str(w.get("Nama Pemenang", "") or "").strip()
            detail["nama_pemenang"]             = nama_pemenang or "Belum Ada Pemenang"
            detail["npwp_pemenang"]             = str(w.get("NPWP", "") or "").strip()
            detail["harga_penawaran_pemenang"]  = parse_rupiah(w.get("Harga Penawaran", ""))
            detail["harga_terkoreksi_pemenang"] = parse_rupiah(w.get("Harga Terkoreksi", ""))
            detail["harga_negosiasi"]           = parse_rupiah(w.get("Harga Negosiasi", ""))
    except Exception as e:
        log.warning(f"  pemenang {kode_tender}: {e}")

    # 5. /evaluasi/pemenangberkontrak
    detail["pemenang_berkontrak"] = "Belum Ada Kontrak"
    detail["alamat"]              = "-"
    detail["harga_kontrak"]       = "0"
    detail["nilai_pdn"]           = "0"
    detail["nilai_umk"]           = "0"
    try:
        html_k = fetch_html(session, f"{base_e}/pemenangberkontrak", referer=ref)
        rows_k = parse_tabel_rows(html_k, 1)
        if rows_k:
            k = rows_k[0]
            detail["pemenang_berkontrak"] = str(k.get("Nama Pemenang", "") or "Belum Ada Kontrak").strip()
            detail["alamat"]              = str(k.get("Alamat", "-") or "-").strip()
            detail["harga_kontrak"]       = parse_rupiah(k.get("Harga Kontrak", ""))
            detail["nilai_pdn"]           = parse_rupiah(k.get("Nilai PDN", ""))
            detail["nilai_umk"]           = parse_rupiah(k.get("Nilai UMK", ""))
    except Exception as e:
        log.warning(f"  pemenangberkontrak {kode_tender}: {e}")

    # 6. /jadwal — tanggal kontrak
    detail["kontrak_mulai"]   = "-"
    detail["kontrak_selesai"] = "-"
    try:
        html_j = fetch_html(session, f"{base}/jadwal", referer=ref)
        for _, row in pd.read_html(StringIO(html_j))[0].iterrows():
            vals = [str(v) for v in row.values]
            if any("Penandatanganan Kontrak" in v for v in vals):
                detail["kontrak_mulai"]   = vals[2] if len(vals) > 2 else "-"
                detail["kontrak_selesai"] = vals[3] if len(vals) > 3 else "-"
                break
    except Exception as e:
        log.warning(f"  jadwal {kode_tender}: {e}")

    # Gabung skor ke peserta_list + tandai pemenang
    for p in peserta_list:
        skor = skor_map.get(p["nama_peserta"], {})
        p["skor_administrasi"] = skor.get("skor_administrasi")
        p["skor_teknis"]       = skor.get("skor_teknis")
        p["skor_harga"]        = skor.get("skor_harga")
        p["skor_akhir"]        = skor.get("skor_akhir")
        # harga_negosiasi dari skor_map jika ada, else dari pemenang
        p["harga_negosiasi"]   = skor.get("harga_negosiasi", "0")
        p["alasan_gugur"]      = skor.get("alasan_gugur", "")
        p["is_pemenang"]       = (nama_pemenang != "" and p["nama_peserta"] == nama_pemenang)

    return detail, peserta_list


def _to_float(val):
    try:
        f = float(val)
        return None if (f != f) else f  # NaN check
    except:
        return None


def get_detail_nontender(session, kode_lpse, endpoint, kode_tender, suffix):
    """Detail untuk Non Tender / Pencatatan — endpoint berbeda, tanpa peserta."""
    base   = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}"
    base_e = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    ref    = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    detail = {}

    try:
        html = fetch_html(session, f"{base}/{suffix}", referer=ref)
        kv = parse_tabel_kv(html, 0)
        detail["jenis_pengadaan"]  = kv.get("Jenis Pengadaan", "-")
        detail["satuan_kerja"]     = kv.get("Satuan Kerja", "-")
        detail["nilai_pagu"]       = kv.get("Nilai Pagu Paket", "0")
        detail["tahun_anggaran"]   = kv.get("Tahun Anggaran", "-")
        detail["nilai_hps"]        = kv.get("Nilai HPS Paket", "0")
        detail["metode_pengadaan"] = kv.get("Metode Pengadaan", "-")
    except Exception as e:
        log.warning(f"  {suffix} {kode_tender}: {e}")

    detail["nama_pemenang"] = "Belum Ada Pemenang"
    try:
        html_w = fetch_html(session, f"{base_e}/pemenang", referer=ref)
        rows_w = parse_tabel_rows(html_w, 1)
        if rows_w:
            detail["nama_pemenang"] = str(rows_w[0].get("Nama Pemenang", "") or "Belum Ada Pemenang").strip()
    except Exception as e:
        log.warning(f"  pemenang {kode_tender}: {e}")

    detail["pemenang_berkontrak"] = "Belum Ada Kontrak"
    detail["alamat"]              = "-"
    detail["harga_kontrak"]       = "0"
    try:
        html_k = fetch_html(session, f"{base_e}/pemenangberkontrak", referer=ref)
        rows_k = parse_tabel_rows(html_k, 1)
        if rows_k:
            k = rows_k[0]
            detail["pemenang_berkontrak"] = str(k.get("Nama Pemenang", "") or "Belum Ada Kontrak").strip()
            detail["alamat"]              = str(k.get("Alamat", "-") or "-").strip()
            detail["harga_kontrak"]       = parse_rupiah(k.get("Harga Kontrak", ""))
    except Exception as e:
        log.warning(f"  pemenangberkontrak {kode_tender}: {e}")

    detail["kontrak_mulai"]   = "-"
    detail["kontrak_selesai"] = "-"
    try:
        html_j = fetch_html(session, f"{base}/jadwal", referer=ref)
        for _, row in pd.read_html(StringIO(html_j))[0].iterrows():
            vals = [str(v) for v in row.values]
            if any("Penandatanganan Kontrak" in v for v in vals):
                detail["kontrak_mulai"]   = vals[2] if len(vals) > 2 else "-"
                detail["kontrak_selesai"] = vals[3] if len(vals) > 3 else "-"
                break
    except Exception as e:
        log.warning(f"  jadwal {kode_tender}: {e}")

    return detail


def scrape_satu_lpse(kode_lpse, nama_lpse, endpoint, endpoint_dt, suffix, tabel, tahun):
    log.info(f"=== Mulai: {nama_lpse} ({endpoint}, tahun {tahun}) ===")
    is_tender = (tabel == "tender")

    try:
        session = buat_session()
        token   = get_session(session, kode_lpse, endpoint)
        rows    = get_list_paket(session, token, kode_lpse, endpoint, endpoint_dt, tahun)
        log.info(f"  {len(rows)} paket ditemukan")
    except Exception as e:
        log.error(f"  Gagal ambil list: {e}")
        return 0, 0

    sb = _sb()
    ok = err = 0

    for row in rows:
        try:
            kode_tender = str(row[0]).strip()
            nama_paket  = strip_html(row[1])
            instansi    = strip_html(row[2])
            tahapan     = strip_html(row[3])
            if not nama_paket or nama_paket == "nan":
                continue

            if is_tender:
                # Kolom ekstra dari DataTables
                sistem_evaluasi    = strip_html(row[7]) if len(row) > 7 else "-"
                metode_kualifikasi = strip_html(row[5]) if len(row) > 5 else "-"

                detail, peserta_list = get_detail_tender(session, kode_lpse, kode_tender)

                data = {
                    "kode_tender":              kode_tender,
                    "nama_paket":               nama_paket,
                    "instansi":                 instansi,
                    "tahapan":                  tahapan,
                    "link_detail":              f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/{suffix}",
                    "sistem_evaluasi":          sistem_evaluasi,
                    "metode_kualifikasi":       metode_kualifikasi,
                    "jenis_pengadaan":          detail.get("jenis_pengadaan", "-"),
                    "satuan_kerja":             detail.get("satuan_kerja", "-"),
                    "nilai_pagu":               detail.get("nilai_pagu", "0"),
                    "tahun_anggaran":           detail.get("tahun_anggaran", "-"),
                    "nilai_hps":                detail.get("nilai_hps", "0"),
                    "metode_pengadaan":         detail.get("metode_pengadaan", "-"),
                    "jenis_kontrak":            detail.get("jenis_kontrak", "-"),
                    "lokasi_pekerjaan":         detail.get("lokasi_pekerjaan", "-"),
                    "bobot_teknis":             detail.get("bobot_teknis"),
                    "bobot_biaya":              detail.get("bobot_biaya"),
                    "jumlah_peserta":           detail.get("jumlah_peserta"),
                    "alasan_ulang":             detail.get("alasan_ulang", ""),
                    "kode_rup":                 detail.get("kode_rup", ""),
                    "sumber_dana":              detail.get("sumber_dana", ""),
                    "tanggal_pembuatan":        detail.get("tanggal_pembuatan", ""),
                    "nama_pemenang":            detail.get("nama_pemenang", "Belum Ada Pemenang"),
                    "npwp_pemenang":            detail.get("npwp_pemenang", ""),
                    "harga_penawaran_pemenang": detail.get("harga_penawaran_pemenang", "0"),
                    "harga_terkoreksi_pemenang":detail.get("harga_terkoreksi_pemenang", "0"),
                    "harga_negosiasi":          detail.get("harga_negosiasi", "0"),
                    "pemenang_berkontrak":      detail.get("pemenang_berkontrak", "Belum Ada Kontrak"),
                    "alamat":                   detail.get("alamat", "-"),
                    "harga_kontrak":            detail.get("harga_kontrak", "0"),
                    "nilai_pdn":                detail.get("nilai_pdn", "0"),
                    "nilai_umk":                detail.get("nilai_umk", "0"),
                    "kontrak_mulai":            detail.get("kontrak_mulai", "-"),
                    "kontrak_selesai":          detail.get("kontrak_selesai", "-"),
                    "diambil_pada":             datetime.now().isoformat(),
                }
                sb.table("tender").upsert(data).execute()

                # Upsert peserta
                if peserta_list:
                    for p in peserta_list:
                        p["kode_tender"]  = kode_tender
                        p["diambil_pada"] = datetime.now().isoformat()
                    sb.table("tender_peserta").upsert(
                        peserta_list, on_conflict="kode_tender,urutan"
                    ).execute()

            else:
                detail = get_detail_nontender(session, kode_lpse, endpoint, kode_tender, suffix)
                data = {
                    "kode_tender":         kode_tender,
                    "nama_paket":          nama_paket,
                    "instansi":            instansi,
                    "tahapan":             tahapan,
                    "link_detail":         f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/{suffix}",
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
                sb.table(tabel).upsert(data).execute()

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
    tahun       = os.environ.get("SCRAPE_TAHUN", "2026")
    kategori    = os.environ.get("SCRAPE_KATEGORI", "Tender")
    filter_kode = os.environ.get("SCRAPE_KODE_LPSE", "").strip()

    if kategori not in CONFIG_KATEGORI:
        log.error(f"Kategori tidak valid: {kategori}. Pilih: {list(CONFIG_KATEGORI)}")
        sys.exit(1)

    config      = CONFIG_KATEGORI[kategori]
    endpoint    = config["endpoint"]
    endpoint_dt = config["endpoint_dt"]
    suffix      = config["suffix"]
    tabel       = config["tabel"]

    targets = DAFTAR_LPSE
    if filter_kode:
        targets = [l for l in DAFTAR_LPSE if l["kode"] == filter_kode]
        if not targets:
            log.error(f"Kode LPSE tidak ditemukan: {filter_kode}")
            sys.exit(1)

    log.info(f"Scrape {kategori} tahun {tahun} -> tabel '{tabel}' | {len(targets)} LPSE")

    total_ok = total_err = 0
    for i, lpse in enumerate(targets):
        ok, err = scrape_satu_lpse(
            lpse["kode"], lpse["nama"],
            endpoint, endpoint_dt, suffix, tabel, tahun
        )
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
