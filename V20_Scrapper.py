import streamlit as st
import urllib.request
import urllib.parse
import http.cookiejar
import json
import re
import logging
import pandas as pd
from io import StringIO
from supabase import create_client
import os
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SPSE Scraper V2.0", page_icon="🏗️", layout="wide")

st.markdown("""
    <style>
        .stCheckbox { margin-bottom: -15px; }
        .block-container { padding-top: 1rem; padding-bottom: 2rem; }
        h1 { padding-top: 0rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. SETUP DATA ---
_dir1 = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
_env_path = os.path.join(_dir1, "secret_supabase.env")
if not os.path.exists(_env_path):
    _env_path = os.path.join(os.getcwd(), "secret_supabase.env")
load_dotenv(dotenv_path=_env_path)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
STOP_FILE = os.path.join(os.path.dirname(_env_path), "stop_signal.txt")
LOG_FILE = os.path.join(os.path.dirname(_env_path), "scraper.log")

logging.basicConfig(
    filename=LOG_FILE, level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S", encoding="utf-8"
)
log = logging.getLogger("scraper")

# --- MAPPING KATEGORI ---
CONFIG_KATEGORI = {
    "Tender": {
        "tabel": "tender",
        "endpoint": "lelang"
    },
    "Non Tender": {
        "tabel": "non_tender",
        "endpoint": "nontender"
    },
    "Pencatatan": {
        "tabel": "pencatatan",
        "endpoint": "pencatatan"
    }
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
    {"nama": "Provinsi Kalsel",         "kode": "kalselprov"}
]

BASE_URL = "https://spse.inaproc.id"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"

try:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
except:
    st.error("Koneksi Supabase Gagal.")
    st.stop()

# --- 3. UTILITY ---
def toggle_all():
    state = st.session_state.chk_all
    for lpse in DAFTAR_LPSE:
        st.session_state[f"chk_{lpse['kode']}"] = state

def get_last_update(kode_lpse, kategori_selected):
    nama_tabel = CONFIG_KATEGORI[kategori_selected]["tabel"]
    try:
        response = supabase.table(nama_tabel) \
            .select('diambil_pada') \
            .ilike('link_detail', f'%{kode_lpse}%') \
            .order('diambil_pada', desc=True) \
            .limit(1) \
            .execute()
        if response.data and response.data[0]['diambil_pada']:
            tgl_raw = response.data[0]['diambil_pada'].split('+')[0].replace("Z", "")
            try:    waktu_utc = datetime.strptime(tgl_raw, "%Y-%m-%dT%H:%M:%S.%f")
            except:
                try: waktu_utc = datetime.strptime(tgl_raw, "%Y-%m-%dT%H:%M:%S")
                except: return f"⚠️ {tgl_raw[:10]}"
            waktu_wib = waktu_utc + timedelta(hours=7)
            return f"🟢 {waktu_wib.strftime('%d/%m %H:%M')}"
        return "⚪ Kosong"
    except: return "-"

# --- 4. ENGINE SCRAPING (urllib, tanpa Selenium) ---

def buat_opener():
    """Buat urllib opener dengan cookie jar."""
    cj = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

def fetch_html(opener, url, referer=None, data=None):
    """GET atau POST dengan cookie otomatis. Kembalikan HTML string."""
    headers = {
        "User-Agent": UA,
        "Referer": referer or url,
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    }
    body = None
    if data:
        body = urllib.parse.urlencode(data).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
        headers["X-Requested-With"] = "XMLHttpRequest"
    req = urllib.request.Request(url, data=body, headers=headers)
    with opener.open(req, timeout=20) as r:
        return r.read().decode("utf-8", errors="replace")

def get_session(opener, kode_lpse, endpoint):
    """GET halaman list → dapat session cookie + CSRF token."""
    url = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    html = fetch_html(opener, url)
    m = re.search(r"authenticityToken = '([a-f0-9]+)'", html)
    return m.group(1) if m else ""

def get_list_paket(opener, token, kode_lpse, endpoint, tahun):
    """Loop pagination DataTables API → list semua kode paket."""
    url = f"{BASE_URL}/{kode_lpse}/dt/{endpoint}?tahun={tahun}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    semua = []
    start = 0
    PAGE = 100
    while True:
        data = {
            "draw": str(start // PAGE + 1),
            "start": str(start),
            "length": str(PAGE),
            "authenticityToken": token,
        }
        try:
            raw = fetch_html(opener, url, referer=referer, data=data)
            parsed = json.loads(raw)
            rows = parsed.get("data", [])
            if not rows:
                break
            semua.extend(rows)
            start += PAGE
            # Kalau total sudah terpenuhi (sentinel 2147483647 diabaikan)
            if len(rows) < PAGE:
                break
        except Exception as e:
            log.warning(f"  get_list_paket error start={start}: {e}")
            break
    return semua

def parse_tabel1(html):
    """Ambil Tabel index 1 dari HTML (data detail). Kembalikan dict baris pertama atau {}."""
    try:
        tables = pd.read_html(StringIO(html))
        if len(tables) > 1 and len(tables[1]) > 0:
            row = tables[1].iloc[0]
            return {str(k): str(v) for k, v in row.items()}
    except:
        pass
    return {}

def get_detail_paket(opener, kode_lpse, endpoint, kode_tender):
    """Ambil detail paket dari 4 endpoint. Kembalikan dict field."""
    base_lelang  = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}"
    base_evaluasi = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    detail = {}

    # --- Pengumuman: pagu, HPS, metode, satker, jenis, tahun anggaran ---
    try:
        html = fetch_html(opener, f"{base_lelang}/pengumumanlelang", referer=referer)
        tables = pd.read_html(StringIO(html))
        if tables:
            t = tables[0]
            # Tabel 0: key-value. Baris normal: col0=key, col1=value.
            # Baris khusus (Pagu+HPS): col0=Nilai Pagu, col1=nilainya, col2=Nilai HPS, col3=nilainya
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

    # --- Pemenang ---
    detail["nama_pemenang"] = "Belum Ada Pemenang"
    try:
        html = fetch_html(opener, f"{base_evaluasi}/pemenang", referer=referer)
        d = parse_tabel1(html)
        if d:
            detail["nama_pemenang"] = d.get("Nama Pemenang", "Belum Ada Pemenang")
    except Exception as e:
        log.warning(f"  pemenang {kode_tender}: {e}")

    # --- Pemenang Berkontrak ---
    detail["pemenang_berkontrak"] = "Belum Ada Kontrak"
    detail["alamat"]              = "-"
    detail["harga_kontrak"]       = "0"
    try:
        html = fetch_html(opener, f"{base_evaluasi}/pemenangberkontrak", referer=referer)
        d = parse_tabel1(html)
        if d:
            detail["pemenang_berkontrak"] = d.get("Nama Pemenang", "Belum Ada Kontrak")
            detail["alamat"]              = d.get("Alamat", "-")
            detail["harga_kontrak"]       = d.get("Harga Kontrak", "0")
    except Exception as e:
        log.warning(f"  pemenangberkontrak {kode_tender}: {e}")

    # --- Jadwal: tanggal penandatanganan kontrak ---
    detail["kontrak_mulai"]  = "-"
    detail["kontrak_selesai"] = "-"
    try:
        html = fetch_html(opener, f"{base_lelang}/jadwal", referer=referer)
        tables = pd.read_html(StringIO(html))
        if tables:
            for _, row in tables[0].iterrows():
                vals = [str(v) for v in row.values]
                if any("Penandatanganan Kontrak" in v for v in vals):
                    # kolom: no | tahap | mulai | sampai | perubahan
                    detail["kontrak_mulai"]  = vals[2] if len(vals) > 2 else "-"
                    detail["kontrak_selesai"] = vals[3] if len(vals) > 3 else "-"
                    break
    except Exception as e:
        log.warning(f"  jadwal {kode_tender}: {e}")

    return detail

def scrape_satu_lpse(target, tahun_pilihan, kategori_pilihan, supabase_client):
    if os.path.exists(STOP_FILE):
        return "⛔ Dibatalkan"

    config      = CONFIG_KATEGORI[kategori_pilihan]
    tabel       = config["tabel"]
    endpoint    = config["endpoint"]
    nama_lpse   = target["nama"]
    kode_lpse   = target["kode"]

    log.info(f"Mulai scrape: {nama_lpse} ({kategori_pilihan})")

    try:
        opener = buat_opener()
        token  = get_session(opener, kode_lpse, endpoint)
        rows   = get_list_paket(opener, token, kode_lpse, endpoint, tahun_pilihan)
        log.info(f"  {nama_lpse}: {len(rows)} paket ditemukan di list")
    except Exception as e:
        log.error(f"  {nama_lpse}: Gagal ambil list — {e}")
        return f"❌ {nama_lpse}: Gagal ambil list"

    paket_counter = 0
    error_counter = 0

    for row in rows:
        if os.path.exists(STOP_FILE):
            return f"🛑 {nama_lpse}: STOP"

        try:
            kode_tender = str(row[0]).strip()
            nama_paket  = str(row[1]).strip()
            instansi    = str(row[2]).strip()
            tahapan     = str(row[3]).strip()
            link_detail = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/pengumumanlelang"

            if not nama_paket or nama_paket == "nan":
                continue

            detail = get_detail_paket(opener, kode_lpse, endpoint, kode_tender)

            data = {
                "kode_tender":        kode_tender,
                "nama_paket":         nama_paket,
                "instansi":           instansi,
                "tahapan":            tahapan,
                "link_detail":        link_detail,
                "jenis_pengadaan":    detail.get("jenis_pengadaan", "-"),
                "satuan_kerja":       detail.get("satuan_kerja", "-"),
                "nilai_pagu":         detail.get("nilai_pagu", "0"),
                "tahun_anggaran":     detail.get("tahun_anggaran", "-"),
                "nilai_hps":          detail.get("nilai_hps", "0"),
                "metode_pengadaan":   detail.get("metode_pengadaan", "-"),
                "nama_pemenang":      detail.get("nama_pemenang", "Belum Ada Pemenang"),
                "pemenang_berkontrak":detail.get("pemenang_berkontrak", "Belum Ada Kontrak"),
                "alamat":             detail.get("alamat", "-"),
                "harga_kontrak":      detail.get("harga_kontrak", "0"),
                "kontrak_mulai":      detail.get("kontrak_mulai", "-"),
                "kontrak_selesai":    detail.get("kontrak_selesai", "-"),
                "diambil_pada":       datetime.now().isoformat(),
            }

            supabase_client.table(tabel).upsert(data).execute()
            paket_counter += 1

        except Exception as e:
            error_counter += 1
            log.warning(f"  Error paket {row[0] if row else '?'}: {str(e)[:150]}")
            continue

    msg = f"{nama_lpse}: ✅ {paket_counter} data"
    if error_counter:
        msg += f", ⚠️ {error_counter} error"
    log.info(msg)
    return msg

# --- 5. UI ---
st.title("🏗️ SPSE Scraper V2.0 (urllib — No Chrome)")

col_thn, col_kat = st.columns([1, 2])
with col_thn:
    tahun_sekarang = datetime.now().year
    tahun_input = st.number_input("Tahun", min_value=2020, max_value=2030, value=tahun_sekarang)
with col_kat:
    kategori_input = st.radio("Kategori (Target Tabel)", ["Tender", "Non Tender", "Pencatatan"], horizontal=True)

st.write("---")
c1, c2, c3 = st.columns([0.5, 4, 3])
c1.checkbox("", key="chk_all", on_change=toggle_all)
c2.markdown("**DAERAH**")
c3.markdown(f"**LAST UPDATE (Tabel: {CONFIG_KATEGORI[kategori_input]['tabel']})**")

target_dipilih = []
with st.container():
    for lpse in DAFTAR_LPSE:
        col_a, col_b, col_c = st.columns([0.5, 4, 3])
        with col_a:
            if f"chk_{lpse['kode']}" not in st.session_state:
                st.session_state[f"chk_{lpse['kode']}"] = False
            if col_a.checkbox("", key=f"chk_{lpse['kode']}"):
                target_dipilih.append(lpse)
        with col_b: st.markdown(f"{lpse['nama']}")
        with col_c: st.markdown(get_last_update(lpse['kode'], kategori_input))

st.write("---")
col_opts, _ = st.columns([2, 3])
with col_opts:
    max_workers_input = st.slider("Parallel Workers", min_value=1, max_value=5, value=2)

st.info("ℹ️ V2.0 menggunakan urllib (tanpa Chrome). Headless/minimize tidak diperlukan lagi.")

c_start, c_stop = st.columns([4, 1])

with c_start:
    if st.button(f"🚀 GAS SCRAPE ({len(target_dipilih)})", type="primary", use_container_width=True):
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)
        if not target_dipilih:
            st.warning("Pilih daerah!")
        else:
            status_box = st.status(
                f"Menghisap data ke tabel: **{CONFIG_KATEGORI[kategori_input]['tabel']}**...",
                expanded=True
            )
            p_bar = status_box.progress(0)
            with ThreadPoolExecutor(max_workers=max_workers_input) as executor:
                futures = [
                    executor.submit(scrape_satu_lpse, t, tahun_input, kategori_input, supabase)
                    for t in target_dipilih
                ]
                done = 0
                for f in futures:
                    res = f.result()
                    done += 1
                    p_bar.progress(int((done / len(futures)) * 100))
                    status_box.write(res)

            if not os.path.exists(STOP_FILE):
                status_box.update(label="Selesai!", state="complete", expanded=False)
                st.balloons()
                st.cache_data.clear()
                import time; time.sleep(2)
                st.rerun()

with c_stop:
    if st.button("⛔ STOP", type="secondary"):
        with open(STOP_FILE, "w") as f:
            f.write("STOP")
        st.toast("Stop signal sent.")
