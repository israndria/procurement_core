import streamlit as st
import urllib.request
import urllib.parse
import json
import re
import logging
import subprocess
import pandas as pd
import cloudscraper
from io import StringIO, BytesIO
from supabase import create_client
import os
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SPSE Scraper V2.1", page_icon="🏗️", layout="wide")

st.markdown("""
    <style>
        .stCheckbox { margin-bottom: -15px; }
        .block-container { padding-top: 1rem; padding-bottom: 2rem; }
        h1 { padding-top: 0rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. SETUP DATA ---
def _load_env_manual():
    """Baca .env langsung tanpa load_dotenv."""
    import pathlib
    kandidat = [
        pathlib.Path(__file__).resolve().parent / "secret_supabase.env",
        pathlib.Path(os.getcwd()) / "secret_supabase.env",
        pathlib.Path(r"D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\secret_supabase.env"),
    ]
    errs = []
    for path in kandidat:
        try:
            with open(str(path), encoding="utf-8-sig") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, val = line.split("=", 1)
                        os.environ[key.strip()] = val.strip().strip('"').strip("'")
            return str(path)
        except Exception as e:
            errs.append(f"{path}: {e}")
    raise FileNotFoundError(f"secret_supabase.env tidak ditemukan.\n" + "\n".join(errs))

try:
    _env_path = _load_env_manual()
except FileNotFoundError as _fe:
    st.error(str(_fe))
    st.stop()
_dir1 = os.path.dirname(_env_path)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
STOP_FILE = os.path.join(os.path.dirname(_env_path), "stop_signal.txt")
LOG_FILE  = os.path.join(os.path.dirname(_env_path), "scraper.log")

logging.basicConfig(
    filename=LOG_FILE, level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S", encoding="utf-8"
)
log = logging.getLogger("scraper")
# Suppress noise dari httpx/httpcore (Supabase HTTP log)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

CONFIG_KATEGORI = {
    "Tender":     {"tabel": "tender",     "endpoint": "lelang",     "endpoint_dt": "lelang", "suffix": "pengumumanlelang"},
    "Non Tender": {"tabel": "non_tender", "endpoint": "nontender",  "endpoint_dt": "pl",     "suffix": "pengumumanpl"},
    "Pencatatan": {"tabel": "pencatatan", "endpoint": "pencatatan", "endpoint_dt": "nonspk", "suffix": "pengumuman"},
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
    return re.sub(r"<[^>]+>", "", str(text)).strip()

try:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError(f"Env tidak terbaca. Path: {_env_path} | URL: {SUPABASE_URL}")
except Exception as _e:
    st.error(f"Koneksi Supabase Gagal: {_e}")
    st.stop()

def _sb():
    """Buat Supabase client baru dengan fresh httpx.Client() — hindari stale connection pool."""
    import httpx
    from supabase.lib.client_options import SyncClientOptions
    return create_client(SUPABASE_URL, SUPABASE_KEY,
                         options=SyncClientOptions(httpx_client=httpx.Client()))

# Global untuk backward-compat (snapshot/notifikasi di Tab Scraper)
supabase = _sb()

# --- 3. UTILITY ---
def toggle_all():
    state = st.session_state.chk_all
    for lpse in DAFTAR_LPSE:
        st.session_state[f"chk_{lpse['kode']}"] = state

def parse_wib(tgl_raw):
    tgl_raw = tgl_raw.split('+')[0].replace("Z", "")
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(tgl_raw, fmt)
        except: pass
    return None

def _query_last_update(kategori_selected):
    """Query last update per LPSE — 1 query kecil per LPSE agar tidak kena batas 1000 row Supabase."""
    nama_tabel = CONFIG_KATEGORI[kategori_selected]["tabel"]
    hasil = {}
    sb = _sb()
    for lpse in DAFTAR_LPSE:
        kode = lpse["kode"]
        try:
            r = sb.table(nama_tabel) \
                .select('diambil_pada') \
                .ilike('link_detail', f'%/{kode}/%') \
                .order('diambil_pada', desc=True).limit(1).execute()
            if r.data and r.data[0].get('diambil_pada'):
                wib = parse_wib(r.data[0]['diambil_pada'])
                hasil[kode] = f"🟢 {wib.strftime('%d/%m %H:%M')}" if wib else "⚪ Kosong"
            else:
                hasil[kode] = "⚪ Kosong"
        except:
            hasil[kode] = "-"
    return hasil

def get_all_last_update(kategori_selected):
    """Bulk query 1x untuk semua LPSE, dengan cache TTL 120 detik di session_state."""
    import time
    ck = f"_lu_{kategori_selected}"
    ct = f"_lu_ts_{kategori_selected}"
    now = time.time()
    if ck in st.session_state and (now - st.session_state.get(ct, 0)) < 120:
        return st.session_state[ck]
    try:
        hasil = _query_last_update(kategori_selected)
    except Exception as e:
        st.warning(f"⚠️ Gagal baca last update: {e}")
        hasil = {}
    st.session_state[ck] = hasil
    st.session_state[ct] = now
    return hasil

def get_last_update(kode_lpse, kategori_selected, bulk_map=None):
    if bulk_map is not None:
        return bulk_map.get(kode_lpse, "⚪ Kosong")
    # Fallback: query langsung (tidak dipakai di UI normal)
    nama_tabel = CONFIG_KATEGORI[kategori_selected]["tabel"]
    try:
        response = supabase.table(nama_tabel) \
            .select('diambil_pada') \
            .ilike('link_detail', f'%{kode_lpse}%') \
            .order('diambil_pada', desc=True).limit(1).execute()
        if response.data and response.data[0]['diambil_pada']:
            wib = parse_wib(response.data[0]['diambil_pada'])
            if wib: return f"🟢 {wib.strftime('%d/%m %H:%M')}"
        return "⚪ Kosong"
    except: return "-"

# --- 4. ENGINE SCRAPING ---
def buat_opener():
    """Buat cloudscraper session — meniru TLS fingerprint Chrome agar lolos Cloudflare."""
    return cloudscraper.create_scraper(
        browser={"browser": "chrome", "platform": "windows", "mobile": False}
    )

def fetch_html(opener, url, referer=None, data=None):
    """GET atau POST menggunakan cloudscraper session.
    Parameter 'opener' sekarang adalah cloudscraper session (bukan urllib opener).
    """
    headers = {
        "Referer": referer or url,
        "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    }
    if data:
        headers["Accept"] = "application/json, text/javascript, */*; q=0.01"
        headers["X-Requested-With"] = "XMLHttpRequest"
        r = opener.post(url, data=data, headers=headers, timeout=20)
    else:
        headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        r = opener.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    return r.text

def get_session(opener, kode_lpse, endpoint):
    import time
    url = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    html = fetch_html(opener, url)
    time.sleep(2)
    m = re.search(r"authenticityToken = '([a-f0-9]+)'", html)
    return m.group(1) if m else ""

def get_list_paket(opener, token, kode_lpse, endpoint, endpoint_dt, tahun):
    import time
    url = f"{BASE_URL}/{kode_lpse}/dt/{endpoint_dt}?tahun={tahun}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    semua = []; start = 0; PAGE = 100; MAX_RETRY = 3
    while True:
        data = {"draw": str(start // PAGE + 1), "start": str(start),
                "length": str(PAGE), "authenticityToken": token}
        for attempt in range(MAX_RETRY):
            try:
                raw = fetch_html(opener, url, referer=referer, data=data)
                rows = json.loads(raw).get("data", [])
                break
            except Exception as e:
                if attempt < MAX_RETRY - 1:
                    log.warning(f"  Retry {attempt+1} get_list start={start}: {e}")
                    time.sleep(5 * (attempt + 1))
                else:
                    log.error(f"  get_list gagal start={start}: {e}")
                    return semua
        if not rows: break
        semua.extend(rows); start += PAGE; time.sleep(1)
        if len(rows) < PAGE: break
    return semua

def parse_tabel1(html):
    try:
        tables = pd.read_html(StringIO(html))
        if len(tables) > 1 and len(tables[1]) > 0:
            row = tables[1].iloc[0]
            return {str(k): str(v) for k, v in row.items()}
    except: pass
    return {}

def get_detail_paket(opener, kode_lpse, endpoint, kode_tender, suffix="pengumumanlelang"):
    base_lelang   = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}"
    base_evaluasi = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    detail = {}
    try:
        html = fetch_html(opener, f"{base_lelang}/{suffix}", referer=referer)
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
        log.warning(f"  {suffix} {kode_tender}: {e}")

    detail["nama_pemenang"] = "Belum Ada Pemenang"
    try:
        d = parse_tabel1(fetch_html(opener, f"{base_evaluasi}/pemenang", referer=referer))
        if d: detail["nama_pemenang"] = d.get("Nama Pemenang", "Belum Ada Pemenang")
    except Exception as e: log.warning(f"  pemenang {kode_tender}: {e}")

    detail["pemenang_berkontrak"] = "Belum Ada Kontrak"
    detail["alamat"] = "-"; detail["harga_kontrak"] = "0"
    try:
        d = parse_tabel1(fetch_html(opener, f"{base_evaluasi}/pemenangberkontrak", referer=referer))
        if d:
            detail["pemenang_berkontrak"] = d.get("Nama Pemenang", "Belum Ada Kontrak")
            detail["alamat"]              = d.get("Alamat", "-")
            detail["harga_kontrak"]       = d.get("Harga Kontrak", "0")
    except Exception as e: log.warning(f"  berkontrak {kode_tender}: {e}")

    detail["kontrak_mulai"] = "-"; detail["kontrak_selesai"] = "-"
    try:
        html = fetch_html(opener, f"{base_lelang}/jadwal", referer=referer)
        for _, row in pd.read_html(StringIO(html))[0].iterrows():
            vals = [str(v) for v in row.values]
            if any("Penandatanganan Kontrak" in v for v in vals):
                detail["kontrak_mulai"]   = vals[2] if len(vals) > 2 else "-"
                detail["kontrak_selesai"] = vals[3] if len(vals) > 3 else "-"
                break
    except Exception as e: log.warning(f"  jadwal {kode_tender}: {e}")
    return detail

def scrape_satu_lpse(target, tahun_pilihan, kategori_pilihan, incremental=False):
    if os.path.exists(STOP_FILE): return "⛔ Dibatalkan"
    # Buat Supabase client + httpx client baru per thread — hindari konflik event loop Streamlit
    import httpx
    from supabase.lib.client_options import SyncClientOptions
    sb = create_client(SUPABASE_URL, SUPABASE_KEY,
                       options=SyncClientOptions(httpx_client=httpx.Client()))
    config = CONFIG_KATEGORI[kategori_pilihan]
    tabel = config["tabel"]; endpoint = config["endpoint"]
    endpoint_dt = config.get("endpoint_dt", endpoint)
    suffix = config.get("suffix", "pengumumanlelang")
    nama_lpse = target["nama"]; kode_lpse = target["kode"]
    mode_label = "[INCREMENTAL]" if incremental else "[FULL]"
    log.info(f"Mulai scrape: {nama_lpse} ({kategori_pilihan}) {mode_label}")
    import time as _time
    try:
        opener = buat_opener()
        token  = get_session(opener, kode_lpse, endpoint)
        _time.sleep(1)  # jeda kecil sebelum get_list agar tidak langsung hit server
        rows   = get_list_paket(opener, token, kode_lpse, endpoint, endpoint_dt, tahun_pilihan)
        log.info(f"  {nama_lpse}: {len(rows)} paket ditemukan di SPSE")
    except Exception as e:
        log.error(f"  {nama_lpse}: Gagal ambil list — {e}")
        return f"❌ {nama_lpse}: Gagal ambil list"

    # Incremental: ambil snapshot DB (kode_tender + tahapan) untuk LPSE ini saja
    db_map = {}
    if incremental:
        try:
            snap = sb.table(tabel).select("kode_tender,tahapan") \
                .ilike("link_detail", f"%/{kode_lpse}/%").execute()
            db_map = {r["kode_tender"]: r["tahapan"] for r in snap.data}
            log.info(f"  DB snapshot: {len(db_map)} paket tersimpan untuk {kode_lpse}")
        except Exception as e:
            log.warning(f"  Gagal ambil snapshot DB, fallback ke full: {e}")
            incremental = False

    ok = 0; err = 0; skip = 0; jml_baru = 0; jml_berubah = 0
    for row in rows:
        if os.path.exists(STOP_FILE): return f"🛑 {nama_lpse}: STOP"
        try:
            kode_tender = str(row[0]).strip()
            nama_paket  = strip_html(row[1])
            instansi    = strip_html(row[2])
            tahapan     = strip_html(row[3])
            if not nama_paket or nama_paket == "nan": continue

            # Incremental: skip jika sudah ada dan tahapan tidak berubah
            if incremental and kode_tender in db_map:
                if db_map[kode_tender] == tahapan:
                    skip += 1
                    continue
                else:
                    jml_berubah += 1
            elif incremental:
                jml_baru += 1

            detail = get_detail_paket(opener, kode_lpse, endpoint, kode_tender, suffix)
            data = {
                "kode_tender": kode_tender, "nama_paket": nama_paket,
                "instansi": instansi, "tahapan": tahapan,
                "link_detail": f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/{suffix}",
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
            sb.table(tabel).upsert(data).execute()
            ok += 1
        except Exception as e:
            err += 1
            log.warning(f"  Error {row[0] if row else '?'}: [{type(e).__name__}] {str(e)[:150]}")

    if incremental:
        msg = f"{nama_lpse}: ✅ {ok} diproses ({jml_baru} baru, {jml_berubah} berubah), ⏭️ {skip} dilewati"
    else:
        msg = f"{nama_lpse}: ✅ {ok} data"
    if err: msg += f", ⚠️ {err} error"
    log.info(msg)
    return msg

# ============================================================
# --- 5. UI UTAMA (TABS) ---
# ============================================================
st.title("🏗️ SPSE Scraper V2.1")
tab_scraper, tab_dashboard, tab_log = st.tabs(["🚀 Scraper", "📊 Dashboard", "📋 Log & Status"])

# ============================================================
# TAB 1: SCRAPER
# ============================================================
with tab_scraper:
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

    # Bulk query 1x untuk semua LPSE
    bulk_map = get_all_last_update(kategori_input)

    target_dipilih = []
    for lpse in DAFTAR_LPSE:
        col_a, col_b, col_c = st.columns([0.5, 4, 3])
        if f"chk_{lpse['kode']}" not in st.session_state:
            st.session_state[f"chk_{lpse['kode']}"] = False
        if col_a.checkbox("", key=f"chk_{lpse['kode']}"): target_dipilih.append(lpse)
        col_b.markdown(lpse['nama'])
        col_c.markdown(get_last_update(lpse['kode'], kategori_input, bulk_map))

    st.write("---")
    col_opts, col_mode = st.columns([2, 3])
    with col_opts:
        max_workers_input = st.slider("Parallel Workers", min_value=1, max_value=3, value=1,
                                      help="Non Tender & Pencatatan disarankan pakai 1 worker agar tidak kena rate limit Cloudflare (429)")
    with col_mode:
        mode_scrape = st.radio(
            "Mode Scrape",
            ["🔄 Full (semua paket)", "⚡ Incremental (baru & berubah)"],
            horizontal=True, index=1,
            help="Incremental: hanya fetch detail paket baru atau yang berubah tahapan. Lebih cepat."
        )
        incremental_mode = "Incremental" in mode_scrape
    st.info("ℹ️ V2.1 menggunakan cloudscraper (bypass Cloudflare TLS fingerprint).")

    c_start, c_stop = st.columns([4, 1])
    with c_start:
        if st.button(f"🚀 GAS SCRAPE ({len(target_dipilih)})", type="primary", use_container_width=True):
            if os.path.exists(STOP_FILE): os.remove(STOP_FILE)
            if not target_dipilih:
                st.warning("Pilih daerah!")
            else:
                tabel_aktif = CONFIG_KATEGORI[kategori_input]["tabel"]
                try:
                    snap = _sb().table(tabel_aktif).select("kode_tender").execute()
                    st.session_state["snapshot_sebelum"] = {r["kode_tender"] for r in snap.data}
                except: st.session_state["snapshot_sebelum"] = set()

                status_box = st.status(
                    f"Menghisap data ke tabel: **{CONFIG_KATEGORI[kategori_input]['tabel']}**...",
                    expanded=True)
                p_bar = status_box.progress(0)
                hasil_list = []
                with ThreadPoolExecutor(max_workers=max_workers_input) as executor:
                    futures = [executor.submit(scrape_satu_lpse, t, tahun_input, kategori_input, incremental_mode)
                               for t in target_dipilih]
                    done = 0
                    for f in futures:
                        res = f.result(); done += 1
                        p_bar.progress(int((done / len(futures)) * 100))
                        status_box.write(res)
                        hasil_list.append(res)

                # Simpan hasil ke session_state agar tetap tampil setelah selesai
                st.session_state["hasil_scrape_terakhir"] = {
                    "waktu": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "tabel": tabel_aktif,
                    "mode": "Incremental" if incremental_mode else "Full",
                    "hasil": hasil_list,
                }

                if not os.path.exists(STOP_FILE):
                    try:
                        snap_sesudah = _sb().table(tabel_aktif).select("kode_tender,nama_paket,instansi").execute()
                        kode_sesudah = {r["kode_tender"] for r in snap_sesudah.data}
                        kode_baru = kode_sesudah - st.session_state.get("snapshot_sebelum", set())
                        if kode_baru:
                            paket_baru = [r for r in snap_sesudah.data if r["kode_tender"] in kode_baru]
                            status_box.write(f"🆕 **{len(kode_baru)} paket baru ditemukan:**")
                            for p in paket_baru[:10]:
                                status_box.write(f"  - `{p['kode_tender']}` {p['nama_paket']} ({p['instansi']})")
                            if len(kode_baru) > 10:
                                status_box.write(f"  ... dan {len(kode_baru)-10} lainnya")
                    except: pass

                    status_box.update(label="Selesai!", state="complete", expanded=False)
                    st.balloons()
                    st.cache_data.clear()
                    # Reset cache last_update agar langsung fresh setelah scrape
                    for _k in list(st.session_state.keys()):
                        if _k.startswith("_lu_ts_"):
                            st.session_state[_k] = 0
                    import time; time.sleep(2)
                    st.rerun()

    with c_stop:
        if st.button("⛔ STOP", type="secondary"):
            with open(STOP_FILE, "w") as f: f.write("STOP")
            st.toast("Stop signal sent.")

    # Tampilkan hasil scraping terakhir (persisten setelah selesai)
    if "hasil_scrape_terakhir" in st.session_state:
        h = st.session_state["hasil_scrape_terakhir"]
        with st.expander(f"📋 Hasil Scraping Terakhir — {h['waktu']} | Tabel: {h['tabel']} | Mode: {h['mode']}", expanded=True):
            for baris in h["hasil"]:
                st.markdown(baris)

# ============================================================
# TAB 2: DASHBOARD
# ============================================================
with tab_dashboard:
    st.subheader("📊 Dashboard Data Tender")

    fc1, fc2, fc3, fc4 = st.columns([1, 2, 2, 2])
    with fc1:
        dash_tahun = st.number_input("Tahun", min_value=2020, max_value=2030,
                                      value=datetime.now().year, key="dash_tahun")
    with fc2:
        dash_kat = st.selectbox("Tabel", ["tender", "non_tender", "pencatatan"], key="dash_kat")
    with fc3:
        dash_instansi = st.selectbox("Instansi", ["Semua"] + [l["nama"] for l in DAFTAR_LPSE], key="dash_inst")
    with fc4:
        dash_search = st.text_input("Cari nama paket...", key="dash_search")

    @st.cache_data(ttl=60)
    def get_daftar_tahapan(tabel):
        """Ambil daftar nilai unik tahapan dari DB untuk filter dropdown."""
        try:
            r = _sb().table(tabel).select("tahapan").execute()
            vals = sorted({row["tahapan"] for row in r.data if row.get("tahapan")})
            return vals
        except:
            return []

    daftar_tahapan = get_daftar_tahapan(dash_kat)
    dash_tahapan = st.selectbox(
        "Filter Tahapan",
        ["Semua"] + daftar_tahapan,
        key="dash_tahapan"
    )

    @st.cache_data(ttl=300)
    def load_dashboard(tabel, tahun, instansi_filter, tahapan_filter, search):
        try:
            q = _sb().table(tabel).select("*").ilike("tahun_anggaran", f"%{tahun}%")
            if instansi_filter != "Semua":
                q = q.ilike("instansi", f"%{instansi_filter}%")
            if tahapan_filter != "Semua":
                q = q.eq("tahapan", tahapan_filter)
            if search:
                q = q.ilike("nama_paket", f"%{search}%")
            r = q.order("diambil_pada", desc=True).limit(500).execute()
            return pd.DataFrame(r.data) if r.data else pd.DataFrame()
        except Exception as e:
            return pd.DataFrame()

    df = load_dashboard(dash_kat, dash_tahun, dash_instansi, dash_tahapan, dash_search)

    if not df.empty:
        m1, m2, m3, m4 = st.columns(4)
        total = len(df)
        selesai = len(df[df["tahapan"].str.contains("Selesai", na=False)])
        berkontrak = len(df[df["pemenang_berkontrak"].ne("Belum Ada Kontrak") & df["pemenang_berkontrak"].notna()])
        ada_pemenang = len(df[df["nama_pemenang"].ne("Belum Ada Pemenang") & df["nama_pemenang"].notna()])
        m1.metric("Total Paket", total)
        m2.metric("Tender Selesai", selesai)
        m3.metric("Ada Pemenang", ada_pemenang)
        m4.metric("Sudah Berkontrak", berkontrak)

        st.write("---")

        kolom_tampil = ["kode_tender", "nama_paket", "instansi", "tahapan",
                        "nilai_hps", "nama_pemenang", "harga_kontrak", "kontrak_mulai", "link_detail"]
        kolom_ada = [c for c in kolom_tampil if c in df.columns]
        df_tampil = df[kolom_ada].copy()

        st.dataframe(
            df_tampil,
            use_container_width=True,
            height=400,
            column_config={
                "link_detail": st.column_config.LinkColumn("Link Detail", display_text="Buka"),
                "nama_paket":  st.column_config.TextColumn("Nama Paket", width="large"),
                "nilai_hps":   st.column_config.TextColumn("Nilai HPS"),
                "harga_kontrak": st.column_config.TextColumn("Harga Kontrak"),
            }
        )

        st.write("---")
        ex1, ex2 = st.columns(2)
        with ex1:
            csv_buf = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Download CSV", csv_buf,
                               file_name=f"spse_{dash_kat}_{dash_tahun}.csv",
                               mime="text/csv")
        with ex2:
            try:
                xl_buf = BytesIO()
                df.to_excel(xl_buf, index=False, engine="openpyxl")
                xl_buf.seek(0)
                st.download_button("⬇️ Download Excel", xl_buf.getvalue(),
                                   file_name=f"spse_{dash_kat}_{dash_tahun}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except:
                st.caption("Excel export butuh openpyxl")
    else:
        st.info("Belum ada data. Jalankan scraper dulu di tab Scraper.")

    if st.button("🔄 Refresh Dashboard"):
        st.cache_data.clear()
        st.rerun()

# ============================================================
# TAB 3: LOG & STATUS
# ============================================================
with tab_log:
    st.subheader("📋 Log Scraper & Status Task Scheduler")

    st.markdown("**Status Task Scheduler (`POKJA_ScrapeSpse`)**")
    try:
        ps_cmd = (
            "Get-ScheduledTaskInfo -TaskName 'POKJA_ScrapeSpse' | "
            "Select-Object LastRunTime, NextRunTime, LastTaskResult | "
            "ConvertTo-Json"
        )
        result = subprocess.run(
            ["powershell", "-Command", ps_cmd],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.returncode == 0 and result.stdout.strip():
            info = json.loads(result.stdout)
            sc1, sc2, sc3, sc4 = st.columns([2, 2, 2, 1])
            last_result = info.get("LastTaskResult", -1)
            status_label = "✅ Sukses" if last_result == 0 else f"❌ Error (kode {last_result})"

            def parse_ps_date(raw):
                """Parse /Date(1234567890123+0700)/ dari PowerShell JSON ke string WIB."""
                if not raw: return "-"
                m = re.search(r"/Date\((\d+)", str(raw))
                if m:
                    ts = int(m.group(1)) / 1000  # ms → detik
                    dt = datetime.utcfromtimestamp(ts) + timedelta(hours=7)
                    return dt.strftime("%d/%m/%Y %H:%M")
                return str(raw)[:16]

            sc1.metric("Last Run", parse_ps_date(info.get("LastRunTime")))
            sc2.metric("Next Run", parse_ps_date(info.get("NextRunTime")))
            sc3.metric("Status", status_label)
            with sc4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("▶️ Jalankan Sekarang", use_container_width=True):
                    try:
                        subprocess.run(
                            ["powershell", "-Command", "Start-ScheduledTask -TaskName 'POKJA_ScrapeSpse'"],
                            capture_output=True, text=True, timeout=10,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                        st.toast("✅ Task Scheduler dijalankan!", icon="🚀")
                    except Exception as ex:
                        st.error(f"Gagal trigger: {ex}")
        else:
            st.warning("Task Scheduler tidak ditemukan atau tidak bisa dibaca.")
    except Exception as e:
        st.warning(f"Tidak bisa baca Task Scheduler: {e}")

    st.write("---")

    st.markdown("**Log Scraper (100 baris terakhir)**")
    try:
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, encoding="utf-8") as f:
                semua_baris = f.readlines()
            # Filter baris noise dari httpx (HTTP Request: POST/GET ke supabase/spse)
            FILTER_NOISE = ("HTTP Request:", "httpx", "httpcore")
            bersih = [b for b in semua_baris if not any(n in b for n in FILTER_NOISE)]
            tail = "".join(bersih[-100:])
            st.code(tail, language=None)
        else:
            st.info("Log belum ada.")
    except Exception as e:
        st.error(f"Gagal baca log: {e}")

    if st.button("🔄 Refresh Log"):
        st.rerun()
