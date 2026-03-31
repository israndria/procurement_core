import streamlit as st
import urllib.request
import urllib.parse
import http.cookiejar
import json
import re
import logging
import subprocess
import pandas as pd
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
_dir1 = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
_env_path = os.path.join(_dir1, "secret_supabase.env")
if not os.path.exists(_env_path):
    _env_path = os.path.join(os.getcwd(), "secret_supabase.env")
load_dotenv(dotenv_path=_env_path)
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

CONFIG_KATEGORI = {
    "Tender":     {"tabel": "tender",     "endpoint": "lelang"},
    "Non Tender": {"tabel": "non_tender", "endpoint": "nontender"},
    "Pencatatan": {"tabel": "pencatatan", "endpoint": "pencatatan"},
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
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
except:
    st.error("Koneksi Supabase Gagal.")
    st.stop()

# --- 3. UTILITY ---
def toggle_all():
    state = st.session_state.chk_all
    for lpse in DAFTAR_LPSE:
        st.session_state[f"chk_{lpse['kode']}"] = state

def parse_wib(tgl_raw):
    tgl_raw = tgl_raw.split('+')[0].replace("Z", "")
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(tgl_raw, fmt) + timedelta(hours=7)
        except: pass
    return None

def get_last_update(kode_lpse, kategori_selected):
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
    import time
    url = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    html = fetch_html(opener, url)
    time.sleep(2)
    m = re.search(r"authenticityToken = '([a-f0-9]+)'", html)
    return m.group(1) if m else ""

def get_list_paket(opener, token, kode_lpse, endpoint, tahun):
    import time
    url = f"{BASE_URL}/{kode_lpse}/dt/{endpoint}?tahun={tahun}"
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

def get_detail_paket(opener, kode_lpse, endpoint, kode_tender):
    base_lelang   = f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}"
    base_evaluasi = f"{BASE_URL}/{kode_lpse}/evaluasi/{kode_tender}"
    referer = f"{BASE_URL}/{kode_lpse}/{endpoint}"
    detail = {}
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

def scrape_satu_lpse(target, tahun_pilihan, kategori_pilihan, supabase_client):
    if os.path.exists(STOP_FILE): return "⛔ Dibatalkan"
    config = CONFIG_KATEGORI[kategori_pilihan]
    tabel = config["tabel"]; endpoint = config["endpoint"]
    nama_lpse = target["nama"]; kode_lpse = target["kode"]
    log.info(f"Mulai scrape: {nama_lpse} ({kategori_pilihan})")
    try:
        opener = buat_opener()
        token  = get_session(opener, kode_lpse, endpoint)
        rows   = get_list_paket(opener, token, kode_lpse, endpoint, tahun_pilihan)
        log.info(f"  {nama_lpse}: {len(rows)} paket")
    except Exception as e:
        log.error(f"  {nama_lpse}: Gagal ambil list — {e}")
        return f"❌ {nama_lpse}: Gagal ambil list"

    ok = 0; err = 0
    for row in rows:
        if os.path.exists(STOP_FILE): return f"🛑 {nama_lpse}: STOP"
        try:
            kode_tender = str(row[0]).strip()
            nama_paket  = strip_html(row[1])
            instansi    = strip_html(row[2])
            tahapan     = strip_html(row[3])
            if not nama_paket or nama_paket == "nan": continue
            detail = get_detail_paket(opener, kode_lpse, endpoint, kode_tender)
            data = {
                "kode_tender": kode_tender, "nama_paket": nama_paket,
                "instansi": instansi, "tahapan": tahapan,
                "link_detail": f"{BASE_URL}/{kode_lpse}/{endpoint}/{kode_tender}/pengumumanlelang",
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
            ok += 1
        except Exception as e:
            err += 1
            log.warning(f"  Error {row[0] if row else '?'}: {str(e)[:150]}")
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
# TAB 1: SCRAPER (existing)
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

    target_dipilih = []
    for lpse in DAFTAR_LPSE:
        col_a, col_b, col_c = st.columns([0.5, 4, 3])
        if f"chk_{lpse['kode']}" not in st.session_state:
            st.session_state[f"chk_{lpse['kode']}"] = False
        if col_a.checkbox("", key=f"chk_{lpse['kode']}"): target_dipilih.append(lpse)
        col_b.markdown(lpse['nama'])
        col_c.markdown(get_last_update(lpse['kode'], kategori_input))

    st.write("---")
    col_opts, _ = st.columns([2, 3])
    with col_opts:
        max_workers_input = st.slider("Parallel Workers", min_value=1, max_value=5, value=2)
    st.info("ℹ️ V2.1 menggunakan urllib (tanpa Chrome).")

    c_start, c_stop = st.columns([4, 1])
    with c_start:
        if st.button(f"🚀 GAS SCRAPE ({len(target_dipilih)})", type="primary", use_container_width=True):
            if os.path.exists(STOP_FILE): os.remove(STOP_FILE)
            if not target_dipilih:
                st.warning("Pilih daerah!")
            else:
                # Simpan snapshot kode_tender sebelum scrape (untuk deteksi paket baru)
                tabel_aktif = CONFIG_KATEGORI[kategori_input]["tabel"]
                try:
                    snap = supabase.table(tabel_aktif).select("kode_tender").execute()
                    st.session_state["snapshot_sebelum"] = {r["kode_tender"] for r in snap.data}
                except: st.session_state["snapshot_sebelum"] = set()

                status_box = st.status(
                    f"Menghisap data ke tabel: **{CONFIG_KATEGORI[kategori_input]['tabel']}**...",
                    expanded=True)
                p_bar = status_box.progress(0)
                with ThreadPoolExecutor(max_workers=max_workers_input) as executor:
                    futures = [executor.submit(scrape_satu_lpse, t, tahun_input, kategori_input, supabase)
                               for t in target_dipilih]
                    done = 0
                    for f in futures:
                        res = f.result(); done += 1
                        p_bar.progress(int((done / len(futures)) * 100))
                        status_box.write(res)

                if not os.path.exists(STOP_FILE):
                    # Deteksi paket baru
                    try:
                        snap_sesudah = supabase.table(tabel_aktif).select("kode_tender,nama_paket,instansi").execute()
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
                    import time; time.sleep(2)
                    st.rerun()

    with c_stop:
        if st.button("⛔ STOP", type="secondary"):
            with open(STOP_FILE, "w") as f: f.write("STOP")
            st.toast("Stop signal sent.")

# ============================================================
# TAB 2: DASHBOARD
# ============================================================
with tab_dashboard:
    st.subheader("📊 Dashboard Data Tender")

    # Filter row
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

    @st.cache_data(ttl=300)
    def load_dashboard(tabel, tahun, instansi_filter, search):
        try:
            q = supabase.table(tabel).select("*").ilike("tahun_anggaran", f"%{tahun}%")
            if instansi_filter != "Semua":
                q = q.ilike("instansi", f"%{instansi_filter}%")
            if search:
                q = q.ilike("nama_paket", f"%{search}%")
            r = q.order("diambil_pada", desc=True).limit(500).execute()
            return pd.DataFrame(r.data) if r.data else pd.DataFrame()
        except Exception as e:
            return pd.DataFrame()

    df = load_dashboard(dash_kat, dash_tahun, dash_instansi, dash_search)

    if not df.empty:
        # Metrics
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

        # Tabel
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

        # Export
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

    # Status Task Scheduler
    st.markdown("**Status Task Scheduler (`POKJA_ScrapeSpse`)**")
    try:
        ps_cmd = (
            "Get-ScheduledTaskInfo -TaskName 'POKJA_ScrapeSpse' | "
            "Select-Object LastRunTime, NextRunTime, LastTaskResult | "
            "ConvertTo-Json"
        )
        result = subprocess.run(
            ["powershell", "-Command", ps_cmd],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and result.stdout.strip():
            info = json.loads(result.stdout)
            sc1, sc2, sc3 = st.columns(3)
            last_run = info.get("LastRunTime", "-")
            next_run = info.get("NextRunTime", "-")
            last_result = info.get("LastTaskResult", -1)
            status_label = "✅ Sukses" if last_result == 0 else f"❌ Error (kode {last_result})"
            sc1.metric("Last Run", str(last_run)[:16] if last_run else "-")
            sc2.metric("Next Run", str(next_run)[:16] if next_run else "-")
            sc3.metric("Status", status_label)
        else:
            st.warning("Task Scheduler tidak ditemukan atau tidak bisa dibaca.")
    except Exception as e:
        st.warning(f"Tidak bisa baca Task Scheduler: {e}")

    st.write("---")

    # Live Log
    st.markdown("**Log Scraper (100 baris terakhir)**")
    try:
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, encoding="utf-8") as f:
                semua_baris = f.readlines()
            tail = "".join(semua_baris[-100:])
            st.code(tail, language=None)
        else:
            st.info("Log belum ada.")
    except Exception as e:
        st.error(f"Gagal baca log: {e}")

    if st.button("🔄 Refresh Log"):
        st.rerun()
