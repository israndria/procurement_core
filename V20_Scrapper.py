import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import logging
from supabase import create_client
import os
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

# --- 1. KONFIGURASI HALAMAN ---
if not os.environ.get("POKJA_MULTIPAGE"):
    st.set_page_config(page_title="SPSE Scraper V1.c", page_icon="🏗️", layout="wide")

st.markdown("""
    <style>
        .stCheckbox { margin-bottom: -15px; }
        .block-container { padding-top: 1rem; padding-bottom: 2rem; }
        h1 { padding-top: 0rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. SETUP DATA ---
# Cari secret_supabase.env: coba __file__ dulu, fallback ke CWD
_dir1 = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
_env_path = os.path.join(_dir1, "secret_supabase.env")
if not os.path.exists(_env_path):
    _env_path = os.path.join(os.getcwd(), "secret_supabase.env")
load_dotenv(dotenv_path=_env_path)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
STOP_FILE = os.path.join(os.path.dirname(_env_path), "stop_signal.txt")
LOG_FILE = os.path.join(os.path.dirname(_env_path), "scraper.log")

# Setup logging ke file
logging.basicConfig(
    filename=LOG_FILE, level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S", encoding="utf-8"
)
log = logging.getLogger("scraper")

# --- PENTING: MAPPING NAMA TABEL DI SINI ---
# Pastikan nama tabel di Supabase kamu sesuai dengan yang ada di kanan (nilai dict)
CONFIG_KATEGORI = {
    "Tender": {
        "tabel": "tender",        # Nama tabel di Supabase
        "endpoint": "lelang"      # URL di website SPSE
    },
    "Non Tender": {
        "tabel": "non_tender",    # Nama tabel di Supabase (biasanya pake underscore)
        "endpoint": "nontender"   # URL di website SPSE
    },
    "Pencatatan": {
        "tabel": "pencatatan",    # Nama tabel di Supabase
        "endpoint": "pencatatan"  # URL di website SPSE
    }
}

DAFTAR_LPSE = [
    {"nama": "Kabupaten Tapin", "kode": "tapinkab"},
    {"nama": "Kota Banjarmasin", "kode": "banjarmasinkota"},
    {"nama": "Kota Banjarbaru", "kode": "banjarbarukota"},
    {"nama": "Kabupaten Tanah Laut", "kode": "tanahlautkab"},
    {"nama": "Kabupaten Barito Kuala", "kode": "baritokualakab"},
    {"nama": "Kabupaten Banjar", "kode": "banjarkab"},
    {"nama": "Kabupaten Tanah Bumbu", "kode": "tanahbumbukab"},
    {"nama": "Kabupaten Kotabaru", "kode": "kotabarukab"},
    {"nama": "Kabupaten HSS", "kode": "hulusungaiselatankab"},
    {"nama": "Kabupaten HST", "kode": "hstkab"},
    {"nama": "Kabupaten HSU", "kode": "hsu"},
    {"nama": "Kabupaten Balangan", "kode": "balangankab"},
    {"nama": "Kabupaten Tabalong", "kode": "tabalongkab"},
    {"nama": "Provinsi Kalsel", "kode": "kalselprov"}
]

try:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
except:
    st.error("Koneksi Supabase Gagal.")
    st.stop()

# --- 3. UTILITY ---
def toggle_all():
    state = st.session_state.v20_chk_all
    for lpse in DAFTAR_LPSE:
        st.session_state[f"v20_chk_{lpse['kode']}"] = state

def get_last_update(kode_lpse, kategori_selected):
    # Ambil nama tabel target berdasarkan pilihan user
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
            try: waktu_utc = datetime.strptime(tgl_raw, "%Y-%m-%dT%H:%M:%S.%f")
            except: 
                try: waktu_utc = datetime.strptime(tgl_raw, "%Y-%m-%dT%H:%M:%S")
                except: return f"⚠️ {tgl_raw[:10]}"
            waktu_wib = waktu_utc + timedelta(hours=7)
            return f"🟢 {waktu_wib.strftime('%d/%m %H:%M')}"
        return "⚪ Kosong"
    except: return "-"

# --- 4. ENGINE SCRAPING (MULTI-TABLE) ---
def create_driver(headless=False, minimize=False):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
    elif minimize:
        # Off-screen: Chrome tidak pernah muncul di layar
        options.add_argument("--window-position=-2400,-2400")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    return webdriver.Chrome(options=options)

def scrape_satu_lpse(target, tahun_pilihan, kategori_pilihan, supabase_client, headless=False, minimize=False):
    if os.path.exists(STOP_FILE): return "⛔ Dibatalkan"

    config = CONFIG_KATEGORI[kategori_pilihan]
    nama_tabel_target = config["tabel"]
    endpoint_target = config["endpoint"]

    driver = None
    nama_lpse = target["nama"]
    kode_lpse = target["kode"]

    def keep_minimized():
        """Re-minimize Chrome setelah switch_to agar tidak muncul ke depan"""
        if minimize and not headless:
            try: driver.minimize_window()
            except: pass

    url = f"https://spse.inaproc.id/{kode_lpse}/{endpoint_target}?tahun={tahun_pilihan}"
    log.info(f"Mulai scrape: {nama_lpse} ({kategori_pilihan})")

    MAX_RETRY = 2
    for attempt in range(MAX_RETRY + 1):
      try:
        if driver: driver.quit(); driver = None
        driver = create_driver(headless=headless, minimize=minimize)
        wait = WebDriverWait(driver, 15)
        driver.get(url)
        time.sleep(4)

        # Cek apakah halaman berhasil dimuat
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table")))
        except:
            if attempt < MAX_RETRY:
                log.warning(f"  Retry {attempt+1}/{MAX_RETRY} untuk {nama_lpse} (halaman tidak termuat)")
                continue
            else:
                log.error(f"  Gagal load halaman {nama_lpse} setelah {MAX_RETRY} retry")
                return f"❌ {nama_lpse}: Gagal load halaman"

        halaman = 1
        paket_counter = 0
        error_counter = 0
        halaman_kosong = 0

        while True:
            if os.path.exists(STOP_FILE): driver.quit(); return f"🛑 {nama_lpse}: STOP"

            try:
                driver.switch_to.window(driver.window_handles[0])
                keep_minimized()
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table")))
                rows = driver.find_elements(By.CSS_SELECTOR, "table.table tbody tr")
            except: break 
            
            paket_ditemukan = False
            
            for i, row in enumerate(rows):
                if os.path.exists(STOP_FILE): driver.quit(); return "STOP"

                try:
                    row_class = row.get_attribute("class") or ""
                    if "data-td-text" in row_class: continue
                    links = row.find_elements(By.TAG_NAME, "a")
                    if len(links) == 0: continue 
                    
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) < 2: continue

                    kode_tender = cols[0].text.strip()
                    link_elem = links[0] 
                    nama_paket = link_elem.text.strip()
                    link_detail = link_elem.get_attribute("href")
                    
                    if not nama_paket: continue

                    instansi = cols[2].text.strip() if len(cols) > 2 else "-"
                    tahapan = cols[3].text.strip() if len(cols) > 3 else "-"
                    
                    paket_ditemukan = True
                    
                    # Buka Detail
                    driver.execute_script("window.open(arguments[0]);", link_detail)
                    driver.switch_to.window(driver.window_handles[-1])
                    keep_minimized()
                    
                    # Logic Ambil Data Detail
                    jenis_pengadaan = "-"
                    satuan_kerja = "-"
                    nilai_pagu = "0"
                    tahun_anggaran = "-"
                    nilai_hps = "0"
                    metode_pengadaan = "-"
                    
                    try:
                        wait_detail = WebDriverWait(driver, 5)
                        try: jenis_pengadaan = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Jenis Pengadaan')]/following-sibling::td"))).text.strip()
                        except: pass
                        try: satuan_kerja = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Satuan Kerja')]/following-sibling::td"))).text.strip()
                        except: pass
                        try: nilai_pagu = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Nilai Pagu Paket')]/following-sibling::td"))).text.strip()
                        except: pass
                        try: tahun_anggaran = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Tahun Anggaran')]/following-sibling::td"))).text.strip()
                        except: pass
                        try: nilai_hps = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Nilai HPS')]/following-sibling::td"))).text.strip()
                        except: pass
                        try: metode_pengadaan = wait_detail.until(EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Metode Pengadaan')]/following-sibling::td"))).text.strip()
                        except: pass
                    except: pass
                    
                    nama_pemenang = "Belum Ada Pemenang"
                    pemenang_berkontrak = "Belum Ada Kontrak"
                    alamat = "-"
                    harga_kontrak = "0"

                    # Logic Pemenang (Dinamis: Cek dulu apakah tabnya ada, baru klik)
                    try:
                        pemenang_tabs = driver.find_elements(By.XPATH, "//a[text()='Pemenang']")
                        if len(pemenang_tabs) > 0:
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Pemenang']"))).click()
                            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='main']/div/table")))
                            try:
                                elemen_td = driver.find_element(By.XPATH, "//*[@id='main']/div/table/tbody/tr[7]/td/table/tbody/tr[2]/td[1]")
                                nama_pemenang = re.sub(r"[^A-Za-z0-9\s\.,&()\-]", "", elemen_td.text.strip())
                            except: pass
                    except: pass

                    try:
                        if len(driver.find_elements(By.XPATH, "//a[contains(text(), 'Pemenang Berkontrak')]")) > 0:
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Pemenang Berkontrak')]"))).click()
                            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='main']/div/table")))
                            try:
                                elemen_td = driver.find_element(By.XPATH, "//*[@id='main']/div/table/tbody/tr[7]/td/table/tbody/tr[2]/td[1]")
                                pemenang_berkontrak = re.sub(r"[^A-Za-z0-9\s\.,&()\-]", "", elemen_td.text.strip())
                                elemen_td_alamat = driver.find_element(By.XPATH, "//*[@id='main']/div/table/tbody/tr[7]/td/table/tbody/tr[2]/td[2]")
                                alamat = re.sub(r"[^A-Za-z0-9\s\.,&()\-]", "", elemen_td_alamat.text.strip())
                                elemen_td_harga = driver.find_element(By.XPATH, "//*[@id='main']/div/table/tbody/tr[7]/td/table/tbody/tr[2]/td[4]")
                                harga_kontrak = re.sub(r"[^A-Za-z0-9\s\.,&()\-]", "", elemen_td_harga.text.strip())
                            except: pass
                    except: pass

                    kontrak_mulai, kontrak_selesai = "", ""
                    try:
                        driver.execute_script("window.open(arguments[0]);", f"https://spse.inaproc.id/{kode_lpse}/{endpoint_target}/{kode_tender}/jadwal")
                        driver.switch_to.window(driver.window_handles[-1])
                        keep_minimized()
                        wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
                        rows_jadwal = driver.find_elements(By.XPATH, "//table/tbody/tr")
                        for tr in rows_jadwal:
                            kolom = tr.find_elements(By.TAG_NAME, "td")
                            if len(kolom) >= 5 and "Penandatanganan Kontrak" in kolom[1].text:
                                kontrak_mulai, kontrak_selesai = kolom[2].text.strip(), kolom[3].text.strip()
                                break
                    except:
                        kontrak_mulai, kontrak_selesai = "-", "-"
                    finally:
                        # Tutup semua tab kecuali tab utama
                        while len(driver.window_handles) > 1:
                            driver.switch_to.window(driver.window_handles[-1])
                            driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        keep_minimized()

                    # DATA UNTUK DISIMPAN
                    data = {
                        "kode_tender": kode_tender, "nama_paket": nama_paket, "instansi": instansi,
                        "tahapan": tahapan, "link_detail": link_detail, 
                        "jenis_pengadaan": jenis_pengadaan, "satuan_kerja": satuan_kerja, 
                        "nilai_pagu": nilai_pagu, "tahun_anggaran": tahun_anggaran,
                        "nilai_hps": nilai_hps, "metode_pengadaan": metode_pengadaan,
                        "nama_pemenang": nama_pemenang, "pemenang_berkontrak": pemenang_berkontrak,
                        "alamat": alamat, "harga_kontrak": harga_kontrak,
                        "kontrak_mulai": kontrak_mulai, "kontrak_selesai": kontrak_selesai,
                        "diambil_pada": datetime.now().isoformat()
                    }
                    
                    # === DISIMPAN KE TABEL SESUAI KATEGORI ===
                    supabase_client.table(nama_tabel_target).upsert(data).execute()
                    paket_counter += 1
                    
                except Exception as e:
                    error_counter += 1
                    log.warning(f"  Error paket {kode_tender if 'kode_tender' in dir() else '?'}: {str(e)[:100]}")
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    continue

            if not paket_ditemukan: halaman_kosong += 1
            else: halaman_kosong = 0
            
            if halaman_kosong >= 2: break
            
            try:
                next_btn = driver.find_element(By.XPATH, "//a[contains(text(),'Berikutnya')] | //a[contains(@class,'next')]")
                if "disabled" in (next_btn.get_attribute("class") or ""): break
                next_btn.click()
                halaman += 1
                time.sleep(3)
            except: break
        
        msg = f"{nama_lpse}: ✅ {paket_counter} data"
        if error_counter: msg += f", ⚠️ {error_counter} error"
        log.info(msg)
        return msg

      except Exception as e:
        if attempt < MAX_RETRY:
            log.warning(f"  Retry {attempt+1}/{MAX_RETRY} untuk {nama_lpse}: {str(e)[:100]}")
            continue
        log.error(f"❌ {nama_lpse}: {str(e)[:200]}")
        return f"❌ {nama_lpse}: Error Sistem"
      finally:
        if driver:
            driver.quit()
            driver = None

    return f"❌ {nama_lpse}: Gagal setelah {MAX_RETRY} retry"

def main():
    # --- 5. UI ---
    st.title("🏗️ SPSE Scraper V1.c (Multi-Table)")

    # Bagian Kontrol
    col_thn, col_kat = st.columns([1, 2])

    with col_thn:
        tahun_sekarang = datetime.now().year
        tahun_input = st.number_input("Tahun", min_value=2020, max_value=2030, value=tahun_sekarang)

    with col_kat:
        kategori_input = st.radio("Kategori (Target Tabel)", ["Tender", "Non Tender", "Pencatatan"], horizontal=True)

    st.write("---")
    c1, c2, c3 = st.columns([0.5, 4, 3])
    c1.checkbox("", key="v20_chk_all", on_change=toggle_all)
    c2.markdown("**DAERAH**")
    c3.markdown(f"**LAST UPDATE (Tabel: {CONFIG_KATEGORI[kategori_input]['tabel']})**")

    target_dipilih = []
    with st.container():
        for lpse in DAFTAR_LPSE:
            col_a, col_b, col_c = st.columns([0.5, 4, 3])
            with col_a:
                if f"v20_chk_{lpse['kode']}" not in st.session_state: st.session_state[f"v20_chk_{lpse['kode']}"] = False
                if col_a.checkbox("", key=f"v20_chk_{lpse['kode']}"): target_dipilih.append(lpse)
            with col_b: st.markdown(f"{lpse['nama']}")
            with col_c:
                st.markdown(get_last_update(lpse['kode'], kategori_input))

    st.write("---")
    col_opts1, col_opts2, col_opts3 = st.columns(3)
    with col_opts1:
        headless_mode = st.checkbox("🖥️ Headless (SPSE mungkin blokir)", value=False)
    with col_opts2:
        minimize_mode = st.checkbox("🔽 Minimize Chrome", value=True)
    with col_opts3:
        max_workers_input = st.slider("Parallel Workers", min_value=1, max_value=5, value=2)

    c_start, c_stop = st.columns([4, 1])

    with c_start:
        if st.button(f"🚀 GAS SCRAPE ({len(target_dipilih)})", type="primary", use_container_width=True):
            if os.path.exists(STOP_FILE): os.remove(STOP_FILE)
            if not target_dipilih:
                st.warning("Pilih daerah!")
            else:
                status_box = st.status(f"Menghisap data ke tabel: **{CONFIG_KATEGORI[kategori_input]['tabel']}**...", expanded=True)
                p_bar = status_box.progress(0)
                with ThreadPoolExecutor(max_workers=max_workers_input) as executor:
                    futures = [executor.submit(scrape_satu_lpse, t, tahun_input, kategori_input, supabase, headless_mode, minimize_mode) for t in target_dipilih]
                    done = 0
                    for f in futures:
                        res = f.result()
                        done += 1
                        p_bar.progress(int((done/len(futures))*100))
                        status_box.write(res)

                if not os.path.exists(STOP_FILE):
                    status_box.update(label="Selesai!", state="complete", expanded=False)
                    st.balloons()
                    st.cache_data.clear()
                    time.sleep(2)
                    st.rerun()

    with c_stop:
        if st.button("⛔ STOP", type="secondary"):
            with open(STOP_FILE, "w") as f: f.write("STOP")
            st.toast("Stop signal sent.")

if __name__ == "__main__":
    main()