import streamlit as st
import pandas as pd
import datetime
import os
import time
import io
import re
import urllib.request
import hashlib

# --- LIBRARY GOOGLE ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ==========================================
# ⚙️ KONFIGURASI PATH OTOMATIS (PORTABLE) ⚙️
# ==========================================
# Script ini akan mendeteksi folder tempat dia berada
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# File akan selalu dicari di folder yang sama dengan script
DB_FILE_PATH = os.path.join(BASE_DIR, 'database_tender.csv')
CRED_FILE_PATH = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_FILE_PATH = os.path.join(BASE_DIR, 'token.json')

# ==========================================
# 🛑 DAFTAR ANGGOTA POKJA 🛑
# ==========================================
DAFTAR_ANGGOTA = [
    "Mas Boy", "Pak Bambang", "Pak Aidy", "Mas Agit", "Ka Adam",
    "Pak Hernov", "Pak Herry Setiawan", "Pak Rahmad Jabug",
    "Ka Ade", "Pak Ian", "Erwin", "Pak Arya",
    "Mas Budi Rachman", "Pak Azis"
]

# --- KONFIGURASI APLIKASI ---
TAHUN_SEKARANG = datetime.datetime.now().year
JUDUL_APLIKASI = f"Jadwal Tender/Seleksi Tahun Anggaran {TAHUN_SEKARANG}"
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'primary'

st.set_page_config(page_title=JUDUL_APLIKASI, layout="wide")
st.title(f"📅 {JUDUL_APLIKASI}")
st.caption(f"📂 Lokasi Database: {DB_FILE_PATH}") # Info lokasi file (untuk memastikan)

# --- FUNGSI FORMAT TANGGAL INDONESIA ---
BULAN_INDO_MAP = {
    'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04', 'Mei': '05', 'Juni': '06',
    'Juli': '07', 'Agustus': '08', 'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
}
LIST_HARI = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
LIST_BULAN = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']

def get_indonesian_timestamp():
    """Mengembalikan waktu sekarang: Selasa, 3 Februari 2026, 21:18"""
    now = datetime.datetime.now()
    nama_hari = LIST_HARI[now.weekday()]
    nama_bulan = LIST_BULAN[now.month]
    jam_menit = now.strftime("%H:%M")
    return f"{nama_hari}, {now.day} {nama_bulan} {now.year}, {jam_menit}"

def parse_spse_date(date_str):
    try:
        if pd.isna(date_str) or str(date_str).strip() == '-': return None
        clean_str = re.sub(r'\s*\(.*?\)', '', str(date_str))
        parts = clean_str.split(' ')
        if len(parts) >= 3:
            tgl, bln, thn = parts[0].zfill(2), BULAN_INDO_MAP.get(parts[1], '01'), parts[2]
            jam = parts[3] if len(parts) > 3 else "00:00"
            return datetime.datetime.fromisoformat(f"{thn}-{bln}-{tgl}T{jam}:00")
    except: return None

# --- FUNGSI GOOGLE CALENDAR ---
def get_service():
    creds = None
    # Cek token di folder yang sama
    if os.path.exists(TOKEN_FILE_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CRED_FILE_PATH):
                st.error(f"File credentials.json tidak ditemukan di: {CRED_FILE_PATH}")
                st.stop()
            flow = InstalledAppFlow.from_client_secrets_file(CRED_FILE_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE_PATH, 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)

def delete_existing_events_by_source(service, source_url):
    try:
        page_token = None
        while True:
            events_result = service.events().list(
                calendarId=CALENDAR_ID, q=source_url, singleEvents=True, pageToken=page_token
            ).execute()
            events = events_result.get('items', [])
            for event in events:
                if source_url in event.get('description', ''):
                    service.events().delete(calendarId=CALENDAR_ID, eventId=event['id']).execute()
            page_token = events_result.get('nextPageToken')
            if not page_token: break
        return True
    except Exception as e:
        return False

# --- FUNGSI DATABASE (CRUD) ---
def compute_hash(df_jadwal: pd.DataFrame) -> str:
    content = df_jadwal[['Tahap', 'Mulai', 'Sampai', 'Perubahan']].to_csv(index=False)
    return hashlib.md5(content.encode()).hexdigest()

def update_local_database(df_result, hash_map: dict = None):
    new_data = []
    now_str = get_indonesian_timestamp()

    grouped = df_result.groupby('Source')
    for url, group in grouped:
        pokja = group.iloc[0]['Anggota_Pokja']
        nama_pkt = group.iloc[0]['Nama_Paket']
        new_data.append({
            'url': url,
            'members': pokja,
            'nama_paket': nama_pkt,
            'last_sync': now_str,
            'content_hash': hash_map.get(url, '') if hash_map else ''
        })

    df_new = pd.DataFrame(new_data)

    if os.path.exists(DB_FILE_PATH):
        try:
            df_old = pd.read_csv(DB_FILE_PATH)
            for col in ['last_sync', 'nama_paket', 'content_hash']:
                if col not in df_old.columns: df_old[col] = ''

            df_old.set_index('url', inplace=True)
            df_new.set_index('url', inplace=True)
            df_old.update(df_new)
            df_old.reset_index(inplace=True)
            df_new.reset_index(inplace=True)
            df_final = pd.concat([df_old, df_new]).drop_duplicates(subset='url', keep='last')
        except:
            df_final = df_new
    else:
        df_final = df_new

    df_final.to_csv(DB_FILE_PATH, index=False)
    return len(df_final)

def remove_from_database(url_to_remove):
    if os.path.exists(DB_FILE_PATH):
        df = pd.read_csv(DB_FILE_PATH)
        df = df[df['url'] != url_to_remove]
        df.to_csv(DB_FILE_PATH, index=False)
        return True
    return False

def clear_all_database():
    if os.path.exists(DB_FILE_PATH):
        df = pd.DataFrame(columns=['url', 'members', 'nama_paket', 'last_sync'])
        df.to_csv(DB_FILE_PATH, index=False)
        return True
    return False

# --- FUNGSI SCRAPING (urllib — tanpa Chrome/Selenium) ---
def fetch_jadwal_urllib(url: str, members: str) -> pd.DataFrame | None:
    """Ambil tabel jadwal langsung via urllib — ringan, tanpa browser."""
    referer = url.replace('/jadwal', '/pengumuman')
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer':    referer,
        'Accept':     'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            html = resp.read().decode('utf-8', errors='replace')
    except Exception as e:
        st.error(f"Gagal fetch {url}: {e}")
        return None

    tender_id  = url.split('/')[-2]
    nama_final = f"Tender {tender_id}"
    title_m    = re.search(r'\[[\d]+\]\s*(.+?)(?:\s*</title>|\s*$)', html, re.IGNORECASE)
    if title_m: nama_final = title_m.group(1).strip()

    try:
        dfs = pd.read_html(io.StringIO(html), flavor='lxml')
    except Exception as e:
        st.error(f"Gagal parse HTML {url}: {e}")
        return None

    target_df = None
    for t in dfs:
        if any('tahap' in str(c).lower() for c in t.columns):
            target_df = t; break

    if target_df is None:
        st.warning(f"Tabel jadwal tidak ditemukan: {url}")
        return None

    cols      = target_df.columns
    c_tahap   = next((c for c in cols if 'tahap'     in str(c).lower()), None)
    c_mulai   = next((c for c in cols if 'mulai'     in str(c).lower()), None)
    c_sampai  = next((c for c in cols if 'sampai'    in str(c).lower()), None)
    c_ubah    = next((c for c in cols if 'perubahan' in str(c).lower()), None)

    if not c_tahap or not c_mulai: return None

    sel_cols  = [c for c in [c_tahap, c_mulai, c_sampai, c_ubah] if c]
    df_clean  = target_df[sel_cols].copy()
    df_clean.columns = ['Tahap', 'Mulai'] + (['Sampai'] if c_sampai else []) + (['Perubahan'] if c_ubah else [])
    if 'Sampai'    not in df_clean.columns: df_clean['Sampai']    = df_clean['Mulai']
    if 'Perubahan' not in df_clean.columns: df_clean['Perubahan'] = '0'
    df_clean['Perubahan']    = df_clean['Perubahan'].fillna('0')
    df_clean['Source']       = url
    df_clean['Nama_Paket']   = nama_final
    df_clean['Anggota_Pokja'] = members
    df_clean = df_clean.dropna(subset=['Tahap'])
    return df_clean


def start_driver_and_process_slots(slot_data):
    """Proses semua slot URL — tanpa Chrome, pakai urllib."""
    all_data   = []
    hash_map   = {}
    status_text = st.empty()
    progress_bar = st.progress(0)
    total = len(slot_data)

    for i, item in enumerate(slot_data):
        url     = item['url']
        members = item['members']
        tender_id = url.split('/')[-2]

        status_text.write(f"🔄 ({i+1}/{total}) Memproses: {tender_id}...")
        progress_bar.progress(i / total)

        df = fetch_jadwal_urllib(url, members)
        if df is not None:
            all_data.append(df)
            hash_map[url] = compute_hash(df)

    progress_bar.progress(1.0)
    status_text.success("✅ Selesai!")
    return all_data, hash_map

# --- FORMATTING OUTPUT ---
def format_desc(url, perubahan_val, pokja_names):
    try: jml = int(float(str(perubahan_val).replace(',','.'))) 
    except: jml = 0
    status = f"⚠️ PERINGATAN: {jml}x Perubahan" if jml > 0 else "✅ Aman"
    pokja = f"\n\n👥 POKJA:\n{pokja_names}" if pokja_names else ""
    return f"🔗 Link: {url}\n\n{status}{pokja}"

def get_reminders(judul_tahap):
    judul_lower = str(judul_tahap).lower()
    if any(x in judul_lower for x in ['pembuktian', 'pembukaan', 'penunjukan']):
        return {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 24 * 60}]}
    elif 'penjelasan' in judul_lower:
        return {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 3 * 60}]}
    else:
        return {'useDefault': True}

# --- FUNGSI EKSEKUSI UTAMA ---
def run_single_update(url_target, members_target):
    slot_data = [{'url': url_target, 'members': members_target}]
    res, hash_map = start_driver_and_process_slots(slot_data)
    if res:
        df_res = pd.concat(res, ignore_index=True)
        update_local_database(df_res, hash_map)
        svc = get_service()
        for url, grp in df_res.groupby('Source'):
            delete_existing_events_by_source(svc, url)
            for _, r in grp.iterrows():
                ds = parse_spse_date(r['Mulai'])
                de = parse_spse_date(r['Sampai'])
                if ds:
                    if not de: de = ds + datetime.timedelta(hours=1)
                    evt = {
                        'summary': f"{r['Tahap']} - {r['Nama_Paket']}",
                        'description': format_desc(r['Source'], r['Perubahan'], r['Anggota_Pokja']),
                        'start': {'dateTime': ds.isoformat(), 'timeZone': 'Asia/Jakarta'},
                        'end': {'dateTime': de.isoformat(), 'timeZone': 'Asia/Jakarta'},
                        'reminders': get_reminders(r['Tahap'])
                    }
                    try: svc.events().insert(calendarId=CALENDAR_ID, body=evt).execute()
                    except: pass
        return True
    return False

# ============================================
# 🖥️ UI TAB SYSTEM (STREAMLIT)
# ============================================
tab_monitor, tab_tambah = st.tabs(["👀 Pantau & Update", "➕ Tambah Paket"])

# --- TAB 1: DASHBOARD ---
with tab_monitor:
    col_reset_1, col_reset_2 = st.columns([8, 2])
    with col_reset_2:
        if st.button("🗑️ Reset Database", type="primary", use_container_width=True, help="Hapus semua data (Reset Tahun)"):
            if clear_all_database():
                st.toast("Database berhasil direset!", icon="🗑️")
                time.sleep(1)
                st.rerun()

    if os.path.exists(DB_FILE_PATH):
        try:
            df_saved = pd.read_csv(DB_FILE_PATH)
        except:
            df_saved = pd.DataFrame()
        
        if not df_saved.empty:
            if 'last_sync' not in df_saved.columns: df_saved['last_sync'] = "-"
            if 'nama_paket' not in df_saved.columns: df_saved['nama_paket'] = df_saved['url']
            
            # HEADER (Sticky Logic akan dihandle oleh container di bawah)
            h_no, h_info, h_sync, h_act = st.columns([0.5, 6.0, 2.0, 1.5]) 
            h_no.markdown("**No**")
            h_info.markdown("**Nama Paket**")
            h_sync.markdown("**Terakhir Update**") 
            h_act.markdown("**Aksi**")
            st.divider()

            # CONTAINER SCROLLABLE (Sticky Header Effect)
            with st.container(height=500):
                for index, row in df_saved.iterrows():
                    c_no, c_info, c_sync, c_act = st.columns([0.5, 6.0, 2.0, 1.5])
                    
                    c_no.write(f"{index + 1}")
                    
                    with c_info:
                        st.write(f"**{row['nama_paket']}**")
                        st.caption(f"👥 {row['members']}")
                    
                    c_sync.caption(row['last_sync'])
                    
                    with c_act:
                        b_col1, b_col2 = st.columns(2)
                        # Tombol UPDATE (Icon)
                        if b_col1.button("🔄", key=f"upd_{index}", help="Update Data ke Calendar"):
                            with st.spinner("Processing..."):
                                if run_single_update(row['url'], row['members']):
                                    st.toast("Berhasil Update!", icon="✅")
                                    time.sleep(1)
                                    st.rerun()
                                else: st.error("Gagal")
                        
                        # Tombol DELETE (Icon)
                        if b_col2.button("🗑️", key=f"del_{index}", help="Hapus Paket"):
                            if remove_from_database(row['url']):
                                st.rerun()
                    
                    st.divider()
        else:
            st.info("Belum ada paket yang dipantau. Silakan masuk ke Tab 'Tambah Paket'.")
    else:
        st.info("Database baru akan dibuat otomatis setelah Anda menambah paket.")

# --- TAB 2: TAMBAH PAKET ---
with tab_tambah:
    if 'num_slots' not in st.session_state: st.session_state['num_slots'] = 1
    slot_inputs = []

    for i in range(st.session_state['num_slots']):
        c1, c2 = st.columns([3, 2])
        with c1: url_in = st.text_input(f"Link Paket #{i+1}", key=f"u{i}", placeholder="https://lpse.../jadwal/...")
        with c2: pokja_in = st.multiselect(f"Anggota Pokja #{i+1}", DAFTAR_ANGGOTA, key=f"p{i}")
        if url_in.strip(): slot_inputs.append({'url': url_in.strip(), 'members': ", ".join(pokja_in)})

    c_btn1, c_btn2 = st.columns([1, 1])
    if c_btn1.button("➕ Tambah Baris"): st.session_state['num_slots'] += 1; st.rerun()
    if c_btn2.button("🔄 Reset Form"): st.session_state['num_slots'] = 1; st.rerun()

    st.divider()

    if st.button("🚀 EKSEKUSI PAKET BARU", type="primary", use_container_width=True):
        if slot_inputs:
            with st.status("Sedang Bekerja...", expanded=True) as s:
                res, hash_map = start_driver_and_process_slots(slot_inputs)
                if res:
                    df_res = pd.concat(res, ignore_index=True)
                    update_local_database(df_res, hash_map)

                    svc = get_service()
                    for url, grp in df_res.groupby('Source'):
                        s.write(f"📅 Memproses: {grp.iloc[0]['Nama_Paket']}")
                        delete_existing_events_by_source(svc, url)
                        for _, r in grp.iterrows():
                            ds = parse_spse_date(r['Mulai'])
                            de = parse_spse_date(r['Sampai'])
                            if ds:
                                if not de: de = ds + datetime.timedelta(hours=1)
                                evt = {
                                    'summary': f"{r['Tahap']} - {r['Nama_Paket']}",
                                    'description': format_desc(r['Source'], r['Perubahan'], r['Anggota_Pokja']),
                                    'start': {'dateTime': ds.isoformat(), 'timeZone': 'Asia/Jakarta'},
                                    'end': {'dateTime': de.isoformat(), 'timeZone': 'Asia/Jakarta'},
                                    'reminders': get_reminders(r['Tahap'])
                                }
                                try: svc.events().insert(calendarId=CALENDAR_ID, body=evt).execute()
                                except: pass
                    s.update(label="Selesai!", state="complete", expanded=False)
                    st.success("Sukses! Data telah masuk ke Google Calendar.")
                    time.sleep(2); st.rerun()
                else: st.error("Gagal mengambil data. Pastikan link benar.")