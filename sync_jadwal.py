"""
sync_jadwal.py — Auto-sync jadwal tender SPSE ke Google Calendar
================================================================
Dijalankan via:
  - GitHub Actions (cron setiap 3 jam) — tanpa laptop
  - Manual: python sync_jadwal.py
  - Windows Task Scheduler (opsional, sebagai backup lokal)

Tidak membutuhkan Chrome/Selenium — murni urllib + pandas.
"""

import os
import io
import re
import json
import hashlib
import urllib.request
import datetime
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ============================================================
# KONFIGURASI
# ============================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, 'database_tender.csv')
LOG_PATH    = os.path.join(BASE_DIR, 'sync_log.txt')
TOKEN_PATH  = os.path.join(BASE_DIR, 'token.json')
CRED_PATH   = os.path.join(BASE_DIR, 'credentials.json')

CALENDAR_ID = 'primary'
SCOPES      = ['https://www.googleapis.com/auth/calendar']

BULAN_MAP = {
    'Januari':'01','Februari':'02','Maret':'03','April':'04',
    'Mei':'05','Juni':'06','Juli':'07','Agustus':'08',
    'September':'09','Oktober':'10','November':'11','Desember':'12'
}


# ============================================================
# LOGGING
# ============================================================
def log(msg: str):
    ts  = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


# ============================================================
# GOOGLE CALENDAR AUTH
# ============================================================
def get_service():
    """
    Mendukung 2 mode auth:
    1. GitHub Actions → baca dari env var GOOGLE_TOKEN_JSON
    2. Lokal           → baca dari token.json
    """
    creds = None

    # Mode GitHub Actions
    token_env = os.environ.get('GOOGLE_TOKEN_JSON')
    if token_env:
        creds = Credentials.from_authorized_user_info(json.loads(token_env), SCOPES)

    # Mode lokal
    elif os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log("🔄 Refreshing Google token...")
            creds.refresh(Request())
            # Simpan token baru ke file (mode lokal)
            if not token_env and os.path.exists(TOKEN_PATH):
                with open(TOKEN_PATH, 'w') as f:
                    f.write(creds.to_json())
        else:
            raise RuntimeError(
                "Token Google tidak valid dan tidak bisa di-refresh. "
                "Jalankan V19_Scheduler.py sekali secara lokal untuk login ulang."
            )

    return build('calendar', 'v3', credentials=creds)


# ============================================================
# SCRAPING (tanpa Selenium)
# ============================================================
def fetch_jadwal(url: str) -> pd.DataFrame | None:
    """
    Ambil tabel jadwal dari halaman SPSE menggunakan urllib.
    Tidak membutuhkan Chrome/Selenium.
    """
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
        log(f"  ⚠️  Gagal fetch {url}: {e}")
        return None

    # Ambil nama paket dari <title>
    nama_paket = None
    title_match = re.search(r'\[[\d]+\]\s*(.+?)(?:\s*</title>|\s*$)', html, re.IGNORECASE)
    if title_match:
        nama_paket = title_match.group(1).strip()

    # Parse tabel
    try:
        dfs = pd.read_html(io.StringIO(html), flavor='lxml')
    except Exception as e:
        log(f"  ⚠️  Gagal parse HTML {url}: {e}")
        return None

    target = None
    for df in dfs:
        if any('tahap' in str(c).lower() for c in df.columns):
            target = df
            break

    if target is None:
        log(f"  ⚠️  Tabel jadwal tidak ditemukan: {url}")
        return None

    # Normalisasi kolom
    cols     = target.columns
    c_tahap  = next((c for c in cols if 'tahap'     in str(c).lower()), None)
    c_mulai  = next((c for c in cols if 'mulai'     in str(c).lower()), None)
    c_sampai = next((c for c in cols if 'sampai'    in str(c).lower()), None)
    c_ubah   = next((c for c in cols if 'perubahan' in str(c).lower()), None)

    if not c_tahap or not c_mulai:
        return None

    df_clean = target[[c for c in [c_tahap, c_mulai, c_sampai, c_ubah] if c]].copy()
    df_clean.columns = ['Tahap', 'Mulai'] + (['Sampai'] if c_sampai else []) + (['Perubahan'] if c_ubah else [])
    if 'Sampai'    not in df_clean.columns: df_clean['Sampai']    = df_clean['Mulai']
    if 'Perubahan' not in df_clean.columns: df_clean['Perubahan'] = '0'
    df_clean['Perubahan'] = df_clean['Perubahan'].fillna('0')
    df_clean = df_clean.dropna(subset=['Tahap'])
    df_clean['Nama_Paket'] = nama_paket or f"Tender {url.split('/')[-2]}"

    return df_clean


def compute_hash(df: pd.DataFrame) -> str:
    """Hash konten jadwal — dipakai untuk deteksi perubahan."""
    content = df[['Tahap', 'Mulai', 'Sampai', 'Perubahan']].to_csv(index=False)
    return hashlib.md5(content.encode()).hexdigest()


# ============================================================
# GOOGLE CALENDAR — HELPERS
# ============================================================
def parse_date(date_str: str) -> datetime.datetime | None:
    try:
        if pd.isna(date_str) or str(date_str).strip() in ('-', ''):
            return None
        clean = re.sub(r'\s*\(.*?\)', '', str(date_str))
        parts = clean.split()
        if len(parts) >= 3:
            tgl = parts[0].zfill(2)
            bln = BULAN_MAP.get(parts[1], '01')
            thn = parts[2]
            jam = parts[3] if len(parts) > 3 else '00:00'
            return datetime.datetime.fromisoformat(f"{thn}-{bln}-{tgl}T{jam}:00")
    except Exception:
        pass
    return None


def get_reminders(tahap: str) -> dict:
    t = str(tahap).lower()
    if any(x in t for x in ['pembuktian', 'pembukaan', 'penunjukan']):
        return {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 24 * 60}]}
    elif 'penjelasan' in t:
        return {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 3 * 60}]}
    return {'useDefault': True}


def format_desc(url: str, perubahan, pokja: str, diff_info: str = '') -> str:
    """Format deskripsi event GCal dengan info perubahan."""
    try:
        raw = str(perubahan).strip()
        # Ekstrak angka dari "2 kali perubahan" / "3x" / "Tidak Ada"
        import re as _re
        match = _re.search(r'(\d+)', raw)
        jml = int(match.group(1)) if match else 0
    except:
        jml = 0

    if jml > 0:
        status = f"⚠️ PERINGATAN: {jml}x Perubahan"
    else:
        status = "✅ Aman"

    diff_block = ""
    if diff_info:
        diff_block = f"\n\n📋 PERUBAHAN TERDETEKSI:\n{diff_info}"

    pokja_str = f"\n\n👥 POKJA:\n{pokja}" if pokja else ""
    return f"🔗 Link: {url}\n\n{status}{diff_block}{pokja_str}"


def fetch_old_events(service, url: str) -> dict:
    """Ambil event lama dari GCal sebelum dihapus, return {summary: start_time}."""
    old_events = {}
    page_token = None
    while True:
        result = service.events().list(
            calendarId=CALENDAR_ID,
            q=url,
            singleEvents=True,
            pageToken=page_token
        ).execute()
        for ev in result.get('items', []):
            if url in ev.get('description', ''):
                summary = ev.get('summary', '')
                start = ev.get('start', {}).get('dateTime', '')
                if summary and start:
                    old_events[summary] = start
        page_token = result.get('nextPageToken')
        if not page_token:
            break
    return old_events


def fetch_jadwal_history(url: str) -> str:
    """
    Ambil history perubahan dari SPSE.
    Format return: '1x : 13 April 2026 10:00 - 13 April 2026 10:59\n2x : 14 April 2026 09:00 - 14 April 2026 11:00'
    """
    try:
        tender_id = url.split('/')[-2]
        base_url = '/'.join(url.split('/')[:-2])
        history_url = f"{base_url}/jadwal/{tender_id}/history"

        referer = url.replace('/jadwal', '/pengumuman')
        req = urllib.request.Request(history_url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Referer':    referer,
        })
        with urllib.request.urlopen(req, timeout=20) as resp:
            html = resp.read().decode('utf-8', errors='replace')

        dfs = pd.read_html(io.StringIO(html), flavor='lxml')
        if not dfs:
            return ""

        # Cari tabel yang punya kolom 'Tanggal Edit'
        for df in dfs:
            cols = [str(c).lower() for c in df.columns]
            if any('tanggal' in c or 'edit' in c for c in cols):
                if len(df) == 0:
                    return ""

                # Format per baris: "1x : 13 April 2026 10:00 - 13 April 2026 10:59"
                lines = []
                for _, row in df.iterrows():
                    vals = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip()]
                    if len(vals) >= 4:
                        # vals[0]=No, vals[1]=Tanggal Edit, vals[2]=Mulai, vals[3]=Sampai
                        no = vals[0]
                        mulai = vals[2]
                        sampai = vals[3]
                        lines.append(f"{no}x : {mulai} - {sampai}")

                return '\n'.join(lines) if lines else ""

    except Exception as e:
        log(f"  ⚠️ Gagal fetch history: {e}")
    return ""


def build_diff_info(old_events: dict, df_new: pd.DataFrame) -> str:
    """Bandingkan event lama vs baru, return string info perubahan."""
    if not old_events:
        return ""

    changes = []
    change_num = 0
    for _, row in df_new.iterrows():
        summary = f"{row['Tahap']} - {row['Nama_Paket']}"
        if summary in old_events:
            old_start = old_events[summary]
            # Parse old start time
            try:
                # old_start format: 2026-04-06T10:00:00+08:00
                old_dt = datetime.datetime.fromisoformat(old_start.replace('+08:00', ''))
                old_formatted = old_dt.strftime('%d %B %Y %H:%M')
                # Ganti nama bulan ke Indonesia
                for eng, indo in [('January','Januari'),('February','Februari'),('March','Maret'),
                                   ('April','April'),('May','Mei'),('June','Juni'),
                                   ('July','Juli'),('August','Agustus'),('September','September'),
                                   ('October','Oktober'),('November','November'),('December','Desember')]:
                    old_formatted = old_formatted.replace(eng, indo)
            except:
                old_formatted = old_start

            new_start = row['Mulai']
            new_sampai = row.get('Sampai', '')

            if str(old_formatted) != str(new_start):
                change_num += 1
                changes.append(f"  {change_num}. {row['Tahap']}: {old_formatted} → {new_start}")
                if new_sampai and str(new_sampai) != str(new_start):
                    changes[-1] += f" s/d {new_sampai}"

    if changes:
        return '\n'.join(changes)
    return ""


def delete_events_by_url(service, url: str):
    """Hapus semua event GCal yang mengandung URL ini di description."""
    page_token = None
    while True:
        result = service.events().list(
            calendarId=CALENDAR_ID,
            q=url,
            singleEvents=True,
            pageToken=page_token
        ).execute()
        for ev in result.get('items', []):
            if url in ev.get('description', ''):
                try:
                    service.events().delete(calendarId=CALENDAR_ID, eventId=ev['id']).execute()
                except Exception:
                    pass
        page_token = result.get('nextPageToken')
        if not page_token:
            break


def insert_events(service, df: pd.DataFrame, url: str, members: str, diff_info: str = ''):
    """Insert semua tahap sebagai event GCal."""
    inserted = 0
    for _, row in df.iterrows():
        ds = parse_date(row['Mulai'])
        de = parse_date(row['Sampai'])
        if not ds:
            continue
        if not de:
            de = ds + datetime.timedelta(hours=1)
        evt = {
            'summary':     f"{row['Tahap']} - {row['Nama_Paket']}",
            'description': format_desc(url, row['Perubahan'], members, diff_info=diff_info),
            'start':       {'dateTime': ds.isoformat(), 'timeZone': 'Asia/Jakarta'},
            'end':         {'dateTime': de.isoformat(), 'timeZone': 'Asia/Jakarta'},
            'reminders':   get_reminders(row['Tahap']),
        }
        try:
            service.events().insert(calendarId=CALENDAR_ID, body=evt).execute()
            inserted += 1
        except Exception as e:
            log(f"    ⚠️  Gagal insert event '{row['Tahap']}': {e}")
    return inserted


# ============================================================
# DATABASE
# ============================================================
def load_db() -> pd.DataFrame:
    cols = ['url', 'members', 'nama_paket', 'last_sync', 'content_hash']
    if os.path.exists(DB_PATH):
        try:
            df = pd.read_csv(DB_PATH)
            for c in cols:
                if c not in df.columns:
                    df[c] = ''
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=cols)


def save_db(df: pd.DataFrame):
    df.to_csv(DB_PATH, index=False)


def now_str() -> str:
    now = datetime.datetime.now()
    hari  = ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'][now.weekday()]
    bulan = ['','Januari','Februari','Maret','April','Mei','Juni',
             'Juli','Agustus','September','Oktober','November','Desember'][now.month]
    return f"{hari}, {now.day} {bulan} {now.year}, {now.strftime('%H:%M')}"


# ============================================================
# MAIN SYNC
# ============================================================
def sync_all():
    log("=" * 60)
    log("🚀 Mulai sync jadwal tender...")

    db = load_db()
    if db.empty:
        log("📭 Database kosong — tidak ada URL untuk disync.")
        return {'updated': 0, 'unchanged': 0, 'failed': 0}

    service     = get_service()
    updated     = 0
    unchanged   = 0
    failed      = 0

    for idx, row in db.iterrows():
        url     = str(row['url']).strip()
        members = str(row.get('members', ''))
        old_hash = str(row.get('content_hash', ''))

        log(f"\n🔍 [{idx+1}/{len(db)}] {row.get('nama_paket', url)}")

        df_jadwal = fetch_jadwal(url)
        if df_jadwal is None:
            log("  ❌ Gagal fetch — skip.")
            failed += 1
            continue

        new_hash = compute_hash(df_jadwal)
        nama_paket = df_jadwal['Nama_Paket'].iloc[0]

        if new_hash == old_hash:
            log("  ✅ Tidak ada perubahan.")
            unchanged += 1
            continue

        # Ada perubahan (atau pertama kali) — update calendar
        if old_hash:
            log("  🔄 Perubahan terdeteksi! Update Google Calendar...")
        else:
            log("  ➕ Entry baru — insert ke Google Calendar...")

        # Ambil info perubahan dari history SPSE
        diff_info = ""
        if old_hash:
            diff_info = fetch_jadwal_history(url)
            # Fallback: bandingkan event lama vs baru jika history kosong
            if not diff_info:
                old_events = fetch_old_events(service, url)
                if old_events:
                    diff_info = build_diff_info(old_events, df_jadwal)

        delete_events_by_url(service, url)
        n = insert_events(service, df_jadwal, url, members, diff_info=diff_info)
        log(f"  📅 {n} event berhasil dimasukkan.")

        # Update database
        db.at[idx, 'content_hash'] = new_hash
        db.at[idx, 'last_sync']    = now_str()
        db.at[idx, 'nama_paket']   = nama_paket
        updated += 1

    save_db(db)
    log(f"\n📊 Selesai — Updated: {updated} | Unchanged: {unchanged} | Failed: {failed}")
    log("=" * 60)
    return {'updated': updated, 'unchanged': unchanged, 'failed': failed}


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == '__main__':
    result = sync_all()
    # Exit code 1 kalau semua URL gagal (agar GitHub Actions tandai failed)
    if result['failed'] > 0 and result['updated'] == 0 and result['unchanged'] == 0:
        exit(1)
