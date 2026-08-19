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
import requests
import datetime
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from calendar_sync_targets import (
    TargetRegistryError,
    folder_identity_matches,
    load_targets,
    upsert_target,
)
from spse_public_http import get_public as _get_public_spse

# ============================================================
# KONFIGURASI
# ============================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, 'database_tender.csv')
LOG_PATH    = os.path.join(BASE_DIR, 'sync_log.txt')
_CANONICAL_SECRET_ROOT = os.path.join(os.path.dirname(BASE_DIR), 'Secrets')
_CONFIGURED_SECRET_ROOT = os.path.normpath(os.environ.get('POKJA_SECRET_ROOT', '')) if os.environ.get('POKJA_SECRET_ROOT') else ''
SECRET_ROOT = (
    _CANONICAL_SECRET_ROOT
    if os.path.exists(os.path.join(_CANONICAL_SECRET_ROOT, 'secret_supabase.env'))
    else (_CONFIGURED_SECRET_ROOT or _CANONICAL_SECRET_ROOT)
)
RUNTIME_ROOT = os.path.normpath(os.environ.get(
    'POKJA_RUNTIME_ROOT',
    os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~/AppData/Local')), 'POKJA2026', 'Asisten_Pokja'),
))
TOKEN_PATH  = os.path.join(RUNTIME_ROOT, 'state', 'token.json')
CRED_PATH   = os.path.join(SECRET_ROOT, 'credentials.json')

CALENDAR_ID = 'primary'
SCOPES      = ['https://www.googleapis.com/auth/calendar']
# LPSE Tapin memakai WITA (UTC+8), bukan WIB/Asia-Jakarta.
CALENDAR_TIMEZONE = 'Asia/Makassar'
SPSE_BASE_URL = os.environ.get('POKJA_SPSE_BASE_URL', 'https://spse.inaproc.id/tapinkab').rstrip('/')

_SPSE_SESSION = None

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
        else:
            raise RuntimeError(
                "Token Google tidak valid dan tidak bisa di-refresh. "
                "Jalankan V19_Scheduler.py sekali secara lokal untuk login ulang."
            )

    # Simpan token terbaru — lokal ke token.json, CI ke refreshed_token.json
    if token_env:
        with open('refreshed_token.json', 'w') as f:
            f.write(creds.to_json())
    elif os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'w') as f:
            f.write(creds.to_json())

    return build('calendar', 'v3', credentials=creds)


# ============================================================
# SCRAPING (tanpa Selenium)
# ============================================================
def _get_spse_session():
    """Pakai cloudscraper bila tersedia; fallback requests untuk environment minimal."""
    global _SPSE_SESSION
    if _SPSE_SESSION is not None:
        return _SPSE_SESSION
    try:
        import cloudscraper
        _SPSE_SESSION = cloudscraper.create_scraper(
            browser={'browser': 'chrome', 'platform': 'windows', 'mobile': False}
        )
    except ImportError:
        _SPSE_SESSION = requests
    return _SPSE_SESSION


def _get_spse(url: str, headers: dict, timeout: int = 30):
    session = _get_spse_session()
    return _get_public_spse(
        session,
        url,
        headers=headers,
        timeout=timeout,
        fallback=requests,
        log_fn=log,
    )


def fetch_jadwal(url: str) -> pd.DataFrame | None:
    """
    Ambil tabel jadwal dari halaman publik SPSE (/lelang/{id}/jadwal).
    Menggunakan urllib.request murni karena data jadwal adalah halaman publik biasa.
    """
    referer = url.replace('/jadwal', '/pengumuman')
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer': referer,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    }
    
    try:
        resp = _get_spse(url, headers=headers, timeout=30)
        resp.raise_for_status()
        html = resp.text
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
    cols = ['Tahap', 'Mulai', 'Sampai', 'Perubahan', 'Nama_Paket']
    content = df[cols].fillna('').astype(str).to_csv(index=False)
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


def _tender_code_from_url(url: str) -> str:
    match = re.search(r'/lelang/(\d+)', str(url or ''))
    return match.group(1) if match else ''


def _list_tender_events(service, url: str) -> list[dict]:
    """Ambil event paket baru (extended property) dan event legacy (URL di deskripsi)."""
    code = _tender_code_from_url(url)
    found = {}
    queries = [{'q': url}, {'privateExtendedProperty': f'source_tender={code}'}] if code else [{'q': url}]
    for query in queries:
        page_token = None
        while True:
            result = service.events().list(
                calendarId=CALENDAR_ID,
                singleEvents=True,
                pageToken=page_token,
                **query,
            ).execute()
            for ev in result.get('items', []):
                private = ev.get('extendedProperties', {}).get('private', {})
                if url in (ev.get('description', '') or '') or private.get('source_tender') == code:
                    found[ev.get('id')] = ev
            page_token = result.get('nextPageToken')
            if not page_token:
                break
    return [ev for ev in found.values() if ev.get('id')]


def fetch_old_events(service, url: str) -> dict:
    """Ambil event lama dari GCal sebelum dihapus, return {summary: start_time}."""
    old_events = {}
    for ev in _list_tender_events(service, url):
        summary = ev.get('summary', '')
        start = ev.get('start', {}).get('dateTime', '')
        if summary and start:
            old_events[summary] = start
    return old_events


def fetch_jadwal_history(url: str) -> dict:
    """
    Ambil history perubahan dari SPSE per tahap menggunakan urllib murni.
    Returns dict: {stage_name_lower: "1x : ...\n2x : ..."}
    """
    referer = url.replace('/jadwal', '/pengumuman')
    
    def _fetch_url(target_url, ref):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Referer': ref,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        }
        resp = _get_spse(target_url, headers=headers, timeout=20)
        resp.raise_for_status()
        return resp.text

    try:
        # Step 1: Fetch halaman jadwal, ekstrak tabel + link history
        html_text = _fetch_url(url, referer)

        rows = re.findall(r'<tr[^>]*>.*?</tr>', html_text, re.DOTALL)
        if not rows:
            return {}

        base_url = '/'.join(url.split('/')[:3])
        stage_history = {}

        for row in rows:
            cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row, re.DOTALL)
            if len(cells) < 5:
                continue
            tahap_name = re.sub(r'<[^>]+>', '', cells[1]).strip()
            if not tahap_name or tahap_name.lower() in ('tahap', 'no'):
                continue

            # Extract history link dari cell Perubahan
            history_links = re.findall(r'href="([^"]*history[^"]*)"', cells[4])
            if not history_links:
                continue

            # Step 2: Fetch history page untuk tahap ini
            hlink = history_links[0]
            full_url = hlink if hlink.startswith('http') else f"{base_url}{hlink}"
            
            try:
                html_history = _fetch_url(full_url, url)
            except Exception:
                continue

            tables = pd.read_html(io.StringIO(html_history), flavor='lxml')
            if not tables:
                continue

            for df in tables:
                cols = [str(c).lower() for c in df.columns]
                if any('tanggal' in c or 'edit' in c for c in cols):
                    if len(df) == 0:
                        continue
                    lines = []
                    for _, row_data in df.iterrows():
                        vals = [str(v).strip() for v in row_data.values if pd.notna(v) and str(v).strip()]
                        if len(vals) >= 4:
                            no, tgl_edit, mulai, sampai = vals[0], vals[1], vals[2], vals[3]
                            lines.append(f"{no}x : {mulai} - {sampai}, diedit pada tanggal : {tgl_edit}")
                    if lines:
                        stage_history[tahap_name.lower()] = '\n'.join(lines)
                    break

        return stage_history
    except Exception as e:
        log(f"  ⚠️ Gagal fetch history: {e}")
    return {}


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
    for ev in _list_tender_events(service, url):
        try:
            service.events().delete(calendarId=CALENDAR_ID, eventId=ev['id']).execute()
        except Exception:
            pass


def insert_events(service, df: pd.DataFrame, url: str, members: str, stage_history: dict = None):
    """Insert semua tahap sebagai event GCal."""
    stage_history = stage_history or {}
    inserted = 0
    for _, row in df.iterrows():
        ds = parse_date(row['Mulai'])
        de = parse_date(row['Sampai'])
        if not ds:
            continue
        if not de:
            de = ds + datetime.timedelta(hours=1)
        # Lookup stage-specific history
        stage_key = str(row['Tahap']).strip().lower()
        stage_diff = stage_history.get(stage_key, '')
        evt = {
            'summary':     f"{row['Tahap']} - {row['Nama_Paket']}",
            'description': format_desc(url, row['Perubahan'], members, diff_info=stage_diff),
            'start':       {'dateTime': ds.isoformat(), 'timeZone': CALENDAR_TIMEZONE},
            'end':         {'dateTime': de.isoformat(), 'timeZone': CALENDAR_TIMEZONE},
            'reminders':   get_reminders(row['Tahap']),
        }
        try:
            service.events().insert(calendarId=CALENDAR_ID, body=evt).execute()
            inserted += 1
        except Exception as e:
            log(f"    ⚠️  Gagal insert event '{row['Tahap']}': {e}")
    return inserted


def _tender_event_body(row, index: int, url: str, members: str, stage_history: dict) -> dict | None:
    ds = parse_date(row['Mulai'])
    de = parse_date(row['Sampai'])
    if not ds:
        return None
    if not de:
        de = ds + datetime.timedelta(hours=1)
    stage_key = str(row['Tahap']).strip().lower()
    return {
        'summary': f"{row['Tahap']} - {row['Nama_Paket']}",
        'description': format_desc(
            url, row['Perubahan'], members,
            diff_info=stage_history.get(stage_key, ''),
        ),
        'start': {'dateTime': ds.isoformat(), 'timeZone': CALENDAR_TIMEZONE},
        'end': {'dateTime': de.isoformat(), 'timeZone': CALENDAR_TIMEZONE},
        'reminders': get_reminders(row['Tahap']),
        'extendedProperties': {
            'private': {
                'source_tender': _tender_code_from_url(url),
                'source_stage_index': str(index),
            }
        },
    }


def reconcile_tender_events(
    service, df: pd.DataFrame, url: str, members: str, stage_history: dict | None = None,
) -> dict:
    """Reconcile event per tahap; tidak delete-before-insert sehingga partial failure aman."""
    stage_history = stage_history or {}
    existing = _list_tender_events(service, url)
    by_index = {}
    by_summary = {}
    for ev in existing:
        private = ev.get('extendedProperties', {}).get('private', {})
        if private.get('source_stage_index') is not None:
            by_index.setdefault(str(private['source_stage_index']), []).append(ev)
        if ev.get('summary'):
            by_summary.setdefault(ev['summary'], []).append(ev)

    valid_rows = []
    for index, (_, row) in enumerate(df.iterrows()):
        body = _tender_event_body(row, index, url, members, stage_history)
        if body is not None:
            valid_rows.append((index, row, body))
    if not valid_rows:
        return {'ok': False, 'inserted': 0, 'updated': 0, 'deleted': 0, 'error': 'Tidak ada tahap dengan tanggal valid.'}

    used_ids = set()
    inserted = updated = 0
    errors = []
    for index, row, body in valid_rows:
        candidates = by_index.get(str(index), []) or by_summary.get(body['summary'], [])
        event = next((ev for ev in candidates if ev.get('id') not in used_ids), None)
        try:
            if event:
                service.events().update(
                    calendarId=CALENDAR_ID, eventId=event['id'], body=body,
                ).execute()
                used_ids.add(event['id'])
                updated += 1
            else:
                response = service.events().insert(calendarId=CALENDAR_ID, body=body).execute()
                if isinstance(response, dict) and response.get('id'):
                    used_ids.add(response['id'])
                inserted += 1
        except Exception as exc:
            errors.append(f"Tahap {index + 1} ({row['Tahap']}) gagal: {exc}")

    # Jangan menghapus event lama bila ada insert/update yang gagal.
    deleted = 0
    if not errors:
        for ev in existing:
            if ev.get('id') in used_ids:
                continue
            try:
                service.events().delete(calendarId=CALENDAR_ID, eventId=ev['id']).execute()
                deleted += 1
            except Exception as exc:
                errors.append(f"Hapus event stale gagal: {exc}")

    return {
        'ok': not errors and inserted + updated == len(valid_rows),
        'inserted': inserted,
        'updated': updated,
        'deleted': deleted,
        'error': ' | '.join(errors)[:1000],
    }


def _tender_events_complete(service, url: str, df: pd.DataFrame) -> bool:
    expected = {
        f"{row['Tahap']} - {row['Nama_Paket']}"
        for _, row in df.iterrows()
        if parse_date(row['Mulai']) is not None
    }
    if not expected:
        return False
    actual = {ev.get('summary') for ev in _list_tender_events(service, url)}
    return expected.issubset(actual)


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


def _read_env_file(path: str) -> dict[str, str]:
    values = {}
    try:
        with open(path, encoding='utf-8') as handle:
            for line in handle:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, value = line.split('=', 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    except OSError:
        pass
    return values


def _load_supabase_tender_rows(codes: list[str] | None = None) -> list[dict]:
    """Ambil metadata Tender untuk kode yang sudah diizinkan registry."""
    roots = [
        os.environ.get('POKJA_SECRET_ROOT', ''),
        SECRET_ROOT,
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'POKJA2026', 'Secrets'),
        os.path.join(os.path.dirname(BASE_DIR), 'Secrets'),
    ]
    env = {}
    for root in dict.fromkeys(os.path.normpath(r) for r in roots if r):
        env.update(_read_env_file(os.path.join(root, 'secret_supabase.env')))
    supabase_url = (env.get('SUPABASE_URL') or os.environ.get('SUPABASE_URL', '')).strip().rstrip('/')
    supabase_key = (env.get('SUPABASE_KEY') or os.environ.get('SUPABASE_KEY', '')).strip()
    if not supabase_url or not supabase_key:
        log('  ⚠️ Discovery Supabase dilewati: secret_supabase.env tidak ditemukan/lengkap.')
        return []
    try:
        params = {
            'select': 'kode_tender,nama_tender,kode_pokja,folder_dibuat',
            'limit': str(max(len(codes), 1)) if codes else '1000',
        }
        if codes:
            params['kode_tender'] = f"in.({','.join(codes)})"
        response = requests.get(
            f'{supabase_url}/rest/v1/draft_paket',
            params=params,
            headers={'apikey': supabase_key, 'Authorization': f'Bearer {supabase_key}'},
            timeout=30,
        )
        response.raise_for_status()
        rows = response.json()
    except Exception as exc:
        log(f'  ⚠️ Discovery Supabase gagal: {exc}')
        return []

    result = []
    for row in rows if isinstance(rows, list) else []:
        code = str(row.get('kode_tender') or '').strip()
        if not code:
            continue
        result.append({
            'url': f'{SPSE_BASE_URL}/lelang/{code}/jadwal',
            'members': str(row.get('kode_pokja') or '').strip(),
            'nama_paket': str(row.get('nama_tender') or code).strip(),
            'folder_name': str(row.get('folder_dibuat') or '').strip(),
        })
    return result


def _pokja_drive_roots() -> list[str]:
    configured = os.environ.get('POKJA_DRIVE_ROOT', '').strip().strip('"')
    return list(dict.fromkeys(os.path.normpath(root) for root in (
        configured,
        r'D:\Dokumen\@ POKJA 2026',
        r'G:\Other computers\My Laptop\@ POKJA 2026',
    ) if root))


def _tender_folder_identity_valid(folder_name: str, kode_tender: str) -> bool:
    return folder_identity_matches(
        folder_name,
        kode_tender,
        [os.path.join(root, '@ Tender 2026') for root in _pokja_drive_roots()],
        '@ Master Data',
        ('C4',),
    )


def _auto_enroll_folder_tenders(existing_targets: list[dict]) -> None:
    """Enroll friend-created SPSE packages after user creates a local folder."""
    known = {
        str(row.get('kode_paket') or '').strip(): row
        for row in existing_targets
    }
    for row in _load_supabase_tender_rows():
        code = str(row.get('url', '')).split('/lelang/')[-1].split('/')[0]
        folder_name = str(row.get('folder_name') or '').strip()
        if not code.isdigit() or not folder_name or code in known:
            continue
        if not _tender_folder_identity_valid(folder_name, code):
            continue
        upsert_target(
            'tender', code,
            name=row.get('nama_paket', code),
            folder_name=folder_name,
            source='folder-auto',
            note='Auto-enrolled karena folder paket ada di root POKJA lokal.',
        )


def _owned_tender_rows(db: pd.DataFrame, targets: list[dict]) -> pd.DataFrame:
    """Bangun kandidat hanya dari allowlist aktif, bukan dari seluruh sumber."""
    csv_by_code = {}
    for _, row in db.iterrows():
        code = _tender_code_from_url(row.get('url', ''))
        if code:
            csv_by_code[code] = row.to_dict()

    codes = [str(target.get('kode_paket') or '').strip() for target in targets]
    metadata = {
        str(row.get('kode_tender') or '').strip(): row
        for row in _load_supabase_tender_rows(codes)
    }
    rows = []
    for target in targets:
        code = str(target.get('kode_paket') or '').strip()
        if not code:
            continue
        if (
            str(target.get('source') or '').strip() == 'folder-auto'
            and not _tender_folder_identity_valid(
                target.get('folder_name', ''), code
            )
        ):
            log(f'  ⏭️ Tender {code} dilewati: folder POKJA tidak terbaca.')
            continue
        old = dict(csv_by_code.get(code, {}))
        source = metadata.get(code, {})
        rows.append({
            **old,
            'url': f'{SPSE_BASE_URL}/lelang/{code}/jadwal',
            'members': str(target.get('kode_pokja') or source.get('kode_pokja') or old.get('members') or '').strip(),
            'nama_paket': str(target.get('nama_paket') or source.get('nama_tender') or old.get('nama_paket') or code).strip(),
            'last_sync': old.get('last_sync', ''),
            'content_hash': old.get('content_hash', ''),
        })
    return pd.DataFrame(rows)


def _merge_database_updates(original: pd.DataFrame, synced: pd.DataFrame) -> pd.DataFrame:
    """Pertahankan CSV legacy; hanya update/append baris target aktif."""
    if original.empty:
        return synced
    result = original.copy()
    result = result.set_index('url')
    for _, row in synced.iterrows():
        url = str(row.get('url', '')).strip()
        if not url:
            continue
        if url not in result.index:
            columns = [column for column in row.index if column != 'url']
            result.loc[url, columns] = row[columns].to_dict()
        else:
            for column in ('members', 'nama_paket', 'last_sync', 'content_hash'):
                if column in row.index:
                    result.loc[url, column] = row[column]
    return result.reset_index()


def _legacy_merge_discovered_tenders(db: pd.DataFrame) -> pd.DataFrame:
    """Legacy helper; scheduler tidak lagi memakai discovery implisit."""
    discovered = _load_supabase_tender_rows()
    if not discovered:
        return db
    existing = {str(url).strip() for url in db.get('url', pd.Series(dtype=str)).tolist()}
    added = 0
    additions = []
    for row in discovered:
        if row['url'] in existing:
            continue
        additions.append(row)
        existing.add(row['url'])
        added += 1
    if additions:
        db = pd.concat([db, pd.DataFrame(additions)], ignore_index=True)
    log(f'  🔎 Discovery Supabase: {len(discovered)} paket berfolder, {added} baru.')
    return db


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

    try:
        all_targets = load_targets('tender', enabled_only=False)
        _auto_enroll_folder_tenders(all_targets)
        targets = load_targets('tender')
    except TargetRegistryError as exc:
        log(f'  ❌ Allowlist Tender tidak tersedia — fail-closed: {exc}')
        return {'updated': 0, 'unchanged': 0, 'empty': 0, 'failed': 1}

    db_source = load_db()
    db = _owned_tender_rows(db_source, targets)
    if db.empty:
        log("📭 Tidak ada target Tender aktif — tidak ada URL untuk discrape.")
        return {'updated': 0, 'unchanged': 0, 'empty': 0, 'failed': 0}

    service     = get_service()
    updated     = 0
    unchanged   = 0
    empty       = 0
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
        if df_jadwal.empty:
            log("  ⚠️ Jadwal kosong di SPSE — skip sementara.")
            empty += 1
            continue
        if 'Nama_Paket' not in df_jadwal.columns:
            log("  ❌ Tabel jadwal tidak valid — skip.")
            failed += 1
            continue

        new_hash = compute_hash(df_jadwal)
        nama_paket = df_jadwal['Nama_Paket'].iloc[0]

        if new_hash == old_hash:
            try:
                if _tender_events_complete(service, url, df_jadwal):
                    log("  ✅ Tidak ada perubahan; event GCal terverifikasi.")
                    unchanged += 1
                    continue
                log("  ⚠️ Hash sama tetapi event GCal hilang/tidak lengkap — reconcile.")
            except Exception as exc:
                log(f"  ❌ Gagal verifikasi event GCal: {exc}")
                failed += 1
                continue

        # Ada perubahan (atau pertama kali) — update calendar
        if old_hash:
            log("  🔄 Perubahan terdeteksi! Update Google Calendar...")
        else:
            log("  ➕ Entry baru — insert ke Google Calendar...")

        # Ambil info perubahan per tahap dari history SPSE
        stage_history = {}
        if old_hash:
            stage_history = fetch_jadwal_history(url)

        cal_result = reconcile_tender_events(
            service, df_jadwal, url, members, stage_history=stage_history,
        )
        if not cal_result['ok']:
            log(f"  ❌ Reconcile GCal gagal: {cal_result['error']}")
            failed += 1
            continue
        log(
            f"  📅 GCal: {cal_result['inserted']} baru, "
            f"{cal_result['updated']} diperbarui, {cal_result['deleted']} stale dihapus."
        )

        # Update database
        db.at[idx, 'content_hash'] = new_hash
        db.at[idx, 'last_sync']    = now_str()
        db.at[idx, 'nama_paket']   = nama_paket
        updated += 1

    save_db(_merge_database_updates(db_source, db))
    log(
        f"\n📊 Selesai — Updated: {updated} | Unchanged: {unchanged} | "
        f"Empty: {empty} | Failed: {failed}"
    )
    log("=" * 60)
    return {'updated': updated, 'unchanged': unchanged, 'empty': empty, 'failed': failed}


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == '__main__':
    result = sync_all()
    # Exit code 1 kalau semua URL gagal (agar GitHub Actions tandai failed)
    if result['failed'] > 0 and result['updated'] == 0 and result['unchanged'] == 0:
        exit(1)
