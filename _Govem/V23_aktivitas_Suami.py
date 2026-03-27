import os
import time
import subprocess
import datetime
import configparser
import json
import glob
import math
import sys
import re
import builtins

# Fix pythonw cp1252 encoding crash — override print() to silently handle emoji
if not getattr(builtins.print, '_safe_wrapped', False):
    _original_print = builtins.print
    def _safe_print(*args, **kwargs):
        try:
            _original_print(*args, **kwargs)
        except (UnicodeEncodeError, ValueError, OSError):
            pass
    _safe_print._safe_wrapped = True
    builtins.print = _safe_print

# Konfigurasi LDPlayer (Sama dengan V22)
LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")

PACKAGE_NAME = "go.id.tapinkab.govem"
CONFIG_FILE = "govem_aktivitas_config.ini" # Config khusus aktivitas

# Google Calendar OAuth (reuse dari V19_Scheduler)
SCHEDULER_DIR = r"D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"
CRED_FILE_PATH = os.path.join(SCHEDULER_DIR, "credentials.json")
TOKEN_FILE_PATH = os.path.join(SCHEDULER_DIR, "token.json")
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'primary'

# Global Config Storage
BUTTON_MAP = {}

# --- DRY RUN FLAG ---
DRY_RUN = "--dry-run" in sys.argv

# =============================================
# KOORDINAT UI GOVEM (terdeteksi via ADB screenshot)
# =============================================
# Dropdown SKP: buka dengan tap (500, 318)
# Koordinat dari uiautomator dump — resolusi emulator 1600x900
SKP_DROPDOWN_XY = (500, 318)
SKP_COORDS = {
    1: (800, 518),   # Melaksanakan Proses pemilihan penyedia barang/jasa [0,469][1600,568]
    2: (800, 617),   # Melaksanakan Kegiatan pengembangan kompetensi [0,568][1600,667]
    3: (800, 716),   # Melaksanakan Kegiatan yang menunjang pengelolaan [0,667][1600,766]
    4: (800, 815),   # Melaksanakan tugas lain dari pimpinan [0,766][1600,865]
    # 5, 6 perlu scroll dulu
}

# Dropdown Jenis: buka dengan tap (1130, 318)
JENIS_DROPDOWN_XY = (1130, 318)
JENIS_COORDS = {
    1: (800, 508),   # Aktifitas [0,475][1600,542]
    2: (800, 592),   # Apel/Shif/Piket/Lainnya [0,559][1600,626]
}

# Navigasi & Form (decoded dari macro .record)
# Step 1: Dashboard → Form (2 tap berurutan)
STEP1_TAP_1 = (1153, 769)  # Klik menu aktivitas
STEP1_TAP_2 = (500, 331)   # Klik buat aktivitas harian
# Step 2: Klik field input teks
STEP2_INPUT_XY = (246, 456)
# Step 5: Tombol simpan/posting
STEP5_SAVE_XY = (795, 897)

# =============================================
# TEMPLATE AKTIVITAS — format: {teks, skp, jenis}
# skp: nomor SKP dropdown (1-4, atau 6 untuk senam)
# jenis: 1=Aktifitas, 2=Apel/Shif/Piket
# =============================================
TEMPLATE_RUTINITAS = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Menelaah dan menganalisa dokumen peraturan presiden no 46 tahun 2025 tentang perubahan kedua atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah SE Kep LKPP No. 1 2025 tentang Penjelasan Atas Pelaksanaan Perpres No. 46 2025 tentang Perubahan Kedua Atas Perpres No. 16 2018 Pada Masa Transisi", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Presiden No. 12 tahun 2021 tentang perubahan pertama atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 12 Tahun 2021 tentang Pedoman Pelaksanaan Pengadaan Barang/Jasa Pemerintah Melalui Penyedia", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 3 Tahun 2021 tentang Pedoman Swakelola", "skp": 4, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

TEMPLATE_JUMAT = [
    {"teks": "melaksanakan senam pagi", "skp": 6, "jenis": 1},
    {"teks": "Menelaah dan menganalisa dokumen peraturan presiden no 46 tahun 2025 tentang perubahan kedua atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah SE Kep LKPP No. 1 2025 tentang Penjelasan Atas Pelaksanaan Perpres No. 46 2025 tentang Perubahan Kedua Atas Perpres No. 16 2018 Pada Masa Transisi", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Presiden No. 12 tahun 2021 tentang perubahan pertama atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 12 Tahun 2021 tentang Pedoman Pelaksanaan Pengadaan Barang/Jasa Pemerintah Melalui Penyedia", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 3 Tahun 2021 tentang Pedoman Swakelola", "skp": 4, "jenis": 1},
]

TEMPLATE_DOKPIL = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Melaksanakan pembuatan dokumen pemilihan paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pembuatan jadwal tayang paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan upload dokumen pemilihan paket tender/seleksi pada SPSE", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan persetujuan paket tender/seleksi yang tayang pada SPSE", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pengeditan paket (checklist syarat kualifikasi dan syarat adm/teknis) tender/seleksi yang akan ditayangkan di SPSE", "skp": 1, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

TEMPLATE_EVALUASI = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Melaksanakan evaluasi kualifikasi adm paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan evaluasi kualifikasi teknis paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan evaluasi administrasi paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan evaluasi teknis paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan evaluasi harga paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

TEMPLATE_PEMBUKTIAN = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Melaksanakan evaluasi pembuktian kualifikasi adm paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan evaluasi pembuktian kualifikasi teknis paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan negosiasi harga paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pembuatan berita acara dan daftar hadir pembuktian kualifikasi paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pembuatan berita acara negosiasi paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

TEMPLATE_REVIU_DPP = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Melaksanakan reviu DPP paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pembuatan berita acara reviu DPP paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan pembuatan daftar hadir reviu DPP paket tender/seleksi", "skp": 1, "jenis": 1},
    {"teks": "Melaksanakan posting dokumen BA Reviu DPP paket tender/seleki ke SPSE", "skp": 1, "jenis": 1},
    {"teks": "Menelaah SE Kep LKPP No. 1 2025 tentang Penjelasan Atas Pelaksanaan Perpres No. 46 2025 tentang Perubahan Kedua Atas Perpres No. 16 2018 Pada Masa Transisi", "skp": 4, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

# Mapping keyword tahapan SPSE → template
# Event summary format dari V19_Scheduler: "{Tahap} - {Nama_Paket}"
TAHAPAN_MAPPING = [
    # (keyword dalam summary event, nama template, template list)
    ("Pengumuman Prakualifikasi",           "DOKPIL",      TEMPLATE_DOKPIL),
    ("Pengumuman Pasca",                    "DOKPIL",      TEMPLATE_DOKPIL),
    ("Download Dokumen Pemilihan",          "DOKPIL",      TEMPLATE_DOKPIL),
    ("Evaluasi Dokumen Kualifikasi",        "EVALUASI",    TEMPLATE_EVALUASI),
    ("Evaluasi Penawaran",                  "EVALUASI",    TEMPLATE_EVALUASI),
    ("Pembukaan dan Evaluasi",              "EVALUASI",    TEMPLATE_EVALUASI),
    ("Evaluasi Administrasi",               "EVALUASI",    TEMPLATE_EVALUASI),
    ("Evaluasi Teknis",                     "EVALUASI",    TEMPLATE_EVALUASI),
    ("Pembuktian Kualifikasi",              "PEMBUKTIAN",  TEMPLATE_PEMBUKTIAN),
    ("Klarifikasi dan Negosiasi",           "PEMBUKTIAN",  TEMPLATE_PEMBUKTIAN),
    ("Penetapan Hasil Kualifikasi",         "REVIU_DPP",   TEMPLATE_REVIU_DPP),
    ("Pengumuman Hasil Prakualifikasi",     "REVIU_DPP",   TEMPLATE_REVIU_DPP),
    ("Penetapan Pemenang",                  "REVIU_DPP",   TEMPLATE_REVIU_DPP),
    ("Pengumuman Pemenang",                 "REVIU_DPP",   TEMPLATE_REVIU_DPP),
]

# Prioritas template (semakin tinggi = lebih diprioritaskan jika ada >1 event)
TEMPLATE_PRIORITY = {
    "DOKPIL": 1,
    "EVALUASI": 4,
    "PEMBUKTIAN": 3,
    "REVIU_DPP": 2,
}


# =============================================
# GOOGLE CALENDAR INTEGRATION
# =============================================
def get_calendar_service():
    """Connect ke Google Calendar API, reuse token dari V19_Scheduler."""
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        print("⚠️ Google API library tidak tersedia. Install: pip install google-api-python-client google-auth-oauthlib")
        return None

    creds = None
    if os.path.exists(TOKEN_FILE_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open(TOKEN_FILE_PATH, 'w') as token:
                    token.write(creds.to_json())
            except Exception as e:
                print(f"⚠️ Gagal refresh token: {e}")
                return None
        else:
            print("⚠️ Token tidak valid dan tidak bisa di-refresh. Jalankan V19_Scheduler dulu untuk authorize.")
            return None

    return build('calendar', 'v3', credentials=creds)


def get_today_events():
    """Ambil semua event Google Calendar hari ini."""
    service = get_calendar_service()
    if not service:
        return []

    today = datetime.date.today()
    time_min = datetime.datetime.combine(today, datetime.time.min).isoformat() + '+08:00'
    time_max = datetime.datetime.combine(today, datetime.time.max).isoformat() + '+08:00'

    try:
        events_result = service.events().list(
            calendarId=CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        print(f"📅 Google Calendar: {len(events)} event hari ini ({today.strftime('%d/%m/%Y')})")
        for e in events:
            summary = e.get('summary', '(tanpa judul)')
            print(f"   • {summary}")
        return events
    except Exception as e:
        print(f"⚠️ Gagal ambil calendar: {e}")
        return []


def map_events_to_template(events, is_jumat):
    """
    Mapping event calendar hari ini ke template aktivitas.
    Return: (template_name, activity_list)

    Logic:
    - Scan semua event, cocokkan summary dengan TAHAPAN_MAPPING
    - Jika ada >1 match, pilih yang prioritas tertinggi (evaluasi > pembuktian > reviu > dokpil)
    - Jika tidak ada match, fallback ke rutinitas/jumat
    """
    best_match = None
    best_priority = -1

    for event in events:
        summary = event.get('summary', '').lower()
        for keyword, template_name, template_list in TAHAPAN_MAPPING:
            if keyword.lower() in summary:
                priority = TEMPLATE_PRIORITY.get(template_name, 0)
                if priority > best_priority:
                    best_priority = priority
                    best_match = (template_name, template_list)
                break  # satu event cukup satu match

    if best_match:
        template_name, template_list = best_match
        print(f"🎯 Template terpilih: {template_name} (prioritas {best_priority})")
        if is_jumat:
            # Jumat: 6 aktivitas, ganti apel pagi→senam, hapus apel sore
            activities = [dict(a) for a in template_list]  # deep copy
            if activities and "apel pagi" in activities[0]["teks"].lower():
                activities[0] = {"teks": "melaksanakan senam pagi", "skp": 6, "jenis": 1}
            if activities and "apel sore" in activities[-1]["teks"].lower():
                activities.pop()
            return template_name, activities
        return template_name, [dict(a) for a in template_list]

    # Fallback
    if is_jumat:
        print("📋 Tidak ada event tender → Template: JUMAT (default)")
        return "JUMAT", [dict(a) for a in TEMPLATE_JUMAT]
    else:
        print("📋 Tidak ada event tender → Template: RUTINITAS (default)")
        return "RUTINITAS", [dict(a) for a in TEMPLATE_RUTINITAS]


# =============================================
# CORE FUNCTIONS (unchanged from v1.0)
# =============================================
def run_command(command):
    if DRY_RUN:
        return ""
    try:
        creationflags = 0
        if os.name == 'nt':
            creationflags = 0x08000000 # CREATE_NO_WINDOW

        if isinstance(command, list):
            result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
        else:
            result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
        return result.stdout.strip()
    except Exception as e:
        print(f"Error executing command: {e}")
        return ""

def load_config():
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
    return config

def save_config(section, key, value):
    config = load_config()
    if not config.has_section(section):
        config.add_section(section)
    config.set(section, key, str(value))
    with open(CONFIG_FILE, "w") as f:
        config.write(f)

# --- REUSED ENGINE FROM V22 (Smart Connect) ---
def connect_adb_smart(idx, launch_if_needed=False):
    if DRY_RUN:
        print(f"🔗 [DRY RUN] Skip ADB connect (emulator {idx})")
        return "dry-run-serial"

    print(f"🔗 Menghubungkan Emulator {idx} (Background Mode: {not launch_if_needed})...")

    if launch_if_needed:
        print("   📱 Launching emulator (foreground)...")
        run_command(f'"{LDCONSOLE}" launch --index {idx}')
        # Loop minimize sampai emulator running
        for attempt in range(30):
            time.sleep(1)
            run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
            s = run_command(f'"{LDCONSOLE}" list2')
            for line in s.splitlines():
                parts = line.split(",")
                if len(parts) > 4 and parts[0] == str(idx) and parts[4] == "1":
                    print(f"   ✅ Emulator running setelah {attempt+1}s, minimize sent!")
                    break
            else:
                continue
            break
        # Extra minimize untuk memastikan
        for _ in range(3):
            time.sleep(1)
            run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
    else:
        print("   🔇 Skip launch (background mode - assuming emulator already running)")

    possible_ports = [5555 + (idx*2), 5557, 5559, 5561, 5563, 5565]
    detected_serial = None

    print("   🔍 Scanning Ports...")
    for p in possible_ports:
        run_command(f'"{ADB}" connect 127.0.0.1:{p}')

    for i in range(10):
        devices = run_command(f'"{ADB}" devices')
        for p in possible_ports:
             if f"127.0.0.1:{p}" in devices and "device" in devices:
                 detected_serial = f"127.0.0.1:{p}"
                 break
        if detected_serial: break
        time.sleep(1)

    if not detected_serial:
        detected_serial = f"127.0.0.1:{5555 + (idx*2)}"

    print(f"   ✅ Terhubung ke: {detected_serial}")
    return detected_serial

def adb_click(serial, x, y):
    run_command([ADB, "-s", serial, "shell", "input", "tap", str(x), str(y)])

def adb_input_text(serial, text, idx=0):
    if DRY_RUN:
        return
    escaped_text = text.replace(" ", "%s")
    escaped_text = escaped_text.replace("'", "")
    escaped_text = escaped_text.replace('"', "")
    escaped_text = escaped_text.replace("&", "")
    escaped_text = escaped_text.replace("(", "")
    escaped_text = escaped_text.replace(")", "")
    escaped_text = escaped_text.replace(";", "")

    # ADB input text limit ~150 chars — kirim per chunk
    CHUNK_SIZE = 120
    if len(escaped_text) <= CHUNK_SIZE:
        run_command([ADB, "-s", serial, "shell", "input", "text", escaped_text])
    else:
        for j in range(0, len(escaped_text), CHUNK_SIZE):
            chunk = escaped_text[j:j+CHUNK_SIZE]
            run_command([ADB, "-s", serial, "shell", "input", "text", chunk])
            time.sleep(0.3)

# --- SMART ACTIVITY SELECTOR (v1.1) ---
def get_suami_activities(day_idx):
    """
    v1.1: Cek Google Calendar dulu, baru fallback ke rutinitas.
    day_idx: 0=Senin, 1=Selasa, ..., 4=Jumat
    """
    is_jumat = (day_idx == 4)

    # 1. Cek Google Calendar
    print("\n🔍 Mengecek Google Calendar...")
    events = get_today_events()

    # 2. Mapping event → template
    template_name, activities = map_events_to_template(events, is_jumat)

    # 3. Validasi jumlah
    expected = 6 if is_jumat else 7
    if len(activities) != expected:
        print(f"⚠️ Template {template_name} punya {len(activities)} aktivitas, expected {expected}")

    print(f"\n📝 Daftar aktivitas ({template_name}):")
    for i, act in enumerate(activities):
        teks = act["teks"]
        print(f"   {i+1}. [SKP{act['skp']} J{act['jenis']}] {teks[:70]}{'...' if len(teks) > 70 else ''}")

    return activities, is_jumat


# --- HELPER: DIRECT JSON REPLAY (BYPASS LDCONSOLE) ---
_CACHED_SERIAL = {}

def parse_json_and_replay(idx, record_filename_no_ext, serial=None):
    if DRY_RUN:
        print(f"   ▶️ [DRY RUN] Skip macro: {record_filename_no_ext[:40]}...")
        return

    global _CACHED_SERIAL

    records_dir = r"D:\LDPlayer\LDPlayer9\vms\operationRecords"
    filepath = os.path.join(records_dir, record_filename_no_ext + ".record")

    if not os.path.exists(filepath):
        print(f"❌ Record tidak ditemukan: {filepath}")
        return

    print(f"   ▶️ Macro: {record_filename_no_ext[:40]}...")

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        ops = data.get('operations', [])

        if serial:
            current_serial = serial
        elif idx in _CACHED_SERIAL:
            current_serial = _CACHED_SERIAL[idx]
        else:
            current_serial = connect_adb_smart(idx, launch_if_needed=False)
            _CACHED_SERIAL[idx] = current_serial

        w_res = 1600
        h_res = 900
        max_in_x = 19092
        max_in_y = 10728

        last_timing = 0

        for op in ops:
            timing = op.get('timing', 0)
            op_id = op.get('operationId')

            delay = (timing - last_timing) / 1000.0 * 0.5
            if delay > 0:
                time.sleep(delay)
            last_timing = timing

            if op_id == 'PutMultiTouch':
                points = op.get('points', [])
                if points and points[0].get('state') == 1:
                    raw_x = points[0].get('x')
                    raw_y = points[0].get('y')

                    real_x = int(raw_x / max_in_x * w_res)
                    real_y = int(raw_y / max_in_y * h_res)

                    run_command([ADB, "-s", current_serial, "shell", "input", "swipe", str(real_x), str(real_y), str(real_x), str(real_y), "100"])

            elif op_id == 'Wait':
                 pass

        print("   ✅ Selesai")

    except Exception as e:
        print(f"   ❌ Gagal: {e}")

def play_record_file(idx, record_filename_no_ext, serial=None):
    parse_json_and_replay(idx, record_filename_no_ext, serial)

# --- SCREENSHOT & NOTIFICATION ---
def take_screenshot(serial, idx):
    """Ambil screenshot emulator via ADB, simpan ke temp file."""
    if DRY_RUN:
        print("📸 [DRY RUN] Skip screenshot")
        return None
    try:
        import tempfile
        screenshot_path = os.path.join(tempfile.gettempdir(), f"govem_aktivitas_{idx}.png")
        result = subprocess.run(
            [ADB, "-s", serial, "exec-out", "screencap", "-p"],
            capture_output=True, timeout=10,
            creationflags=0x08000000 if os.name == 'nt' else 0
        )
        if result.returncode == 0 and result.stdout:
            with open(screenshot_path, 'wb') as f:
                f.write(result.stdout)
            print(f"📸 Screenshot disimpan: {screenshot_path}")
            return screenshot_path
    except Exception as e:
        print(f"⚠️ Gagal screenshot: {e}")
    return None

def show_notification(message, screenshot_path=None):
    """Tampilkan Windows toast notification."""
    try:
        from ctypes import windll
        windll.user32.MessageBeep(0x00000040)  # MB_ICONINFORMATION beep
    except:
        pass

    # Coba Windows toast via PowerShell
    try:
        ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$textNodes = $template.GetElementsByTagName("text")
$textNodes.Item(0).AppendChild($template.CreateTextNode("Govem Bot")) | Out-Null
$textNodes.Item(1).AppendChild($template.CreateTextNode("{message}")) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($template)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Govem Bot").Show($toast)
'''
        subprocess.run(["powershell", "-Command", ps_script],
                       capture_output=True, timeout=10,
                       creationflags=0x08000000 if os.name == 'nt' else 0)
        print(f"🔔 Notifikasi terkirim: {message}")
    except Exception as e:
        print(f"⚠️ Notifikasi gagal (non-critical): {e}")

    # Buka screenshot otomatis
    if screenshot_path and os.path.exists(screenshot_path):
        try:
            os.startfile(screenshot_path)
        except:
            pass

def minimize_emulator(idx):
    """Minimize window emulator via ldconsole."""
    try:
        # Pakai sortWnd untuk minimize
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
        # Fallback: coba via window title
        if os.name == 'nt':
            subprocess.run(
                ["powershell", "-Command",
                 f'(Get-Process -Name LDPlayer -ErrorAction SilentlyContinue | Where-Object {{$_.MainWindowTitle -match "LDPlayer"}}).MainWindowHandle | ForEach-Object {{ [void][Win32]::ShowWindow($_, 6) }}'],
                capture_output=True, timeout=5,
                creationflags=0x08000000
            )
    except:
        pass

# --- HYBRID RUNNER ---
def run_hybrid_automation(idx, background_mode=True, skip_nav=False):
    if DRY_RUN:
        print("\n⏸️ === DRY RUN MODE === (tidak ada ADB yang dieksekusi)")

    print(f"\n{'🔇' if background_mode else '📱'} Mode: {'BACKGROUND' if background_mode else 'FOREGROUND'}")

    serial = connect_adb_smart(idx, launch_if_needed=(not background_mode))

    global _CACHED_SERIAL
    _CACHED_SERIAL[idx] = serial

    is_suami = (idx == 0)
    wd = datetime.datetime.today().weekday()

    activity_texts = []
    is_jumat = False

    if is_suami:
        result = get_suami_activities(wd)
        activity_texts = result[0]
        is_jumat = result[1]
    else:
        print("ℹ Logic Istri belum disetujui detailnya. Skip.")
        return

    if not activity_texts:
        print("❌ Tidak ada daftar aktivitas yang ditemukan.")
        return

    print(f"\n🚀 MEMULAI PENGISIAN {len(activity_texts)} AKTIVITAS")
    if is_jumat:
        print("📅 Mode JUMAT: Step 3.2 akan digunakan (SKP aktivitas 6)")
    if background_mode and not DRY_RUN:
        print("💡 LDPlayer bisa di-minimize, script tetap berjalan!")

    if DRY_RUN:
        print("\n✅ [DRY RUN] Preview selesai. Tidak ada yang dieksekusi.")
        return

    # LOOP ITEMS
    for i, act in enumerate(activity_texts):
        text = act["teks"]
        skp_num = act["skp"]
        jenis_num = act["jenis"]

        print(f"\n📝 [{i+1}/{len(activity_texts)}] Mengisi: {text[:30]}...")

        # STEP 1: NAVIGASI ke Form (HANYA iterasi pertama)
        # Setelah save, form auto-reset — langsung isi tanpa navigasi ulang
        if i == 0 and not skip_nav:
            print("   🧭 Navigasi Dashboard → Form...")
            adb_click(serial, STEP1_TAP_1[0], STEP1_TAP_1[1])
            time.sleep(2)
            adb_click(serial, STEP1_TAP_2[0], STEP1_TAP_2[1])
            time.sleep(3)
        elif i == 0 and skip_nav:
            print("   ⏩ Skip navigasi (sudah di form)")
            time.sleep(1)
        else:
            print("   (Form auto-reset, langsung isi)")
            time.sleep(1)

        # STEP 2: Focus Input (direct ADB tap)
        adb_click(serial, STEP2_INPUT_XY[0], STEP2_INPUT_XY[1])
        time.sleep(0.8)

        # INPUT TEKS
        print(f"   ⌨️ Mengetik...")
        
        # CLEAR FIELD: HOME + Shift+END (select all), then delete
        for _ in range(2):
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "123"]) # END
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "KEYCODE_MOVE_HOME"])
            run_command(f'"{ADB}" -s {serial} shell input keyevent --longpress 112 123') # Shift+END (select all)
            time.sleep(0.1)
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "67"]) # DEL
            time.sleep(0.1)
        time.sleep(0.5)
        
        clean_text = text.replace("'", "").replace('"', "")
        adb_input_text(serial, clean_text, idx)
        time.sleep(2.5)

        # Hide Keyboard
        run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(3.0)

        # STEP 3: Pilih SKP (ADB tap langsung)
        print(f"   📋 Memilih SKP {skp_num}...")
        adb_click(serial, SKP_DROPDOWN_XY[0], SKP_DROPDOWN_XY[1])
        time.sleep(2.0)

        if skp_num >= 5:
            # SKP 5/6 perlu scroll
            print("   📜 Scroll dropdown...")
            run_command([ADB, "-s", serial, "shell", "input", "swipe", "400", "700", "400", "400", "300"])
            time.sleep(0.8)
            # Setelah scroll, SKP 6 muncul di posisi ~790
            adb_click(serial, 400, 790)
        else:
            coords = SKP_COORDS[skp_num]
            adb_click(serial, coords[0], coords[1])
        time.sleep(2.5) # Increased from 2.0

        # STEP 4: Pilih Jenis (ADB tap langsung)
        jenis_label = "Apel" if jenis_num == 2 else "Aktifitas"
        print(f"   🎯 Memilih Jenis {jenis_num} ({jenis_label})...")
        adb_click(serial, JENIS_DROPDOWN_XY[0], JENIS_DROPDOWN_XY[1])
        time.sleep(2.5) # Increased from 2.0
        jenis_coords = JENIS_COORDS[jenis_num]
        adb_click(serial, jenis_coords[0], jenis_coords[1])
        time.sleep(2.5)

        # STEP 5: Simpan (direct ADB tap)
        print("   💾 Simpan & Tunggu Animasi (6s)...")
        adb_click(serial, STEP5_SAVE_XY[0], STEP5_SAVE_XY[1])
        time.sleep(6) # SMART FILL DELAY

    print(f"\n✅ SEMUA AKTIVITAS SELESAI! ({len(activity_texts)} item)")

def main():
    print("🤖 GOVEM HYBRID AUTOMATION (V23 — Smart Calendar v1.3i-speedup)")
    if DRY_RUN:
        print("⏸️ DRY RUN MODE: Hanya preview, tanpa eksekusi ADB\n")

    print("1. Test Run Suami (Manual Trigger)")
    print("2. Test Run Istri (Manual Trigger)")

    # Auto Mode Argument
    target_idx = -1
    if "--auto" in sys.argv:
        if "--index" in sys.argv:
            try:
                i = sys.argv.index("--index")
                target_idx = int(sys.argv[i+1])
            except: pass

        if target_idx != -1:
            run_hybrid_automation(target_idx)
            return

    choice = input("Pilihan (1/2): ")
    if choice == '1':
        run_hybrid_automation(0)
    elif choice == '2':
        run_hybrid_automation(1)

if __name__ == "__main__":
    main()
