import os
import time
import subprocess
import datetime
import schedule
import configparser
import json
import glob
import re
import math
import sys
import logging
import threading
import importlib

# Global Lock untuk File History (Mencegah Race Condition saat tulis paralel)
HISTORY_LOCK = threading.Lock()

# Global Set untuk Per-User Disable (Toggle dari Tray Menu)
# PERSISTEN: Disimpan ke file agar tetap aktif/nonaktif walau laptop di-restart
DISABLED_USERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "disabled_users.json")

def _load_disabled_users():
    """Load daftar user yang dinonaktifkan dari file."""
    if os.path.exists(DISABLED_USERS_FILE):
        try:
            with open(DISABLED_USERS_FILE, 'r') as f:
                data = json.load(f)
                return set(data) if isinstance(data, list) else set()
        except:
            return set()
    return set()

def _save_disabled_users():
    """Simpan daftar user yang dinonaktifkan ke file."""
    try:
        with open(DISABLED_USERS_FILE, 'w') as f:
            json.dump(list(DISABLED_USERS), f)
    except Exception as e:
        print(f"⚠️ Gagal simpan disabled_users.json: {e}")

# Load dari file saat modul di-import (Persist across restart!)
DISABLED_USERS = _load_disabled_users()

def toggle_user(name):
    """Toggle status aktif/nonaktif satu user. Return True jika sekarang AKTIF."""
    if name in DISABLED_USERS:
        DISABLED_USERS.discard(name)
        logger.info(f"✅ [{name}] DIAKTIFKAN kembali.")
        _save_disabled_users()
        return True
    else:
        DISABLED_USERS.add(name)
        logger.info(f"⏸️ [{name}] DINONAKTIFKAN.")
        _save_disabled_users()
        return False

def is_user_enabled(name):
    """Cek apakah user aktif (tidak ada di DISABLED_USERS)."""
    return name not in DISABLED_USERS


# Setup Logging ke File
LOG_FILE = os.path.join(os.path.dirname(__file__), "govem_scheduler.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()  # Juga print ke console
    ]
)
logger = logging.getLogger(__name__)

LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")

# Konfigurasi App
PACKAGE_NAME = "go.id.tapinkab.govem"
EMULATOR_INDEX = 0  # 0 = Suami, 1 = Istri (Sesuaikan nanti)

# File Config untuk menyimpan koordinat
CONFIG_FILE = "govem_config.ini"
# File History untuk mencatat status absen harian
HISTORY_FILE = "attendance_history.json"

# Konfigurasi Multi-User
# Days: 0=Senin, 1=Selasa, ... 4=Jumat, 5=Sabtu, 6=Minggu
USERS = [
    {'name': 'Suami', 'index': 0, 'gps': True, 'days': [0, 1, 2, 3, 4], 'port': None},      # Senin-Jumat
    {'name': 'Istri', 'index': 1, 'gps': True, 'days': [0, 1, 2, 3, 4, 5], 'port': None},    # Senin-Sabtu
    {'name': 'Pancingan', 'index': 2, 'gps': False, 'days': [0, 1, 2, 3, 4, 5, 6], 'port': None} # SETIAP HARI (Pemicu Iklan)
]

# --- HISTORY MANAGEMENT (STATE AWARENESS) ---
def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except: return {}
    return {}

def save_history_entry(name, period):
    # Period: 'pagi' or 'sore'
    with HISTORY_LOCK: # Thread-safe access
        history = load_history()
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    
    if name not in history: history[name] = {}
    if today_str not in history[name]: history[name][today_str] = {}
    
    history[name][today_str][period] = True
    
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=4)
        
def is_already_done(name, period):
    history = load_history()
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    
    # Cek apakah record hari ini ada?
    if name in history and today_str in history[name]:
        return history[name][today_str].get(period, False)
    
    return False

# Sequence Langkah 1-4 (Dari rekaman user)
# Step 1: x=9600, y=2964
# Step 2: x=1644, y=8916
# Step 3: x=7092, y=9792
# Step 4: x=9696, y=9972
# Kita konversi ke Pixel (Asumsi max 32767)
W_RES = 1600
H_RES = 900

# Default Scale (bisa berubah setelah kalibrasi)
MAX_RAW_X = 32767.0
MAX_RAW_Y = 32767.0

def raw_to_pixel(r_x, r_y):
    # Gunakan scale dari config jika ada
    global MAX_RAW_X, MAX_RAW_Y
    config = load_config()
    if config.has_option("CALIBRATION", "max_x"):
        MAX_RAW_X = float(config.get("CALIBRATION", "max_x"))
        MAX_RAW_Y = float(config.get("CALIBRATION", "max_y"))
    
    px = int((r_x / MAX_RAW_X) * W_RES)
    py = int((r_y / MAX_RAW_Y) * H_RES)
    return px, py

# Sequence Pre-Pagi akan di-recalc saat runtime agar ikut scale terbaru
# Format: (RawX, RawY, "action")
# Action: "click" atau "long_press"
RAW_SEQUENCE_PAGI = [
    # (9600, 2964, "click"),       # Step 1 (REMOVED: Rawan salah klik jika icon geser)
    (1644, 8916, "click"),       # Step 2
    (7092, 9300, "long_press"),  # Step 3 (Adjusted: 9792 -> 9300 agar aman dari Nav Bar)
    (9696, 9972, "click")        # Step 4
]



# Sequence Pre-Sore (Step 3 beda, Step 5 beda)
RAW_SEQUENCE_SORE = [
    # (9600, 2964, "click"),       # Step 1 (REMOVED: Rawan salah klik jika icon geser)
    (1644, 8916, "click"),       # Step 2
    (13524, 9828, "long_press"), # Step 3 (Khusus Sore)
    (9696, 9972, "click"),       # Step 4
    (9840, 10020, "click")       # Step 5 (Final Absen Sore)
]

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

def run_command(command, timeout=30): # Default timeout 30 detik
    try:
        # Suppress Window (No Flashing)
        creationflags = 0
        if os.name == 'nt':
            creationflags = 0x08000000 # CREATE_NO_WINDOW
            
        # Jika command adalah list, jalankan tanpa shell=True (Lebih aman untuk quoting)
        if isinstance(command, list):
            proc = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=False, creationflags=creationflags)
        else:
            proc = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True, creationflags=creationflags)
        
        try:
            stdout, stderr = proc.communicate(timeout=timeout)
            return stdout.strip()
        except subprocess.TimeoutExpired:
            # FORCE KILL seluruh process tree (penting untuk ldconsole yang suka hang)
            pid = proc.pid
            try:
                if os.name == 'nt':
                    # taskkill /F /T /PID: kill seluruh tree proses (child + parent)
                    subprocess.run(f'taskkill /F /T /PID {pid}', shell=True, 
                                   creationflags=0x08000000, timeout=5,
                                   capture_output=True)
                else:
                    proc.kill()
            except Exception:
                pass
            # Bersihkan pipe dengan timeout (jangan sampai hang juga)
            try:
                proc.communicate(timeout=3)
            except Exception:
                pass
            logger.info(f"⚠️ Command Timeout ({timeout}s), KILLED PID {pid}: {command}")
            return ""
    except Exception as e:
        logger.info(f"Error executing command: {e}")
        return ""

def launch_emulator(idx, on_boot_callback=None):
    logger.info(f"🚀 Memulai Emulator Index {idx}...")
    run_command(f'"{LDCONSOLE}" launch --index {idx}')
    # Minimize AGRESIF: spam minimize tanpa jeda agar window tidak sempat muncul
    for _ in range(3):
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
        time.sleep(0.2)

    # AUTO-MINIMIZE: kirim minimize tiap loop sambil tunggu boot
    logger.info(f"⏳ [Emu {idx}] Menunggu sistem Android siap (LDConsole)...")
    boot_ready = False
    max_retries = 30  # 30 x 2s = 60 detik max

    for i in range(max_retries):
        # Minimize setiap iterasi agar window tidak muncul
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
        try:
            cmd_boot = [LDCONSOLE, "adb", "--index", str(idx), "--command", "shell getprop sys.boot_completed"]
            res_boot = run_command(cmd_boot, timeout=10)

            if "1" in res_boot:
                logger.info(f"✅ [Emu {idx}] Sistem Siap! Boot Time: {i*2}s (minimized)")
                # Callback segera setelah boot (misal: kill Pancingan)
                if on_boot_callback:
                    on_boot_callback()
                    on_boot_callback = None  # Hanya sekali
                # Extra minimize selama tunggu launcher stabil
                for _ in range(5):
                    time.sleep(2)
                    run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
                boot_ready = True
                break
        except Exception as e:
            logger.info(f"⚠️ [Emu {idx}] Boot check error: {e}")

        time.sleep(2)
        if i % 5 == 0:
            logger.info(f"   ... [Emu {idx}] Waiting Boot ({i*2}s)")

    if not boot_ready:
        logger.info(f"⚠️ [Emu {idx}] Warning: Boot Detect Timeout. Melanjutkan (hope for the best).")

    # Final minimize
    run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
    
    # MINIMIZE via ldconsole sortWnd
    try:
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize', timeout=10)
        logger.info(f"🔽 [Emu {idx}] Window diminimize via sortWnd.")
    except Exception as e:
        logger.info(f"⚠️ [Emu {idx}] Minimize gagal: {e}")
    
    # Dismiss popup/iklan
    dismiss_popup(idx)

def dismiss_popup(idx):
    """
    Ganti logika dismiss popup dengan DELAY MANUAL 60 DETIK.
    Mencegah salah klik (seperti klik search bar) dan memberi waktu user
    untuk menutup iklan secara manual jika diperlukan.
    """
    logger.info(f"[{idx}] Menunggu 60 detik untuk stabilisasi & tutup iklan manual...")
    time.sleep(60) 
    logger.info(f"[{idx}] Waktu tunggu selesai. Lanjut...")


def find_latest_record(search_path):
    files = glob.glob(os.path.join(search_path, "*.record"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def parse_record_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Cari event klik (PutMultiTouch)
        # Struktur: operations -> [ {Points: [{x, y}]} ]
        # Biasanya LDPlayer menyimpan koordinat dalam skala 0-32767 (input event standard)
        # Tapi kita perlu cek resolutionWidth di file
        
        res_w = data.get('recordInfo', {}).get('resolutionWidth', 1600)
        res_h = data.get('recordInfo', {}).get('resolutionHeight', 900)
        
        last_x = -1
        last_y = -1
        
        ops = data.get('operations', [])
        for op in ops:
            if op.get('operationId') == 'PutMultiTouch':
                points = op.get('points', [])
                if points and points[0].get('state') == 1: # 1 = Down, 0 = Up
                    # Ambil koordinat raw
                    raw_x = points[0].get('x')
                    raw_y = points[0].get('y')
                    
                    # Konversi dari skala 32767 ke Pixel
                    # Rumus: Pixel = (Raw / 32767) * Resolution
                    # Note: Kadang LDPlayer pakai 0-10000 atau raw pixel tergantung versi.
                    # Kita asumsi 0-32767 dulu (standard input device android)
                    
                    # Cek kewajaran: Jika raw > res, pasti butuh scaling
                    # Cek kewajaran: Jika raw > res, pasti butuh scaling
                    # Kita pakai constant sementara untuk deteksi, tapi nanti return raw saja sembari info ke user
                    if raw_x > res_w or raw_y > res_h:
                         # Ini raw coordinate
                         pass
                         
                    last_x = raw_x  # Return nilai RAW agar bisa dikalibrasi
                    last_y = raw_y
                    
        return last_x, last_y
                    # Kita ambil klik terakhir agar user bisa rekam "klik salah, lalu klik benar"
                    
        return last_x, last_y
            
    except Exception as e:
        print(f"Error parsing json: {e}")
        return -1, -1

def import_coordinates_wizard():
    print("\n🕵️ IMPORT KOORDINAT DARI REKAMAN LDPLAYER")
    path = r"D:\LDPlayer\LDPlayer9\vms\operationRecords"
    
    # Ambil semua file .record
    files = glob.glob(os.path.join(path, "*.record"))
    
    if not files:
        print("❌ Tidak ada file rekaman ditemukan di folder default.")
        return

    # Sort by time desc (Terbaru diatas)
    files.sort(key=os.path.getmtime, reverse=True)
    
    print(f"Ditemukan {len(files)} file rekaman:")
    for i, f in enumerate(files[:10]): # Tampilkan 10 terbaru
        fname = os.path.basename(f)
        ftime = datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%Y-%m-%d %H:%M')
        print(f"{i+1}. {fname} ({ftime})")
        
    choice = input("\nPilih nomor file (atau 'x' batal): ")
    if choice.lower() == 'x': return
    
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(files):
            target_file = files[idx]
            print(f"📄 Menganalisa: {os.path.basename(target_file)}")
            
            raw_x, raw_y = parse_record_file(target_file)
            
            if raw_x != -1:
                # Konversi RAW ke Pixel
                px, py = raw_to_pixel(raw_x, raw_y)
                
                print(f"\n📍 Raw Data: {raw_x}, {raw_y}")
                print(f"🎯 Converted Pixel: {px}, {py} (Screen: {W_RES}x{H_RES})")
                
                simpan = input("Simpan sebagai koordinat apa? (1=Pagi, 2=Sore, x=Batal): ")
                if simpan == '1':
                    save_config("COORDS", "pagi_x", px)
                    save_config("COORDS", "pagi_y", py)
                    print("✅ Tersimpan untuk PAGI.")
                elif simpan == '2':
                    save_config("COORDS", "sore_x", px)
                    save_config("COORDS", "sore_y", py)
                    print("✅ Tersimpan untuk SORE.")
            else:
                print("❌ Tidak ada data klik (touch) di file rekaman itu.")
        else:
            print("❌ Nomor tidak valid.")
    except ValueError:
        print("❌ Input angka tidak valid.")

def calibrate_resolution_wizard():
    print("\n📏 WIZARD KALIBRASI SKALA")
    print("Langkah kerja:")
    print("1. Buka Operation Recorder di LDPlayer.")
    print("2. Start Record -> Klik SEKALI saja di POJOK KANAN BAWAH layar emulator (Warna Hitam/Kosong tidak apa-apa).")
    print("   Usahakan se-pojok mungkin!")
    print("3. Stop Record.")
    print("4. Tekan Enter di sini untuk membaca file tersebut.")
    
    input("Tekan Enter jika sudah merekam klik pojok...")
    
    path = r"D:\LDPlayer\LDPlayer9\vms\operationRecords"
    latest_file = find_latest_record(path)
    
    if latest_file:
        print(f"📄 Menganalisa file: {os.path.basename(latest_file)}")
        raw_x, raw_y = parse_record_file(latest_file)
        
        if raw_x != -1:
            print(f"📍 Terbaca Koordinat Ujung: X={raw_x}, Y={raw_y}")
            # Kita anggap ini adalah 100% width dan 100% height
            
        if raw_x != -1:
            print(f"📍 Terbaca Koordinat Ujung: X={raw_x}, Y={raw_y}")
            # Kita anggap ini adalah 100% width dan 100% height
            
            save_config("CALIBRATION", "max_x", raw_x)
            save_config("CALIBRATION", "max_y", raw_y)
            
            print(f"✅ Kalibrasi tersimpan! Scale Factor baru: X={raw_x}, Y={raw_y}")
            print("Sekarang coba import ulang koordinat / test langkah.")
        else:
             print("❌ Gagal membaca klik.")
    else:
        print("❌ File rekaman tidak ditemukan.")

def open_app(idx):
    logger.info(f"📱 [Emu {idx}] Membuka Aplikasi {PACKAGE_NAME}...")
    
    # RETRY LOGIC: 2 attempts total
    for attempt in range(2):
        if attempt > 0:
            logger.info(f"🔄 [Emu {idx}] RETRY #{attempt}: Mencoba buka ulang...")
        
        # 1. FORCE KILL DULU (Agar Fresh Start)
        run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
        time.sleep(2)
        
        # 2. RUN APP
        run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
        logger.info(f"⏳ [Emu {idx}] Menunggu Aplikasi muncul di layar...")
        
        # 3. CEK FOCUS VIA LDCONSOLE (ANTI NYASAR)
        app_ready = False
        max_checks = 15 if attempt == 0 else 10  # 30s pertama, 20s retry
        for i in range(max_checks):
            cmd_focus = [LDCONSOLE, "adb", "--index", str(idx), "--command", "shell dumpsys window windows"]
            res = run_command(cmd_focus, timeout=10)
            
            if PACKAGE_NAME in res:
                logger.info(f"✅ [Emu {idx}] Aplikasi {PACKAGE_NAME} terdeteksi aktif!")
                time.sleep(3)  # Buffer render UI
                return True
            
            time.sleep(2)
            if i % 5 == 0: logger.info(f"   ... [Emu {idx}] Loading App ({i*2}s)")
        
        logger.info(f"⚠️ [Emu {idx}] Attempt {attempt+1}: Aplikasi belum terdeteksi.")
    
    # GAGAL TOTAL setelah 2 attempts
    logger.info(f"❌ [Emu {idx}] GAGAL membuka {PACKAGE_NAME} setelah 2 percobaan. ABORT.")
    return False

def set_location(user):
    # Cek apakah user ini butuh set lokasi?
    if not user['gps']:
        logger.info(f"⏩ [Emu {user['index']}] Skip Set Lokasi (Sesuai Config).")
        return

    idx = user['index']
    config = load_config()
    
    # Priority: LOCATION_{idx} -> LOCATION (Legacy) -> Hardcoded
    section_name = f"LOCATION_{idx}"
    
    if config.has_section(section_name):
        lng = config.get(section_name, "longitude")
        lat = config.get(section_name, "latitude")
        logger.info(f"📍 [Emu {idx}] Menggunakan Config Spesifik User {idx}")
    else:
        # Fallback to shared location
        lng = config.get("LOCATION", "longitude", fallback="115.1625796")
        lat = config.get("LOCATION", "latitude", fallback="-2.9338875")
        logger.info(f"📍 [Emu {idx}] Menggunakan Config Global/Default")
    
    # Auto-save ke section spesifik jika belum ada (Migrasi)
    if not config.has_section(section_name):
        save_config(section_name, "longitude", lng)
        save_config(section_name, "latitude", lat)
    
    if lng and lat:
        logger.info(f"📍 Mengunci GPS [Emu {idx}] ke: {lng}, {lat}")
        cmd = [LDCONSOLE, "locate", "--index", str(idx), "--LLI", f"{lng},{lat}"]
        run_command(cmd)
        time.sleep(2)
        logger.info(f"✅ [Emu {idx}] Lokasi terkunci.")
    else:
        logger.info(f"⚠️ [Emu {idx}] Lokasi error/kosong.")

def click(x, y, idx):
    logger.info(f"👉 [Emu {idx}] Klik ke: {x}, {y}")
    
    # HANYA VIA LDCONSOLE --index (PASTI AKURAT, tidak bisa nyasar ke emulator lain)
    cmd = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input tap {x} {y}"]
    run_command(cmd)

def long_press(x, y, idx, duration_ms=2000):
    logger.info(f"👇 [Emu {idx}] Long Press di: {x}, {y} selama {duration_ms}ms")
    
    # HANYA VIA LDCONSOLE --index (PASTI AKURAT)
    cmd = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input swipe {x} {y} {x} {y} {duration_ms}"]
    run_command(cmd)

def verify_attendance(idx):
    logger.info(f"🔄 [Emu {idx}] VERIFIKASI: Restart Aplikasi...")
    run_command([LDCONSOLE, "killapp", "--index", str(idx), "--packagename", PACKAGE_NAME])
    time.sleep(2)
    open_app(idx)
    logger.info(f"✅ [Emu {idx}] Aplikasi dibuka ulang.")



def play_sequence_steps():
    print("\n▶️ MENJALANKAN 4 LANGKAH AWAL (TEST)")
    launch_emulator(0)
    # open_app() # User bilang rekaman dari awal buka emulator, jadi mungkin manual?
    # Tapi aman kita open_app dulu untuk memastikan focus
    
    print("Akan menjalankan 4 langkah. Pastikan Emulator di Homepage (Awal).")
    input("Tekan Enter untuk mulai...")
    
    # Play Sequence hanya untuk testing index 0 (Suami) dulu
    # Atau bisa ditanya mau test punya siapa, tapi default 0 saja
    idx = 0
    print("Akan menjalankan 4 langkah di Emulator 0 (Suami).")
    input("Tekan Enter untuk mulai...")
    
    launch_emulator(idx)
    # Fake User for test
    dummy_user = {'name': 'Suami', 'index': 0, 'gps': True}
    set_location(dummy_user)
    
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_PAGI):
        sx, sy = raw_to_pixel(rx, ry)
        
        if action == "long_press":
            long_press(sx, sy, idx)
        else:
            click(sx, sy, idx)
            
        time.sleep(3)
    
    print("✅ 4 Langkah selesai.")
    print("Sekarang harusnya sudah di halaman konfirmasi Absen?")

def run_diagnostics():
    print("\n🕵️ DIAGNOSA MASALAH")
    print("1. Cek Status Emulator via LDConsole...")
    res = run_command([LDCONSOLE, "list2"])
    print(f"Output: {res}")
    
    if "0,Suami" in res:
        parts = res.split(',')
        if len(parts) > 4:
            running = parts[4]
            print(f"Status Running: {running} (1=Jalan, 0=Stop)")
            if running == '0':
                print("⚠️ SYSTEM MENDETEKSI EMULATOR MATI!")
                print("Solusi: Tutup Emulator, lalu biarkan Script yang membukanya.")
    
    print("\n2. Cek Koneksi ADB...")
    res_adb = run_command([ADB, "devices"])
    print(f"Output ADB:\n{res_adb}")
    
    print("\n3. Cek Port Terbuka (Netstat)...")
    try:
        # Cari PID dnplayer
        task = subprocess.run('tasklist /FI "IMAGENAME eq dnplayer.exe"', capture_output=True, text=True, shell=True)
        print(task.stdout)
    except: pass
    
    input("Tekan Enter untuk kembali...")

    
def calibration_mode():
    # Helper untuk mengambil screenshot dan melihat mouse position sebenarnya agak susah via console script murni
    # Jadi kita pakai metode 'Trial & Error' sengan ADB input
    
    while True:
        print("\nMenu:")
        print("1. 🕵️ IMPORT DARI REKAMAN (Untuk Step 5 / Final)")
        print("2. ▶️  TEST 4 LANGKAH AWAL (Cek apakah sampai halaman absen)")
        print("3. Set Manual Koordinat SORE")
        print("4. Start Scheduler (Otomatis)")
        print("5. 🚑 DIAGNOSA (Cek Error)")
        print("6. 📏 KALIBRASI SKALA (Steps Awal)")
        print("7. 📍 SET MANUAL LOKASI (Anti-Geser)")
        print("8. 🏃 TEST PENGISIAN AKTIVITAS (Integrated V23)")
        print("9. Keluar")
        
        choice = input("Pilihan (1-9): ")
        
        if choice == '1':
            import_coordinates_wizard()
        
        elif choice == '2':
            play_sequence_steps()
        
        elif choice == '8':
            # Test Activity
            u_idx = int(input("Emulator Index (0=Suami, 1=Istri): "))
            u_name = "Suami" if u_idx == 0 else "Istri"
            dummy_u = {'name': u_name, 'index': u_idx}
            trigger_activity(dummy_u)
            
        elif choice == '3':
            x = input("Masukkan X baru untuk SORE: ")
            y = input("Masukkan Y baru untuk SORE: ")
            save_config("COORDS", "sore_x", x)
            save_config("COORDS", "sore_y", y)
            print("💾 Tersimpan!")
            
        elif choice == '4':
            print("🕒 Scheduler berjalan... (Tekan Ctrl+C untuk stop)")
            schedule.every().day.at("06:30").do(absen_pagi)
            schedule.every().day.at("17:30").do(absen_sore)
            
            while True:
                schedule.run_pending()
                time.sleep(1)
        
        elif choice == '5':
            run_diagnostics()

        elif choice == '6':
            calibrate_resolution_wizard()

        elif choice == '7':
            print("\n📍 SET MANUAL LOKASI (Anti-Geser)")
            print("Mau setting lokasi untuk siapa?")
            print("1. Suami")
            print("2. Istri")
            who = input("Pilihan (1/2): ")
            
            target_section = "LOCATION_0" # Default Suami
            if who == '2':
                target_section = "LOCATION_1"
                
            lat = input(f"Masukkan LATITUDE {target_section} (Contoh: -3.4258): ")
            lng = input(f"Masukkan LONGITUDE {target_section} (Contoh: 114.8394): ")
            
            save_config(target_section, "latitude", lat)
            save_config(target_section, "longitude", lng)
            print(f"✅ Lokasi {target_section} tersimpan!")
        
        elif choice == '8':
            break

# --- INTEGRATED ACTIVITY AUTOMATION (V23 Logic Ported to V22) ---

def parse_text_from_py(filepath):
    texts = []
    try:
        if not os.path.exists(filepath):
            print(f"❌ File tidak ditemukan: {filepath}")
            return []
        with open(filepath, 'r') as f:
            content = f.read()
            matches = re.findall(r'pyautogui\.write\((?:"|\')(.+?)(?:"|\')\)', content)
            texts = matches
    except Exception as e:
        print(f"Error parsing activity script: {e}")
    return texts

def get_activity_list(user_name, day_idx):
    # Base Dirs
    base_suami = r"D:\Download\Aktivitas Govem\Aktivitas"
    base_istri = r"D:\Download\Aktivitas Govem\Aktivitas Istri"
    
    target_dir = base_suami if user_name == 'Suami' else base_istri
    
    # Map Day Index to Filename
    # 0=Senin, 4=Jumat, dst.
    filename = ""
    if user_name == 'Suami':
        if day_idx == 4: filename = "Aktivitas Govem (Jumat1).py"
        else: filename = "Aktivitas Govem (Rutinitas).py"
    else: # Istri
        days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        day_name = days[day_idx]
        filename = f"Aktivitas {day_name}.py"
        
    full_path = os.path.join(target_dir, filename)
    print(f"📖 Membaca aktivitas dari: {filename}")
    
    activities = parse_text_from_py(full_path)
    
    # Logic Tambahan: Apel Sore untuk Suami Senin-Kamis
    if user_name == 'Suami' and day_idx < 4 and activities:
        activities.append("Melaksanakan apel sore")
        
    return activities

def adb_input_text_safe(idx, text):
    # Gunakan LDConsole input text (lebih aman daripada raw adb shell input text karena handling spasi)
    # Tapi LDConsole input text kadang error spasi juga.
    # Kita coba escape spasi dengan %s jika pakai raw ADB via LDConsole shell
    escaped = text.replace(" ", "%s")
    cmd = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input text {escaped}"]
    run_command(cmd)

def run_activity_automation(user_obj):
    idx = user_obj['index']
    name = user_obj['name']
    
    print(f"\n🤖 MEMULAI PENGISIAN AKTIVITAS ({name}) - Integrated V22 Engine")
    
    # 1. Tentukan Aktivitas Hari Ini
    hari_ini = datetime.datetime.now().weekday()
    activities = get_activity_list(name, hari_ini)
    
    if not activities:
        print("❌ Tidak ada aktivitas ditemukan untuk hari ini.")
        return

    print(f"📋 Ditemukan {len(activities)} aktivitas. Mulai eksekusi...")
    
    # 2. Loop Aktivitas
    # Asumsi: Sudah login dan di dashboard (karena dipanggil setelah Absen)
    
    # RECORDING NAMES (Harus Exact!)
    REC_NAV = "Step 1 (Menuju Dashboard Isian).record"
    REC_PRE_TYPE = "Step 2 (Klik untuk supaya bisa mengetikcopas kata2).record"
    REC_SKP_APEL = "Step 3 (Membuka dropdown SKP dan memilih aktivitas nomor 5).record"
    REC_SKP_UMUM = "Step 3.1 (Membuka dropdown SKP dan memilih aktivitas nomor 4).record"
    REC_JENIS_UMUM = "Step 4 (Memilih Jenis Aktivitas Nomor 1).record"
    REC_JENIS_APEL = "Step 4.1 (Memilih Jenis Aktivitas Nomor 2).record"
    REC_SAVE = "Step 5 (Posting Aktivitas).record"
    
    # Helper Play Record via DIRECT ADB (Bypass LDConsole)
    def play(rec_name):
        # 1. Cari File
        rec_name_no_ext = rec_name.replace(".record", "")
        records_dir = r"D:\LDPlayer\LDPlayer9\vms\operationRecords"
        filepath = os.path.join(records_dir, rec_name)
        
        if not os.path.exists(filepath):
            print(f"   ❌ Record tidak ditemukan: {rec_name}")
            return

        print(f"   ▶️ Macro (ADB): {rec_name_no_ext}")
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            ops = data.get('operations', [])
            
            # HARDCODED CALIBRATION (19092 -> 1600)
            w_res = 1600
            h_res = 900
            max_in_x = 19092 
            max_in_y = 10728
            
            last_timing = 0
            
            for op in ops:
                timing = op.get('timing', 0)
                op_id = op.get('operationId')
                
                # Hitung delay
                delay = (timing - last_timing) / 1000.0
                if delay > 0:
                    time.sleep(delay)
                last_timing = timing
                
                if op_id == 'PutMultiTouch':
                    points = op.get('points', [])
                    if points and points[0].get('state') == 1:
                        raw_x = points[0].get('x')
                        raw_y = points[0].get('y')
                        
                        # CALIBRATION MATH
                        real_x = int(raw_x / max_in_x * w_res)
                        real_y = int(raw_y / max_in_y * h_res)
                        
                        # EXECUTE VIA LDCONSOLE (ANTI NYASAR)
                        cmd = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input swipe {real_x} {real_y} {real_x} {real_y} 150"]
                        run_command(cmd)
                        
        except Exception as e:
            print(f"   ❌ Gagal replay macro: {e}")
            
        time.sleep(1) # Buffer antar macro

    # Loop
    for i, activity_text in enumerate(activities):
        print(f"\n📝 [{i+1}/{len(activities)}] Mengisi: {activity_text}")
        
        # Step 1: Navigasi ke Form (Hanya jika belum di form?)
        # Tapi record Step 1 itu "Menuju Dashboard Isian".
        # Asumsi: Setiap kali habis save, kita balik ke dashboard atau form reset?
        # Biasanya setelah save, balik ke list. Jadi butuh Step 1 lagi.
        play(REC_NAV)
        time.sleep(3)
        
        # Step 2: Klik Field Input
        play(REC_PRE_TYPE)
        time.sleep(1)
        
        # Input Text
        print(f"   ⌨️ Mengetik: {activity_text}")
        adb_input_text_safe(idx, activity_text)
        time.sleep(1)
        
        # Step 3 & 4 (Hybrid Logic)
        is_apel = (i == 0) or (i == len(activities) - 1)
        
        if is_apel:
            # Apel logic
            play(REC_SKP_APEL)
            time.sleep(2)
            play(REC_JENIS_APEL)
        else:
            # Umum logic
            play(REC_SKP_UMUM)
            time.sleep(2)
            play(REC_JENIS_UMUM)
            
        time.sleep(1)
        
        # Step 5: Save
        play(REC_SAVE)
        print("   💾 Simpan...")
        time.sleep(4) # Tunggu loading simpan
        
    print(f"✅ Semua aktivitas {name} selesai!")

SCREENSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots")

def _take_dashboard_screenshot(user):
    """Restart app (kembali ke dashboard depan) lalu screenshot ke folder screenshots/ dengan timestamp."""
    idx = user['index']
    name = user['name']
    try:
        # Restart app agar kembali ke dashboard depan
        run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
        time.sleep(2)
        run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
        time.sleep(10)  # Tunggu dashboard depan load

        os.makedirs(SCREENSHOT_DIR, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        screenshot_path = os.path.join(SCREENSHOT_DIR, f"govem_{name.lower()}_{timestamp}.png")
        serial = _get_serial(idx)
        if serial:
            result = subprocess.run(
                [ADB, "-s", serial, "exec-out", "screencap", "-p"],
                capture_output=True, timeout=10,
                creationflags=0x08000000 if os.name == 'nt' else 0
            )
            if result.returncode == 0 and result.stdout:
                with open(screenshot_path, 'wb') as f:
                    f.write(result.stdout)
                logger.info(f"📸 [{name}] Screenshot dashboard: {screenshot_path}")
                return screenshot_path
    except Exception as e:
        logger.error(f"⚠️ [{name}] Screenshot gagal: {e}")
    return None

def _send_batch_notification(job_type, screenshot_paths):
    """Kirim 1 notifikasi Windows toast + buka semua screenshot."""
    names = [os.path.basename(p).replace("govem_final_", "").replace(".png", "").capitalize()
             for p in screenshot_paths if p]
    summary = ", ".join(names) if names else "Selesai"
    try:
        ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$textNodes = $template.GetElementsByTagName("text")
$textNodes.Item(0).AppendChild($template.CreateTextNode("Govem Bot - {job_type}")) | Out-Null
$textNodes.Item(1).AppendChild($template.CreateTextNode("{summary} selesai")) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($template)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Govem Bot").Show($toast)
'''
        subprocess.run(["powershell", "-Command", ps_script],
                       capture_output=True, timeout=10,
                       creationflags=0x08000000 if os.name == 'nt' else 0)
        logger.info(f"🔔 Notifikasi {job_type}: {summary}")
    except Exception as e:
        logger.error(f"⚠️ Notifikasi gagal: {e}")
    # Buka semua screenshot
    for p in screenshot_paths:
        if p and os.path.exists(p):
            try:
                os.startfile(p)
            except:
                pass

def _get_serial(idx):
    """Get ADB serial for emulator index."""
    possible_ports = [5554 + (idx*2), 5556, 5558, 5560]
    devices = subprocess.run([ADB, "devices"], capture_output=True, text=True,
                             creationflags=0x08000000 if os.name == 'nt' else 0).stdout
    for p in possible_ports:
        if f"emulator-{p}" in devices:
            return f"emulator-{p}"
        if f"127.0.0.1:{p}" in devices:
            return f"127.0.0.1:{p}"
    return None

def trigger_activity(user_obj):
    """
    Trigger pengisian aktivitas untuk Suami menggunakan V23_aktivitas_Suami.py
    v1.1: importlib.reload() untuk hindari stale module cache,
          restart app sebelum isi untuk clean state.
    """
    idx = user_obj['index']
    name = user_obj['name']

    print(f"\n📝 [{name}] MEMULAI PENGISIAN AKTIVITAS (V23 Engine)")

    # RESTART APP dulu agar mulai dari Dashboard bersih
    logger.info(f"🔄 [{name}] Restart app sebelum isi aktivitas...")
    run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(3)
    run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(25)  # Tunggu app siap

    try:
        # Import + reload untuk hindari stale cache
        import V23_aktivitas_Suami as v23
        importlib.reload(v23)
        v23.run_hybrid_automation(idx, background_mode=True)
        print(f"✅ [{name}] Pengisian Aktivitas Selesai (V23)")
    except Exception as e:
        print(f"❌ [{name}] Gagal menjalankan V23: {e}")
        # Fallback ke versi lama jika V23 gagal
        print(f"   Mencoba fallback ke versi lama...")
        run_activity_automation(user_obj)

def trigger_activity_istri(user_obj):
    """
    Trigger pengisian aktivitas untuk Istri menggunakan V23_aktivitas_Istri.py
    v1.1: importlib.reload() + restart app (sama seperti Suami)
    Note: V23_aktivitas_Istri sudah handle restart app sendiri,
          tapi kita tetap reload module untuk hindari stale cache.
    """
    idx = user_obj['index']
    name = user_obj['name']

    print(f"\n📝 [{name}] MEMULAI PENGISIAN AKTIVITAS (V23 Istri Engine)")

    try:
        # Import + reload untuk hindari stale cache
        import V23_aktivitas_Istri as v23_istri
        importlib.reload(v23_istri)

        # === OVERRIDE HARI TEMPORER ===
        # 25-28 Maret 2026: Istri isi aktivitas Rabu (weekday=2)
        # Hapus blok ini setelah 28 Maret 2026
        import datetime as _dt
        _today = _dt.date.today()
        _override = None
        if _dt.date(2026, 3, 25) <= _today <= _dt.date(2026, 3, 28):
            _override = 2  # Rabu
            logger.info(f"📌 [{name}] Override hari → Rabu (25-28 Maret 2026)")

        v23_istri.run_istri_automation(background_mode=True, override_hari=_override)
        print(f"✅ [{name}] Pengisian Aktivitas Selesai (V23 Istri)")
    except Exception as e:
        print(f"❌ [{name}] Gagal menjalankan V23 Istri: {e}")

def _re_absen(user, tipe, final_x=None, final_y=None):
    """Ulang absen setelah isi aktivitas (PREVENTIF).
    Emulator sudah jalan, tinggal restart app dan klik ulang sequence absen.
    tipe: 'pagi' atau 'sore'
    """
    name = user['name']
    idx = user['index']
    logger.info(f"🔄 [{name}] RE-ABSEN {tipe.upper()} (Preventif setelah isi aktivitas)")
    
    try:
        # Restart app agar fresh
        run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
        time.sleep(3)
        
        if not open_app(idx):
            logger.info(f"⚠️ [{name}] [Re-Absen] Gagal buka app. Skip re-absen.")
            return
        set_location(user)
        
        if tipe == 'pagi':
            sequence = RAW_SEQUENCE_PAGI
        else:
            sequence = RAW_SEQUENCE_SORE
        
        for i, (rx, ry, action) in enumerate(sequence):
            logger.info(f"   🔄 [{name}] [Re-{tipe.capitalize()}] Step {i+1}...")
            sx, sy = raw_to_pixel(rx, ry)
            if action == "long_press":
                long_press(sx, sy, idx)
            else:
                click(sx, sy, idx)
            time.sleep(5)
        
        # Step Final (khusus pagi yang punya koordinat Step 5)
        if tipe == 'pagi' and final_x:
            logger.info(f"   🔄 [{name}] [Re-Pagi] Step 5 (Final): Click {final_x}, {final_y}")
            click(final_x, final_y, idx)
        
        logger.info(f"✅ [{name}] Re-Absen {tipe.upper()} Selesai (Preventif)")
    except Exception as e:
        logger.error(f"⚠️ [{name}] Re-Absen gagal: {e}", exc_info=True)

def absen_pagi(user):
    name = user['name']
    idx = user['index']
    
    # --- LOGIC PANCINGAN ---
    if name == 'Pancingan':
        # Cek apakah sudah dilakukan hari ini (mencegah catch-up ganda)
        if is_already_done(name, 'pagi'):
            logger.info(f"☕ [{name}] Pancingan Pagi SUDAH SELESAI hari ini. Skip.")
            return
        logger.info(f"🎣 [{name}] MEMANCING IKLAN PAGI (Emulator {idx})")
        launch_emulator(idx)
        logger.info(f"✅ [{name}] Selesai Pancingan. Emulator terbuka untuk menampung iklan.")
        save_history_entry(name, 'pagi')  # Catat agar tidak diulang catch-up
        return # STOP di sini, jangan klik absen.
        
    logger.info(f"☀️ [{name}] MEMULAI ABSEN PAGI (Emulator {idx})")

    config = load_config()
    final_x = config.get("COORDS", "pagi_x", fallback=None)
    final_y = config.get("COORDS", "pagi_y", fallback=None)

    if not final_x:
        logger.info(f"⚠️ [{name}] Koordinat Step 5 belum diset! Hanya jalan langkah 1-4.")

    launch_emulator(idx, on_boot_callback=user.get('_on_boot_callback'))
    if not open_app(idx):
        logger.info(f"❌ [{name}] [Pagi] Gagal membuka aplikasi. Abort.")
        return
    set_location(user)
    
    # Jalankan Sequence 1-4
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_PAGI):
        logger.info(f"   ▶️ [{name}] [Pagi] Step {i+1}...")
        sx, sy = raw_to_pixel(rx, ry)
        if action == "long_press":
             long_press(sx, sy, idx)
        else:
             click(sx, sy, idx)
        time.sleep(5)
    
    # Step 5
    if final_x:
        logger.info(f"   ▶️ [{name}] [Pagi] Step 5 (Final): Click {final_x}, {final_y}")
        click(final_x, final_y, idx)
        logger.info(f"✅ [{name}] Absen Pagi Selesai.")
        
        # CATAT SEJARAH
        save_history_entry(name, 'pagi')
        
        # CHAIN REACTION: Jika Istri, lanjut isi Aktivitas setelah absen pagi
        if name == 'Istri':
            trigger_activity_istri(user)
            # PREVENTIF: Absen pagi ULANG setelah isi aktivitas
            _re_absen(user, 'pagi', final_x, final_y)
            
    else:
        logger.info(f"⚠ [{name}] Step 5 dilewati (Manual).")

def absen_sore(user):
    name = user['name']
    idx = user['index']
    
    # --- LOGIC PANCINGAN ---
    if name == 'Pancingan':
        # Cek apakah sudah dilakukan hari ini (mencegah catch-up ganda)
        if is_already_done(name, 'sore'):
            logger.info(f"☕ [{name}] Pancingan Sore SUDAH SELESAI hari ini. Skip.")
            return
        logger.info(f"🎣 [{name}] MEMANCING IKLAN SORE (Emulator {idx})")
        launch_emulator(idx)
        logger.info(f"✅ [{name}] Selesai Pancingan. Emulator terbuka untuk menampung iklan.")
        save_history_entry(name, 'sore')  # Catat agar tidak diulang catch-up
        return # STOP di sini
        
    logger.info(f"🌙 [{name}] MEMULAI ABSEN SORE (Emulator {idx})")

    launch_emulator(idx, on_boot_callback=user.get('_on_boot_callback'))
    if not open_app(idx):
        logger.info(f"❌ [{name}] [Sore] Gagal membuka aplikasi. Abort.")
        return
    set_location(user)
    
    # Jalankan Sequence SORE
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_SORE):
        logger.info(f"   ▶️ [{name}] [Sore] Step {i+1}...")
        sx, sy = raw_to_pixel(rx, ry)
        if action == "long_press":
             long_press(sx, sy, idx)
        else:
             click(sx, sy, idx)
        time.sleep(5)
    
    logger.info(f"✅ [{name}] Absen Sore Selesai.")
    
    # CATAT SEJARAH
    save_history_entry(name, 'sore')
    
    time.sleep(5)
    verify_attendance(idx)
    
    # CHAIN REACTION: Lanjut isi Aktivitas setelah absen sore (KHUSUS SUAMI)
    if name == 'Suami':
        trigger_activity(user)
        # PREVENTIF: Absen sore ULANG setelah isi aktivitas
        _re_absen(user, 'sore', None, None)
    # Istri: Skip aktivitas, hanya absen saja

# Helper functions job_pagi_daily and job_sore_daily removed (replaced by dynamic lambda).

def get_schedule_rules(name, weekday):
    # Return tuple: (pagi_time_str, sore_time_str) or None if holiday
    # Format Time: "HH:MM"
    
    # === RULES PANCINGAN ===
    # SAMA dengan Suami/Istri agar masuk satu batch (callback kill bekerja)
    # _run_batch_sequential sudah handle urutan: Pancingan → Suami → Istri
    # [NORMAL MODE — pasca Ramadan]
    if name == 'Pancingan':
         # Senin-Kamis: Pagi 06:31, Sore 17:01
         if 0 <= weekday <= 3: return ("06:31", "17:01")
         # Jumat: Pagi 06:31, Sore 11:30
         elif weekday == 4: return ("06:31", "11:30")
         # Sabtu: Pagi 06:31, Sore 15:30 (ikut Istri)
         elif weekday == 5: return ("06:31", "15:30")
         # Minggu: libur
         else: return None

    # === RULES ISTRI ===
    # [NORMAL MODE — pasca Ramadan]
    elif name == 'Istri':
        # Senin - Kamis (0-3): Pagi 06:31, Sore 17:01
        if 0 <= weekday <= 3:
            return ("06:31", "17:01")
        # Jumat (4): Pagi 06:31, Sore 11:30
        elif weekday == 4:
            return ("06:31", "11:30")
        # Sabtu (5): Pagi 06:31, Sore 15:30
        elif weekday == 5:
            return ("06:31", "15:30")
        # Minggu (6)
        else:
            return None # Libur

    # === RULES SUAMI ===
    # [NORMAL MODE — pasca Ramadan]
    elif name == 'Suami':
        # Senin - Kamis (0-3): Pagi 06:31, Sore 17:01
        if 0 <= weekday <= 3:
            return ("06:31", "17:01")
        # Jumat (4): Pagi 06:31, Sore 11:30
        elif weekday == 4:
            return ("06:31", "11:30")
        # Sabtu - Minggu (5-6)
        else:
            return None # Libur
            
    return None

def reset_today_history():
    """Hapus status absensi hari ini (Start Fresh)."""
    try:
        if not os.path.exists(HISTORY_FILE):
            return
            
        with open(HISTORY_FILE, 'r') as f:
            history = json.load(f)
            
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        
        changed = False
        for user in history:
            if today_str in history[user]:
                del history[user][today_str]
                changed = True
                print(f"♻️ Reset history hari ini untuk {user}")
                
        if changed:
            with open(HISTORY_FILE, 'w') as f:
                json.dump(history, f, indent=4)
            print("✅ History hari ini berhasil dihapus (Start Fresh).")
            
    except Exception as e:
        print(f"❌ Gagal reset history: {e}")

def run_scheduler(target_users, force_reset=False):
    print("🕒 Scheduler berjalan... (Tekan Ctrl+C untuk stop)")
    
    if force_reset:
        print("\n⚡ FORCE RESET REQUESTED: Menghapus status absensi hari ini...")
        reset_today_history()
        
    print(f"👥 Target User: {', '.join([u['name'] for u in target_users])}")
    
    # 1. REGISTER SCHEDULER JOBS
    # Helper wrapper agar variable 'user' ter-bind dengan benar di closure
    # URUTAN LAUNCH: Pancingan(2) -> Suami(0) -> Istri(1)
    # Pancingan HARUS pertama agar menyerap iklan, mencegah iklan muncul di Suami/Istri
    LAUNCH_ORDER = ['Pancingan', 'Suami', 'Istri']
    
    def _run_batch_sequential(job_type, users_with_rules):
        """Jalankan absen SEQUENTIAL (satu per satu) dengan urutan Pancingan -> Suami -> Istri.
        Ini mencegah ldconsole hang karena 3 emulator launch bersamaan."""
        # Sort users sesuai LAUNCH_ORDER
        sorted_users = sorted(users_with_rules, key=lambda u: LAUNCH_ORDER.index(u['name']) if u['name'] in LAUNCH_ORDER else 99)

        screenshots = []
        pancingan_index = None  # Simpan index Pancingan untuk deferred kill

        def _kill_pancingan_callback():
            """Callback: kill Pancingan 3s setelah emulator berikutnya boot."""
            nonlocal pancingan_index
            if pancingan_index is not None:
                time.sleep(3)
                logger.info(f"🔌 [Pancingan] Auto-kill emulator (iklan sudah terserap)")
                run_command(f'"{LDCONSOLE}" quit --index {pancingan_index}')
                pancingan_index = None

        for u in sorted_users:
            if not is_user_enabled(u['name']):
                logger.info(f"⏸️ [Job] SKIP {job_type} {u['name']} (Dinonaktifkan)")
                continue
            try:
                # Inject callback kill Pancingan ke launch_emulator user berikutnya
                if u['name'] != 'Pancingan' and pancingan_index is not None:
                    u['_on_boot_callback'] = _kill_pancingan_callback

                logger.info(f"🚀 [Job] Memulai {job_type}: {u['name']}")
                if job_type == 'Absen Pagi':
                    absen_pagi(u)
                else:
                    absen_sore(u)
                logger.info(f"✅ [Job] Selesai {job_type}: {u['name']}")
                if u['name'] == 'Pancingan':
                    pancingan_index = u['index']
                else:
                    ss = _take_dashboard_screenshot(u)
                    if ss:
                        screenshots.append(ss)
                    # Auto-kill emulator setelah screenshot selesai
                    logger.info(f"🔌 [{u['name']}] Auto-kill emulator (tugas selesai)")
                    run_command(f'"{LDCONSOLE}" quit --index {u["index"]}')
            except Exception as e:
                logger.error(f"💀 [CRASH] {job_type} {u['name']} GAGAL: {e}", exc_info=True)
        # Safety: kill Pancingan jika belum di-kill (misal semua user lain skip/gagal)
        if pancingan_index is not None:
            logger.info(f"🔌 [Pancingan] Auto-kill emulator (cleanup)")
            run_command(f'"{LDCONSOLE}" quit --index {pancingan_index}')

        # 1 notifikasi gabungan + buka semua screenshot
        if screenshots:
            _send_batch_notification(job_type, screenshots)
    
    # Kumpulkan users per (pagi_time, sore_time) agar yang jadwal sama dilaunched bersama-sama
    from collections import defaultdict
    pagi_groups = defaultdict(list)  # time_str -> [user, ...]
    sore_groups = defaultdict(list)
    
    
    for user in target_users:
        name = user['name']
        
        logger.info(f"📅 Jadwal {name}:")
        for day_idx in range(7):
            rules = get_schedule_rules(name, day_idx)
            if rules:
                pagi_t, sore_t = rules
                day_name = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"][day_idx]
                logger.info(f"   - {day_name}: Pagi {pagi_t}, Sore {sore_t}")
                
                # Kumpulkan user ke group berdasarkan waktu
                pagi_groups[(day_idx, pagi_t)].append(user)
                sore_groups[(day_idx, sore_t)].append(user)
    
    # Register schedule: satu job per (hari, waktu) yang menjalankan semua users secara sequential
    def get_day_scheduler(day_idx):
        """Return scheduler object untuk hari tertentu."""
        if day_idx == 0: return schedule.every().monday
        elif day_idx == 1: return schedule.every().tuesday
        elif day_idx == 2: return schedule.every().wednesday
        elif day_idx == 3: return schedule.every().thursday
        elif day_idx == 4: return schedule.every().friday
        elif day_idx == 5: return schedule.every().saturday
        return None
    
    for (day_idx, time_str), users in pagi_groups.items():
        sched = get_day_scheduler(day_idx)
        if sched:
            user_names = ', '.join(u['name'] for u in users)
            logger.info(f"   📋 Pagi {time_str} ({['Sen','Sel','Rab','Kam','Jum','Sab'][day_idx]}): {user_names}")
            captured_users = list(users)  # capture untuk closure
            def make_pagi_batch(u_list):
                def job():
                    threading.Thread(target=_run_batch_sequential, args=('Absen Pagi', u_list)).start()
                return job
            sched.at(time_str).do(make_pagi_batch(captured_users))
    
    for (day_idx, time_str), users in sore_groups.items():
        sched = get_day_scheduler(day_idx)
        if sched:
            user_names = ', '.join(u['name'] for u in users)
            logger.info(f"   📋 Sore {time_str} ({['Sen','Sel','Rab','Kam','Jum','Sab'][day_idx]}): {user_names}")
            captured_users = list(users)  # capture untuk closure
            def make_sore_batch(u_list):
                def job():
                    threading.Thread(target=_run_batch_sequential, args=('Absen Sore', u_list)).start()
                return job
            sched.at(time_str).do(make_sore_batch(captured_users))

    logger.info("✅ Jadwal Terdaftar. Menunggu waktu eksekusi...")

    # 2. FITUR LATE ARRIVAL (Catch-Up)
    skrg = datetime.datetime.now()
    wd = skrg.weekday()
    
    catchup_pagi = []  # Kumpulkan user yang perlu catch-up
    catchup_sore = []
    
    for user in target_users:
        rules = get_schedule_rules(user['name'], wd)
        if rules:
            pagi_t_str, sore_t_str = rules
            
            # --- CEK PAGI ---
            h_p, m_p = map(int, pagi_t_str.split(':'))
            target_pagi = skrg.replace(hour=h_p, minute=m_p, second=0, microsecond=0)
            batas_pagi = skrg.replace(hour=11, minute=0, second=0, microsecond=0)
            
            if target_pagi < skrg < batas_pagi:
                 # MASIH DALAM WINDOW PAGI (sebelum 11:00)
                 if is_already_done(user['name'], 'pagi'):
                     logger.info(f"☕ [{user['name']}] Absen Pagi SUDAH SELESAI hari ini. (Skip Catch-Up)")
                 else:
                     logger.info(f"⚠️ DETEKSI TELAT STARTUP ({user['name']})!!")
                     logger.info(f"   Jadwal: {pagi_t_str}, Sekarang: {skrg.strftime('%H:%M')}")
                     catchup_pagi.append(user)
            elif skrg >= batas_pagi:
                 # WINDOW PAGI SUDAH LEWAT (>= 11:00)
                 if not is_already_done(user['name'], 'pagi'):
                     logger.info(f"⏰ [{user['name']}] Window Pagi sudah lewat ({skrg.strftime('%H:%M')} > 11:00).")
                     logger.info(f"   📝 Auto-mark Pagi sebagai 'Done (Skipped)' agar history bersih.")
                     save_history_entry(user['name'], 'pagi')

            # --- CEK SORE ---
            h_s, m_s = map(int, sore_t_str.split(':'))
            target_sore = skrg.replace(hour=h_s, minute=m_s, second=0, microsecond=0)
            batas_sore = skrg.replace(hour=20, minute=0, second=0, microsecond=0)
            
            if target_sore < skrg < batas_sore:
                 # MASIH DALAM WINDOW SORE (sebelum 20:00)
                 if is_already_done(user['name'], 'sore'):
                     logger.info(f"🍵 [{user['name']}] Absen Sore SUDAH SELESAI hari ini. (Skip Catch-Up)")
                 else:
                     logger.info(f"⚠️ DETEKSI TELAT STARTUP ({user['name']})!!")
                     logger.info(f"   Jadwal: {sore_t_str}, Sekarang: {skrg.strftime('%H:%M')}")
                     catchup_sore.append(user)
            elif skrg >= batas_sore:
                 # WINDOW SORE SUDAH LEWAT (>= 20:00)
                 if not is_already_done(user['name'], 'sore'):
                     logger.info(f"⏰ [{user['name']}] Window Sore sudah lewat ({skrg.strftime('%H:%M')} > 20:00).")
                     logger.info(f"   📝 Auto-mark Sore sebagai 'Done (Skipped)' agar history bersih.")
                     save_history_entry(user['name'], 'sore')
    
    # JALANKAN CATCH-UP SECARA SEQUENTIAL (Pancingan dulu, baru Suami, baru Istri)
    if catchup_pagi:
        logger.info(f"🚀 CATCH-UP PAGI: {', '.join(u['name'] for u in catchup_pagi)}")
        threading.Thread(target=_run_batch_sequential, args=('Absen Pagi', catchup_pagi)).start()
    if catchup_sore:
        logger.info(f"🚀 CATCH-UP SORE: {', '.join(u['name'] for u in catchup_sore)}")
        threading.Thread(target=_run_batch_sequential, args=('Absen Sore', catchup_sore)).start()
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("🛑 Script Dihentikan User.")
        time.sleep(2)

# Library Mode: No Main Function
# main() removed. Use Govem_Suami.py or Govem_Istri.py