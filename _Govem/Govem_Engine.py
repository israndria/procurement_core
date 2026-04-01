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

# Global Lock untuk ADB kill-server (GLOBAL operation — tidak boleh paralel)
# Ketika 2 emulator berjalan bersamaan dan keduanya timeout, keduanya akan
# mencoba kill-server. Tanpa lock, mereka saling membunuh koneksi satu sama lain.
ADB_KILL_SERVER_LOCK = threading.Lock()

# Path ke aktivitas_trigger.py (di folder V23_Aktivitas_Streamlit)
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
AKTIVITAS_TRIGGER = os.path.join(_SCRIPT_DIR, "..", "V23_Aktivitas_Streamlit", "aktivitas_trigger.py")
AKTIVITAS_TRIGGER = os.path.normpath(AKTIVITAS_TRIGGER)
PYTHON_EXE = r"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe"

# Absolute path ke folder script — dipakai semua path config/log/history

# Global Set untuk Per-User Disable (Toggle dari Tray Menu)
# PERSISTEN: Disimpan ke file agar tetap aktif/nonaktif walau laptop di-restart
DISABLED_USERS_FILE = os.path.join(_SCRIPT_DIR, "disabled_users.json")

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
        logger.error(f"Gagal simpan disabled_users.json: {e}")

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
LOG_FILE = os.path.join(_SCRIPT_DIR, "govem_scheduler.log")
_log_handlers = [logging.FileHandler(LOG_FILE, encoding='utf-8')]
# StreamHandler hanya jika ada console asli (bukan pythonw/devnull)
if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
    import io
    _sh = logging.StreamHandler(io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace'))
    _log_handlers.append(_sh)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=_log_handlers
)
logger = logging.getLogger(__name__)

LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")

# Konfigurasi App
PACKAGE_NAME = "go.id.tapinkab.govem"
EMULATOR_INDEX = 0  # 0 = Suami, 1 = Istri (Sesuaikan nanti)

# File Config untuk menyimpan koordinat
# PENTING: Gunakan _SCRIPT_DIR (absolute) agar tidak bergantung CWD
# (CWD bisa berbeda saat dijalankan via pythonw/autostart/tray)
CONFIG_FILE = os.path.join(_SCRIPT_DIR, "govem_config.ini")
# File History untuk mencatat status absen harian
HISTORY_FILE = os.path.join(_SCRIPT_DIR, "attendance_history.json")

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
    with HISTORY_LOCK:  # Thread-safe: read + modify + write dalam satu lock
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

def launch_emulator(idx, on_boot_callback=None, safe_mode=False):
    logger.info(f"🚀 Memulai Emulator Index {idx}...")

    # RAM CHECK: Pastikan cukup sebelum launch (penting untuk 8 GB RAM)
    # safe_mode=True saat paralel: hanya flush, jangan kill proses (bisa kill emulator saudara)
    try:
        from ram_optimizer import ensure_ram_available, flush_standby_ram, get_ram_report
        logger.info(get_ram_report())
        if safe_mode:
            flush_standby_ram()
        else:
            ensure_ram_available(min_mb=1500, max_wait_seconds=30, aggressive=True)
    except Exception as e:
        logger.warning(f"⚠️ RAM optimizer error (lanjut tanpa optimasi): {e}")

    # Ensure cleanMode 1 (Disable Ads/Popups) — crucial for stabilization
    run_command(f'"{LDCONSOLE}" globalsetting --cleanmode 1', timeout=10)

    run_command(f'"{LDCONSOLE}" launch --index {idx}')
    # Minimize sekali setelah launch (cukup 1x, tidak perlu spam)
    time.sleep(1)
    run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')

    # Tunggu boot tanpa spam minimize (window sudah diminimize di atas)
    logger.info(f"⏳ [Emu {idx}] Menunggu sistem Android siap (Direct ADB)...")
    boot_ready = False
    max_retries = 30  # 30 x 2s = 60 detik max

    # PREEMPTIVE: Connect ADB early
    target_serial = _get_adb_serial(idx)
    run_command(f'"{ADB}" connect {target_serial}', timeout=5)

    for i in range(max_retries):
        try:
            # Pakai _direct_adb (lebih reliable daripada ldconsole adb wrapper)
            res_boot = _direct_adb(idx, "getprop sys.boot_completed", timeout=8)

            if res_boot and "1" in res_boot:
                logger.info(f"✅ [Emu {idx}] Sistem Siap! Boot Time: {i*2}s (Direct ADB)")
                if on_boot_callback:
                    on_boot_callback()
                    on_boot_callback = None
                time.sleep(5)  # Tunggu launcher stabil
                boot_ready = True
                break
        except Exception as e:
            if i % 5 == 0: logger.info(f"⚠️ [Emu {idx}] Boot check error: {e}")

        time.sleep(2)
        if i % 10 == 0:
            run_command(f'"{ADB}" connect {target_serial}', timeout=5)
            logger.info(f"   ... [Emu {idx}] Waiting Boot ({i*2}s)")

    if not boot_ready:
        logger.info(f"⚠️ [Emu {idx}] Boot timeout! Restart emulator sekali...")
        run_command(f'"{LDCONSOLE}" quit --index {idx}')
        time.sleep(5)
        run_command(f'"{LDCONSOLE}" launch --index {idx}')
        time.sleep(1)
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
        # Tunggu boot kedua (max 90 detik)
        for i in range(45):
            try:
                res_boot = _direct_adb(idx, "getprop sys.boot_completed", timeout=10)
                if res_boot and "1" in res_boot:
                    logger.info(f"✅ [Emu {idx}] Boot OK setelah restart! ({i*2}s)")
                    if on_boot_callback:
                        on_boot_callback()
                        on_boot_callback = None
                    time.sleep(5)
                    boot_ready = True
                    break
            except:
                pass
            time.sleep(2)
        if not boot_ready:
            logger.info(f"❌ [Emu {idx}] Boot gagal setelah restart. Melanjutkan (hope for the best).")
            # PASTIKAN memicu callback agar antrean tidak macet untuk user berikutnya
            if on_boot_callback:
                on_boot_callback()
                on_boot_callback = None

    # Final minimize (safety net, 1x saja)
    try:
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize', timeout=10)
        logger.info(f"🔽 [Emu {idx}] Window diminimize via sortWnd.")
    except Exception as e:
        logger.info(f"⚠️ [Emu {idx}] Minimize gagal: {e}")
    
    # Dismiss popup/iklan
    dismiss_popup(idx)

def dismiss_popup(idx):
    """
    Tunggu stabilisasi sebentar. cleanMode sudah aktif via globalsetting,
    jadi tidak butuh 60s lagi. 15s cukup untuk OS siap terima input.
    """
    logger.info(f"[{idx}] Stabilisasi UI (15 detik)...")
    time.sleep(15) 
    logger.info(f"[{idx}] Siap.")


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

def restart_adb_bridge(idx):
    """Restart ADB bridge untuk emulator tertentu — fix hang ADB daemon."""
    logger.info(f"🔧 [Emu {idx}] Restart ADB bridge...")
    port = 5554 + idx * 2  # Emu 0=5554, Emu 1=5556, Emu 2=5558
    serial = f"emulator-{port}"
    addr = f"127.0.0.1:{port + 1}"  # ADB connect pakai port+1 (5555, 5557, 5559)
    run_command(f'"{ADB}" disconnect {serial}', timeout=5)
    run_command(f'"{ADB}" disconnect {addr}', timeout=5)
    time.sleep(1)
    result = run_command(f'"{ADB}" connect {addr}', timeout=5)
    if not result:
        # connect timeout → ADB server dalam state stale (e.g. setelah resume from sleep)
        # kill-server untuk reset total — HARUS PAKAI LOCK karena kill-server adalah
        # operasi GLOBAL yang memutus semua emulator. Tanpa lock, 2 emulator paralel
        # akan saling membunuh koneksi satu sama lain secara bergantian.
        logger.info(f"⚠️ [Emu {idx}] ADB connect timeout — kill-server & retry...")
        with ADB_KILL_SERVER_LOCK:
            logger.info(f"🔒 [Emu {idx}] Acquired kill-server lock...")
            run_command(f'"{ADB}" kill-server', timeout=10)
            time.sleep(2)
            run_command(f'"{ADB}" connect {addr}', timeout=15)
            logger.info(f"🔓 [Emu {idx}] Released kill-server lock.")
    time.sleep(2)
    logger.info(f"✅ [Emu {idx}] ADB bridge restarted ({addr}).")

def _get_adb_serial(idx):
    """Dapatkan ADB serial langsung (bypass ldconsole adb wrapper)."""
    return f"127.0.0.1:{5555 + idx * 2}"

def _direct_adb(idx, cmd, timeout=10):
    """Jalankan ADB shell command langsung ke emulator — lebih reliable dari ldconsole adb."""
    serial = _get_adb_serial(idx)
    result = run_command(f'"{ADB}" -s {serial} shell {cmd}', timeout=timeout)
    return result

def check_app_running(idx):
    """Cek apakah app sedang running — via DIRECT ADB (bypass ldconsole).

    ldconsole adb sering hang untuk Emu 1 saat parallel operation.
    Direct ADB jauh lebih reliable.
    """
    # Method 1: dumpsys activity top
    try:
        res = _direct_adb(idx, "dumpsys activity top", timeout=8)
        if res and PACKAGE_NAME in res:
            logger.info(f"   [Emu {idx}] App terdeteksi via dumpsys activity top")
            return True
    except Exception:
        pass

    # Method 2: pidof
    try:
        res2 = _direct_adb(idx, f"pidof {PACKAGE_NAME}", timeout=5)
        if res2 and res2.strip():
            parts = res2.strip().split()
            if any(p.isdigit() for p in parts):
                logger.info(f"   [Emu {idx}] App terdeteksi via pidof (PID: {res2.strip()})")
                return True
    except Exception:
        pass

    # Method 3: ps grep (note: pipe doesn't work in split, use single command)
    try:
        serial = _get_adb_serial(idx)
        cmd3 = f'"{ADB}" -s {serial} shell ps -A | grep {PACKAGE_NAME}'
        res3 = run_command(cmd3, timeout=8)
        if res3 and PACKAGE_NAME in res3:
            logger.info(f"   [Emu {idx}] App terdeteksi via ps grep")
            return True
    except Exception:
        pass

    return False

def open_app(idx):
    logger.info(f"📱 [Emu {idx}] Membuka Aplikasi {PACKAGE_NAME}...")

    # PREEMPTIVE: Restart ADB bridge (fix hang ADB daemon — terutama Emu 1)
    restart_adb_bridge(idx)

    # RETRY LOGIC: 3 attempts total
    for attempt in range(3):
        if attempt > 0:
            logger.info(f"🔄 [Emu {idx}] RETRY #{attempt}: Mencoba buka ulang...")
            # Restart ADB bridge setiap retry
            restart_adb_bridge(idx)

        # 1. FORCE KILL DULU (Agar Fresh Start)
        run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
        time.sleep(2)

        # 2. RUN APP
        run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
        logger.info(f"⏳ [Emu {idx}] Menunggu Aplikasi muncul di layar...")

        # 3. CEK APP RUNNING (pakai method ringan, bukan dumpsys window)
        max_checks = 20 if attempt == 0 else 12
        for i in range(max_checks):
            if check_app_running(idx):
                logger.info(f"✅ [Emu {idx}] Aplikasi {PACKAGE_NAME} terdeteksi aktif!")
                time.sleep(3)  # Buffer render UI
                return True

            time.sleep(2)
            if i % 5 == 0: logger.info(f"   ... [Emu {idx}] Loading App ({i*2}s)")

        logger.info(f"⚠️ [Emu {idx}] Attempt {attempt+1}: Aplikasi belum terdeteksi.")

    # NUCLEAR RETRY: Quit emulator → relaunch → coba 1x lagi
    logger.info(f"💀 [Emu {idx}] NUCLEAR RETRY: Quit emulator & relaunch...")
    run_command(f'"{LDCONSOLE}" quit --index {idx}')
    time.sleep(5)

    # RAM flush sebelum relaunch
    try:
        from ram_optimizer import ensure_ram_available
        ensure_ram_available(min_mb=1500, max_wait_seconds=20, aggressive=True)
    except Exception:
        pass

    run_command(f'"{LDCONSOLE}" launch --index {idx}')
    time.sleep(1)
    run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')

    # Tunggu boot — pakai direct ADB (ldconsole adb unreliable post-relaunch)
    _nuclear_serial = f"127.0.0.1:{5555 + idx * 2}"
    # Reset ADB server total — wajib setelah nuclear quit+launch agar connect tidak stuck
    # (tanpa ini: connect=NO terus-menerus jika server punya stale state dari sesi sebelumnya)
    # PAKAI LOCK: kill-server adalah operasi GLOBAL — tidak boleh paralel dengan emulator lain
    logger.info(f"🔄 [Emu {idx}] Nuclear: reset ADB server (kill-server)...")
    with ADB_KILL_SERVER_LOCK:
        logger.info(f"🔒 [Emu {idx}] Nuclear acquired kill-server lock...")
        run_command(f'"{ADB}" kill-server', timeout=10)
        time.sleep(3)
        logger.info(f"🔓 [Emu {idx}] Nuclear released kill-server lock.")
    # Bersihkan koneksi lama dulu (stale dari sebelum quit)
    run_command(f'"{ADB}" disconnect {_nuclear_serial}', timeout=5)
    for i in range(60):  # 60×2s = 120s (tambah timeout utk 3-emulator scenario)
        try:
            r_connect = run_command(f'"{ADB}" connect {_nuclear_serial}', timeout=5)
            res_boot = run_command(f'"{ADB}" -s {_nuclear_serial} shell getprop sys.boot_completed', timeout=8)
            if res_boot and "1" in res_boot.strip():
                logger.info(f"✅ [Emu {idx}] Boot OK setelah nuclear restart! ({i*2}s)")
                time.sleep(10)
                break
            if i % 5 == 0:
                # Log diagnostic: apakah connect berhasil? apakah getprop ada output?
                conn_ok = r_connect and ("connected" in r_connect or "already" in r_connect)
                boot_val = repr(res_boot) if res_boot else "None"
                logger.info(f"   ... [Emu {idx}] Nuclear Boot ({i*2}s) | connect={'OK' if conn_ok else 'NO'} | boot={boot_val}")
        except Exception as exc:
            if i % 5 == 0:
                logger.info(f"   ... [Emu {idx}] Nuclear Boot ({i*2}s) | exc: {exc}")
        time.sleep(2)
    else:
        logger.info(f"❌ [Emu {idx}] GAGAL boot setelah nuclear restart. ABORT.")
        return False

    run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
    restart_adb_bridge(idx)

    # Final attempt: buka app
    run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(2)
    run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
    logger.info(f"⏳ [Emu {idx}] Nuclear retry: Menunggu app...")

    for i in range(15):
        if check_app_running(idx):
            logger.info(f"✅ [Emu {idx}] Aplikasi AKHIRNYA terdeteksi setelah nuclear restart!")
            time.sleep(3)
            return True
        time.sleep(2)
        if i % 5 == 0: logger.info(f"   ... [Emu {idx}] Nuclear Loading ({i*2}s)")

    logger.info(f"❌ [Emu {idx}] GAGAL membuka {PACKAGE_NAME} setelah nuclear restart. ABORT.")
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

    # HANYA VIA DIRECT ADB (Lebih cepat & reliable dari ldconsole)
    res = _direct_adb(idx, f"input tap {x} {y}", timeout=10)

    # Jika timeout (res kosong), retry 1x dengan refresh bridge
    if res is None or res == "":
        logger.warning(f"⚠️ Klik timeout atau gagal! Refresh bridge & retry...")
        restart_adb_bridge(idx)
        time.sleep(2)
        res2 = _direct_adb(idx, f"input tap {x} {y}", timeout=10)
        # Return False jika retry juga gagal
        if res2 is None:
            return False
    return True

def long_press(x, y, idx, duration_ms=2000):
    logger.info(f"👇 [Emu {idx}] Long Press di: {x}, {y} selama {duration_ms}ms")
    
    # HANYA VIA DIRECT ADB (Lebih cepat & reliable)
    # Gunakan timeout lebih tinggi (15s) untuk swipe agar tidak dianggap hang
    res = _direct_adb(idx, f"input swipe {x} {y} {x} {y} {duration_ms}", timeout=15)

    # Jika timeout (res kosong), retry 1x dengan refresh bridge
    if res is None or res == "":
        logger.warning(f"⚠️ Long press timeout atau gagal! Refresh bridge & retry...")
        restart_adb_bridge(idx)
        time.sleep(2)
        _direct_adb(idx, f"input swipe {x} {y} {x} {y} {duration_ms}", timeout=15)

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
        print("9. Keluar")

        choice = input("Pilihan (1-9): ")
        
        if choice == '1':
            import_coordinates_wizard()
        
        elif choice == '2':
            play_sequence_steps()
        
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


SCREENSHOT_DIR = os.path.join(_SCRIPT_DIR, "screenshots")

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
    """Dapatkan serial ADB yang tepat untuk index emulator (Anti-Bleeding)."""
    # Port standard LDPlayer: Instance 0=5555, Instance 1=5557, dst.
    target_port = 5555 + (idx * 2)
    target_addr = f"127.0.0.1:{target_port}"
    
    # Cek apakah sudah terhubung sebagai list item
    devices = run_command(f'"{ADB}" devices', timeout=5)
    if target_addr in devices and "device" in devices:
        return target_addr
        
    # Jika tidak ada, coba yang format 'emulator-XXXX'
    emu_port = 5554 + (idx * 2)
    if f"emulator-{emu_port}" in devices:
        return f"emulator-{emu_port}"
        
    return target_addr # Fallback

def trigger_aktivitas(person: str):
    """Jalankan pengisian aktivitas via web.govem + Playwright CDP.
    Dijalankan dalam thread terpisah agar tidak memblokir scheduler Engine.
    person: 'Suami' atau 'Istri'
    """
    def _run():
        logger.info(f"🖥️ [{person}] Memulai pengisian aktivitas via web...")
        try:
            result = subprocess.run(
                [PYTHON_EXE, AKTIVITAS_TRIGGER, person],
                capture_output=True, text=True, timeout=300
            )
            if result.stdout:
                for line in result.stdout.strip().splitlines():
                    logger.info(f"   [aktivitas/{person}] {line}")
            if result.stderr:
                for line in result.stderr.strip().splitlines():
                    logger.warning(f"   [aktivitas/{person}] STDERR: {line}")
            if result.returncode == 0:
                logger.info(f"✅ [{person}] Aktivitas web selesai.")
            else:
                logger.error(f"❌ [{person}] Aktivitas web gagal (exit {result.returncode}).")
        except subprocess.TimeoutExpired:
            logger.error(f"❌ [{person}] Aktivitas web timeout (>5 menit).")
        except Exception as e:
            logger.error(f"❌ [{person}] Aktivitas web error: {e}")

    t = threading.Thread(target=_run, daemon=True, name=f"aktivitas-{person}")
    t.start()


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

    launch_emulator(idx, on_boot_callback=user.get('_on_boot_callback'), safe_mode=user.get('_safe_mode', False))
    if not open_app(idx):
        logger.info(f"❌ [{name}] [Pagi] Gagal membuka aplikasi.")
        return
    set_location(user)
    
    # REFRESH BRIDGE: Pastikan koneksi segar sebelum rangkaian langkah dimulai
    restart_adb_bridge(idx)
    time.sleep(2)

    # Jalankan Sequence 1-4
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_PAGI):
        logger.info(f"   ▶️ [{name}] [Pagi] Step {i+1}...")
        sx, sy = raw_to_pixel(rx, ry)
        if action == "long_press":
             long_press(sx, sy, idx)
        else:
             click(sx, sy, idx)
        time.sleep(5)

    # Step 5 — klik tombol absen final, dengan retry jika gagal
    if final_x:
        logger.info(f"   ▶️ [{name}] [Pagi] Step 5 (Final): Click {final_x}, {final_y}")
        ok = click(final_x, final_y, idx)
        if not ok:
            # Retry sekali lagi setelah bridge fresh
            logger.warning(f"⚠️ [{name}] Step 5 gagal! Retry setelah 5s...")
            time.sleep(5)
            ok = click(final_x, final_y, idx)
            if not ok:
                logger.error(f"❌ [{name}] Step 5 TETAP GAGAL setelah retry. Absen mungkin tidak masuk!")
                # Jangan simpan history — biarkan catch-up / sore ulang menangani
                return
        logger.info(f"✅ [{name}] Absen Pagi Selesai.")

        # CATAT SEJARAH
        save_history_entry(name, 'pagi')

        # TRIGGER AKTIVITAS — hanya Istri setelah absen pagi
        if name == 'Istri':
            trigger_aktivitas('Istri')

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

    launch_emulator(idx, on_boot_callback=user.get('_on_boot_callback'), safe_mode=user.get('_safe_mode', False))
    if not open_app(idx):
        logger.info(f"❌ [{name}] [Sore] Gagal membuka aplikasi. Abort.")
        return
    set_location(user)

    # REFRESH BRIDGE: Pastikan koneksi segar sebelum rangkaian langkah dimulai
    restart_adb_bridge(idx)
    time.sleep(2)
    
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

    # TRIGGER AKTIVITAS — hanya Suami setelah absen sore
    if name == 'Suami':
        trigger_aktivitas('Suami')

    time.sleep(5)
    verify_attendance(idx)

# Helper functions job_pagi_daily and job_sore_daily removed (replaced by dynamic lambda).

def get_schedule_rules(name, weekday):
    # Return tuple: (pagi_time_str, sore_time_str) or None if holiday
    # Format Time: "HH:MM"

    # === RULES PANCINGAN ===
    # Pancingan pagi 1 menit sebelum Suami, sore ikut Suami
    # [UPDATE 2026-04-01: sesuai jam kerja baru]
    if name == 'Pancingan':
         # Senin-Kamis: Pagi 07:00 (sebelum Suami 07:01), Sore 17:01
         if 0 <= weekday <= 3: return ("07:00", "17:01")
         # Jumat: Pagi 07:00, Sore 11:30
         elif weekday == 4: return ("07:00", "11:30")
         # Sabtu: Pagi 07:00, Sore 15:00 (ikut Istri)
         elif weekday == 5: return ("07:00", "15:00")
         # Minggu: libur
         else: return None

    # === RULES ISTRI ===
    # [UPDATE 2026-04-01: pagi 07:31, sore 15:00 Sen-Kam, 13:00 Jumat, 15:00 Sabtu]
    elif name == 'Istri':
        # Senin - Kamis (0-3): Pagi 07:31, Sore 15:00
        if 0 <= weekday <= 3:
            return ("07:31", "15:00")
        # Jumat (4): Pagi 07:31, Sore 13:00
        elif weekday == 4:
            return ("07:31", "13:00")
        # Sabtu (5): Pagi 07:31, Sore 15:00
        elif weekday == 5:
            return ("07:31", "15:00")
        # Minggu (6)
        else:
            return None # Libur

    # === RULES SUAMI ===
    # [UPDATE 2026-04-01: pagi 07:01, sore tetap]
    elif name == 'Suami':
        # Senin - Kamis (0-3): Pagi 07:01, Sore 17:01
        if 0 <= weekday <= 3:
            return ("07:01", "17:01")
        # Jumat (4): Pagi 07:01, Sore 11:30
        elif weekday == 4:
            return ("07:01", "11:30")
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
                logger.info(f"♻️ Reset history hari ini untuk {user}")

        if changed:
            with open(HISTORY_FILE, 'w') as f:
                json.dump(history, f, indent=4)
            logger.info("✅ History hari ini berhasil dihapus (Start Fresh).")

    except Exception as e:
        logger.error(f"❌ Gagal reset history: {e}")

def run_scheduler(target_users, force_reset=False):
    logger.info("🕒 Scheduler berjalan...")

    if force_reset:
        logger.info("⚡ FORCE RESET REQUESTED: Menghapus status absensi hari ini...")
        reset_today_history()

    logger.info(f"👥 Target User: {', '.join([u['name'] for u in target_users])}")
    
    # 1. REGISTER SCHEDULER JOBS
    # Helper wrapper agar variable 'user' ter-bind dengan benar di closure
    # URUTAN LAUNCH: Pancingan(2) -> Suami(0) -> Istri(1)
    # Pancingan HARUS pertama agar menyerap iklan, mencegah iklan muncul di Suami/Istri
    LAUNCH_ORDER = ['Pancingan', 'Suami', 'Istri']
    
    def _run_batch(job_type, users_with_rules):
        """Jalankan absen: Pancingan dulu (serap iklan), lalu Suami + Istri PARALEL."""
        sorted_users = sorted(users_with_rules, key=lambda u: LAUNCH_ORDER.index(u['name']) if u['name'] in LAUNCH_ORDER else 99)

        screenshots = []
        screenshots_lock = threading.Lock()
        pancingan_index = None
        pancingan_killed = threading.Event()

        def _kill_pancingan_callback():
            """Callback: kill Pancingan 3s setelah emulator boot. Thread-safe via Event."""
            nonlocal pancingan_index
            if pancingan_index is not None and not pancingan_killed.is_set():
                pancingan_killed.set()  # Cegah double-kill dari thread lain
                time.sleep(3)
                logger.info(f"🔌 [Pancingan] Auto-kill emulator (iklan sudah terserap)")
                run_command(f'"{LDCONSOLE}" quit --index {pancingan_index}')

        # === STEP 1: PANCINGAN (sequential, harus duluan) ===
        pancingan_user = next((u for u in sorted_users if u['name'] == 'Pancingan'), None)
        if pancingan_user and is_user_enabled('Pancingan'):
            try:
                logger.info(f"🚀 [Job] Memulai {job_type}: Pancingan")
                if job_type == 'Absen Pagi':
                    absen_pagi(pancingan_user)
                else:
                    absen_sore(pancingan_user)
                logger.info(f"✅ [Job] Selesai {job_type}: Pancingan")
                pancingan_index = pancingan_user['index']
            except Exception as e:
                logger.error(f"💀 [CRASH] {job_type} Pancingan GAGAL: {e}", exc_info=True)

        # === STEP 2: SUAMI + ISTRI (STAGGER LAUNCH → PARALEL ABSEN) ===
        # PENTING: Pada laptop 8 GB RAM, launch 2 emulator bersamaan → RAM habis → boot fail.
        # Strategi: Boot emulator SATU-SATU (sequential), lalu jalankan absen PARALEL.
        # Gate signal dikirim via on_boot_callback setelah emulator boot selesai.
        parallel_users = [u for u in sorted_users if u['name'] != 'Pancingan' and is_user_enabled(u['name'])]

        # Gate: Emulator kedua baru launch setelah emulator pertama boot
        launch_gate = threading.Event()
        launch_gate.set()  # User pertama boleh langsung launch

        def _make_boot_callback(name):
            """Buat callback yang: (1) kill Pancingan jika perlu, (2) signal gate."""
            def _callback():
                # Kill Pancingan (siapa duluan boot, dia kill)
                if pancingan_index is not None:
                    _kill_pancingan_callback()
                # Buka gate untuk emulator berikutnya
                logger.info(f"🔓 [{name}] Boot selesai → buka gate untuk emulator berikutnya")
                launch_gate.set()
            return _callback

        def _user_worker(user, is_first):
            """Worker thread untuk satu user (Suami atau Istri)."""
            name = user['name']
            try:
                # Safe mode: jangan kill proses saat paralel (bisa kill emulator saudara)
                user['_safe_mode'] = True
                # Inject combined callback (kill Pancingan + gate signal)
                user['_on_boot_callback'] = _make_boot_callback(name)

                # STAGGER: Tunggu giliran launch (user pertama langsung, sisanya tunggu)
                if not is_first:
                    logger.info(f"⏳ [{name}] Menunggu emulator sebelumnya boot dulu (stagger)...")
                    launch_gate.wait(timeout=300)  # Max 5 menit tunggu
                    logger.info(f"🟢 [{name}] Giliran launch!")

                logger.info(f"🚀 [Job] Memulai {job_type}: {name}")
                if job_type == 'Absen Pagi':
                    absen_pagi(user)
                else:
                    absen_sore(user)
                logger.info(f"✅ [Job] Selesai {job_type}: {name}")

                ss = _take_dashboard_screenshot(user)
                if ss:
                    with screenshots_lock:
                        screenshots.append(ss)
                logger.info(f"🔌 [{name}] Auto-kill emulator (tugas selesai)")
                run_command(f'"{LDCONSOLE}" quit --index {user["index"]}')
            except Exception as e:
                logger.error(f"💀 [CRASH] {job_type} {name} GAGAL: {e}", exc_info=True)
                # PENTING: Jika crash, tetap buka gate agar thread berikutnya tidak stuck selamanya
                launch_gate.set()

        threads = []
        for i, u in enumerate(parallel_users):
            is_first = (i == 0)
            if i > 0:
                launch_gate.clear()  # Reset gate: user berikutnya harus tunggu

            t = threading.Thread(target=_user_worker, args=(u, is_first), name=f"Govem-{u['name']}")
            t.start()
            threads.append(t)
            logger.info(f"🧵 [Stagger] Thread {u['name']} started (is_first={is_first})")

            # Tunggu sebentar agar thread pertama mulai launch dulu
            if is_first and len(parallel_users) > 1:
                time.sleep(2)

        # Tunggu semua thread selesai
        for t in threads:
            t.join()

        # Safety: kill Pancingan jika belum di-kill
        if pancingan_index is not None and not pancingan_killed.is_set():
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
                    threading.Thread(target=_run_batch, args=('Absen Pagi', u_list)).start()
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
                    threading.Thread(target=_run_batch, args=('Absen Sore', u_list)).start()
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
            batas_sore = skrg.replace(hour=23, minute=59, second=0, microsecond=0)
            
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
                     logger.info(f"⏰ [{user['name']}] Window Sore sudah lewat ({skrg.strftime('%H:%M')} > 23:59).")
                     logger.info(f"   📝 Auto-mark Sore sebagai 'Done (Skipped)' agar history bersih.")
                     save_history_entry(user['name'], 'sore')
    
    # JALANKAN CATCH-UP SECARA SEQUENTIAL (Pancingan dulu, baru Suami, baru Istri)
    if catchup_pagi:
        logger.info(f"🚀 CATCH-UP PAGI: {', '.join(u['name'] for u in catchup_pagi)}")
        threading.Thread(target=_run_batch, args=('Absen Pagi', catchup_pagi)).start()
    if catchup_sore:
        logger.info(f"🚀 CATCH-UP SORE: {', '.join(u['name'] for u in catchup_sore)}")
        threading.Thread(target=_run_batch, args=('Absen Sore', catchup_sore)).start()
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("🛑 Script Dihentikan User.")
        time.sleep(2)

# Library Mode: No Main Function
# main() removed. Use Govem_Suami.py or Govem_Istri.py