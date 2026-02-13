import os
import time
import subprocess
import datetime
import schedule
import configparser
import json
import glob
import re
import json
import math
import sys
import logging
import threading

# Global Lock untuk File History (Mencegah Race Condition saat tulis paralel)
HISTORY_LOCK = threading.Lock()

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

def run_command(command, timeout=30): # Default timeout 30 detik (Up dari 15s prevent hang di parallel)
    try:
        # Suppress Window (No Flashing)
        startupinfo = None
        creationflags = 0
        if os.name == 'nt':
            creationflags = 0x08000000 # CREATE_NO_WINDOW
            
        # Jika command adalah list, jalankan tanpa shell=True (Lebih aman untuk quoting)
        if isinstance(command, list):
            result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags, timeout=timeout)
        else:
            result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags, timeout=timeout)
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        print(f"⚠️ Command Timeout: {command}")
        return ""
    except Exception as e:
        print(f"Error executing command: {e}")
        return ""

def launch_emulator(idx):
    print(f"🚀 Memulai Emulator Index {idx}...")
    run_command(f'"{LDCONSOLE}" launch --index {idx}')
    
    # Tunggu fixed wait dihapus, ganti Smart Wait (Karena ADB sudah fix)
    # DIAGNOSA & AUTO-DISCOVERY PORT
    # Masalah: Index 1 seringkali tidak di port 5556 jika hasil clone (bisa 5558, 5560, dst)
    print("🔍 Auto-Discovery Port ADB...")
    
    found_port = None
    
    # 1. Coba Scan Port yang mungkin (5556 - 5564)
    possible_ports = [5554 + (idx*2), 5556, 5558, 5560, 5562, 5564]
    if idx > 0 and 5554 in possible_ports: possible_ports.remove(5554)
    
    # Konek ke semua kemungkinan
    for p in possible_ports:
        run_command(f'"{ADB}" connect 127.0.0.1:{p}')
    
    # Smart Wait Loop untuk Deteksi Device Online + Boot Complete
    print("⏳ Menunggu sistem Android siap (Smart Detect)...")
    boot_ready = False
    max_retries = 30 # 30 x 2s = 60 detik max
    
    for i in range(max_retries):
        devices_out = run_command(f'"{ADB}" devices')
        
        # Cari device yang milik index ini
        detected_p = None
        detected_serial = None
        
        # 1. Cek dari lisr devices
        for p in possible_ports:
            # Check loose match
            if str(p) in devices_out and "device" in devices_out:
                detected_p = p
                # Determine serial format from output logic
                if f"emulator-{p}" in devices_out:
                    detected_serial = f"emulator-{p}"
                elif f"127.0.0.1:{p}" in devices_out:
                     detected_serial = f"127.0.0.1:{p}"
                else:
                     detected_serial = f"emulator-{p}" # Default guess
                break
        
        # 2. Fallback: Jika tidak ketemu di list, coba tembak langsung port DEFAULT
        if not detected_p:
            default_p = 5554 + (idx*2)
            detected_serial = f"127.0.0.1:{default_p}" # Kalau invisible, biasanya harus via IP
            # Cek apakah port ini respond shell?
            check_alive = run_command([ADB, "-s", detected_serial, "shell", "echo", "alive"])
            if "alive" in check_alive:
                detected_p = default_p

        if detected_p and detected_serial:
            # Device ketemu/hidup, sekarang cek apakah Boot Selesai?
            cmd_boot = [ADB, "-s", detected_serial, "shell", "getprop", "sys.boot_completed"]
            res_boot = run_command(cmd_boot)
            
            # Jika return 1 = Booted. Jika kosong/error = Masih loading.
            if "1" in res_boot:
                 print(f"✅ Sistem Siap (Port {detected_p}). Boot Time: {i*2}s")
                 print("☕ Menunggu 10 detik agar Launcher stabil...") # Buffer tambahan penting
                 time.sleep(10)
                 found_port = detected_p
                 boot_ready = True
                 break
        
        time.sleep(2)
        if i % 3 == 0: 
            print(f"   ... Detecting ({i*2}s) [ADB Status: {len(devices_out.splitlines())-1} devices]")

    if not boot_ready:
        print("⚠️ Warning: Smart Detect Timeout. Menggunakan default port & hope for the best.")
        found_port = 5554 + (idx*2)

    # Simpan ke Memory Global USERS
    current_user_obj = next((u for u in USERS if u['index'] == idx), None)
    if current_user_obj:
        current_user_obj['port'] = found_port
        # Simpan Serial juga biar konsisten di command selanjutnya
        # Kalau detected_serial belum ada (timeout case), kita buat default
        if not 'detected_serial' in locals() or not detected_serial:
             detected_serial = f"emulator-{found_port}" 
        current_user_obj['serial'] = detected_serial
    
    # Force Connect Final
    run_command(f'"{ADB}" connect 127.0.0.1:{found_port}')
    
    # Dismiss popup/iklan jika ada
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
    print(f"📱 Membuka Aplikasi {PACKAGE_NAME}...")
    
    # 1. FORCE KILL DULU (Agar Fresh Start)
    run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(1)
    
    # 2. RUN APP
    run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
    print("⏳ Menunggu Aplikasi muncul di layar (Smart Wait)...")
    
    # Ambil Port user & Serial
    user_port = 5554 + (idx * 2)
    target_serial = f"127.0.0.1:{user_port}" # Default fallback
    
    current_user_obj = next((u for u in USERS if u['index'] == idx), None)
    if current_user_obj:
        if current_user_obj.get('port'): user_port = current_user_obj['port']
        if current_user_obj.get('serial'): target_serial = current_user_obj['serial']
    
    # Loop Check Focus
    app_ready = False
    for i in range(20): # Max 40 detik
        # Cek Focused App
        # Gunakan target serial yang sudah validated saat boot
        cmd_focus = [ADB, "-s", target_serial, "shell", "dumpsys", "window", "windows"]
        res = run_command(cmd_focus)
        
        # String detection simpel (Looser)
        # Kadang mCurrentFocus ga muncul jelas, tapi kalau package ada di list windows visible, itu sudah cukup.
        if PACKAGE_NAME in res:
             print(f"✅ Aplikasi {PACKAGE_NAME} terdeteksi aktif!")
             app_ready = True
             break
        
        time.sleep(2)
        if i % 5 == 0: print(f"   ... Loading App ({i*2}s)")
        
    if not app_ready:
        print("⚠️ Warning: Aplikasi belum terdeteksi fokus dalam 40 detik.")
        print("⏩ MELANJUTKAN EKSEKUSI (Optimistic Mode) - Asumsi aplikasi sudah terbuka.")
        return True # DULU False, sekarang True agar tidak stuck.
    
    time.sleep(3) # Buffer dikit biar render UI sempurna
    return True

def set_location(user):
    # Cek apakah user ini butuh set lokasi?
    if not user['gps']:
        print(f"⏩ [Emu {user['index']}] Skip Set Lokasi (Sesuai Config).")
        return

    idx = user['index']
    config = load_config()
    
    # Priority: LOCATION_{idx} -> LOCATION (Legacy) -> Hardcoded
    section_name = f"LOCATION_{idx}"
    
    if config.has_section(section_name):
        lng = config.get(section_name, "longitude")
        lat = config.get(section_name, "latitude")
        print(f"📍 Menggunakan Config Spesifik User {idx}")
    else:
        # Fallback to shared location
        lng = config.get("LOCATION", "longitude", fallback="115.1625796")
        lat = config.get("LOCATION", "latitude", fallback="-2.9338875")
        print(f"📍 Menggunakan Config Global/Default")
    
    # Auto-save ke section spesifik jika belum ada (Migrasi)
    if not config.has_section(section_name):
        save_config(section_name, "longitude", lng)
        save_config(section_name, "latitude", lat)
    
    if lng and lat:
        print(f"📍 Mengunci GPS [Emu {idx}] ke: {lng}, {lat}")
        cmd = [LDCONSOLE, "locate", "--index", str(idx), "--LLI", f"{lng},{lat}"]
        run_command(cmd)
        time.sleep(2)
        print("✅ Lokasi terkunci.")
    else:
        print("⚠️ Lokasi error/kosong.")

def click(x, y, idx):
    print(f"👉 [Emu {idx}] Klik ke: {x}, {y}")
    
    # Cara 1: Via LDConsole (Official) - List Mode untuk hindari masalah Quote
    cmd1 = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input tap {x} {y}"]
    run_command(cmd1)
    
    # Cara 2: Via ADB Direct (Backup) - Skip dulu biar simple, fokus LDConsole
    # Cara 2: Via Direct ADB (Backup dengan Port Dinamis)
    # Cek apakah kita punya port hasil discovery?
    user_port = 5554 + (idx * 2) # Default
    
    current_user_obj = next((u for u in USERS if u['index'] == idx), None)
    if current_user_obj and current_user_obj.get('port'):
         user_port = current_user_obj['port']
         # print(f"   [Debug] Menggunakan Port Temuan: {user_port}")

    serial = f"emulator-{user_port}"
    
    # Target 2: Serial Name
    cmd2 = [ADB, "-s", serial, "shell", "input", "tap", str(x), str(y)]
    res2 = run_command(cmd2)
    if res2 and "error" in res2.lower():
         print(f"   ⚠️ [Direct Serial] Error: {res2}")

    # Target 3: IP Address (Network Mode)
    target_ip = f"127.0.0.1:{user_port}"
    cmd3 = [ADB, "-s", target_ip, "shell", "input", "tap", str(x), str(y)]
    res3 = run_command(cmd3)

def long_press(x, y, idx, duration_ms=2000):
    print(f"👇 [Emu {idx}] Long Press di: {x}, {y} selama {duration_ms}ms")
    # Cara 1: LDConsole
    cmd1 = [LDCONSOLE, "adb", "--index", str(idx), "--command", f"shell input swipe {x} {y} {x} {y} {duration_ms}"]
    run_command(cmd1)

    # Cara 2: Direct ADB
    user_port = 5554 + (idx * 2) # Default
    current_user_obj = next((u for u in USERS if u['index'] == idx), None)
    if current_user_obj and current_user_obj.get('port'):
         user_port = current_user_obj['port']

    serial = f"emulator-{user_port}"
    
    cmd2 = [ADB, "-s", serial, "shell", "input", "swipe", str(x), str(y), str(x), str(y), str(duration_ms)]
    run_command(cmd2)

    # Cara 3: Via IP
    target_ip = f"127.0.0.1:{user_port}"
    cmd3 = [ADB, "-s", target_ip, "shell", "input", "swipe", str(x), str(y), str(x), str(y), str(duration_ms)]
    run_command(cmd3)

def verify_attendance(idx):
    print(f"\n🔄 [Emu {idx}] VERIFIKASI: Restart Aplikasi...")
    run_command([LDCONSOLE, "killapp", "--index", str(idx), "--packagename", PACKAGE_NAME])
    time.sleep(2)
    open_app(idx)
    print("✅ Aplikasi dibuka ulang.")



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
            
            # Get User Serial
            user_port = 5554 + (idx * 2) 
            current_user = next((u for u in USERS if u['index'] == idx), None)
            if current_user and current_user.get('serial'): 
                 serial = current_user['serial']
            else:
                 serial = f"emulator-{user_port}" # Fallback
            
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
                        
                        # EXECUTE SWIPE (Click)
                        # print(f"      Click: {real_x}, {real_y}") # Verbose toggle
                        cmd = [ADB, "-s", serial, "shell", "input", "swipe", str(real_x), str(real_y), str(real_x), str(real_y), "150"]
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

def trigger_activity(user_obj):
    """
    Trigger pengisian aktivitas untuk Suami menggunakan V23_aktivitas_Suami.py
    Script V23 sudah di-refine dengan:
    - Fix untuk Jumat (senam pagi + telaah)
    - ADB input text (proven bekerja di background)
    - Koordinat aktivitas 6 yang tepat
    """
    idx = user_obj['index']
    name = user_obj['name']
    
    print(f"\n📝 [{name}] MEMULAI PENGISIAN AKTIVITAS (V23 Engine)")
    
    try:
        # Import dari V23_aktivitas_Suami
        import V23_aktivitas_Suami as v23
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
    """
    idx = user_obj['index']
    name = user_obj['name']
    
    print(f"\n📝 [{name}] MEMULAI PENGISIAN AKTIVITAS (V23 Istri Engine)")
    
    try:
        # Import dari V23_aktivitas_Istri
        import V23_aktivitas_Istri as v23_istri
        v23_istri.run_istri_automation(background_mode=True)
        print(f"✅ [{name}] Pengisian Aktivitas Selesai (V23 Istri)")
    except Exception as e:
        print(f"❌ [{name}] Gagal menjalankan V23 Istri: {e}")

def absen_pagi(user):
    name = user['name']
    idx = user['index']
    
    # --- LOGIC PANCINGAN ---
    if name == 'Pancingan':
        print(f"\n🎣 [{name}] MEMANCING IKLAN PAGI (Emulator {idx})")
        launch_emulator(idx)
        # open_app(idx) # User confirm app belum install, hanya launch emu
        print(f"✅ [{name}] Selesai Pancingan. Emulator terbuka untuk menampung iklan.")
        return # STOP di sini, jangan klik absen.
        
    print(f"\n☀️ [{name}] MEMULAI ABSEN PAGI (Emulator {idx})")
    
    config = load_config()
    final_x = config.get("COORDS", "pagi_x", fallback=None)
    final_y = config.get("COORDS", "pagi_y", fallback=None)

    if not final_x:
        print("⚠️ Koordinat Step 5 belum diset! Hanya jalan langkah 1-4.")
    
    launch_emulator(idx)
    if not open_app(idx):
        print(f"❌ [Pagi] Gagal membuka aplikasi untuk {name}. Abort.")
        return
    set_location(user)
    
    # Jalankan Sequence 1-4
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_PAGI):
        print(f"   ▶️ [Pagi] Step {i+1}...")
        sx, sy = raw_to_pixel(rx, ry)
        if action == "long_press":
             long_press(sx, sy, idx)
        else:
             click(sx, sy, idx)
        time.sleep(5) # Diperlama jadi 5 detik (biar tidak "ketinggalan")
    
    # Step 5
    if final_x:
        print(f"Step 5: Final Click {final_x}, {final_y}")
        click(final_x, final_y, idx)
        print(f"✅ [{name}] Absen Pagi Selesai.")
        
        # CATAT SEJARAH
        save_history_entry(name, 'pagi')
        
        # CHAIN REACTION: Jika Istri, lanjut isi Aktivitas setelah absen pagi
        if name == 'Istri':
            trigger_activity_istri(user)
            
    else:
        print(f"⚠ [{name}] Step 5 dilewati (Manual).")

def absen_sore(user):
    name = user['name']
    idx = user['index']
    
    # --- LOGIC PANCINGAN ---
    if name == 'Pancingan':
        print(f"\n🎣 [{name}] MEMANCING IKLAN SORE (Emulator {idx})")
        launch_emulator(idx)
        # open_app(idx) # User confirm app belum install
        print(f"✅ [{name}] Selesai Pancingan. Emulator terbuka untuk menampung iklan.")
        return # STOP di sini
        
    print(f"\n🌙 [{name}] MEMULAI ABSEN SORE (Emulator {idx})")
    
    launch_emulator(idx)
    if not open_app(idx):
        print(f"❌ [Sore] Gagal membuka aplikasi untuk {name}. Abort.")
        return
    set_location(user)
    
    # Jalankan Sequence SORE
    for i, (rx, ry, action) in enumerate(RAW_SEQUENCE_SORE):
        print(f"   ▶️ [Sore] Step {i+1}...")
        sx, sy = raw_to_pixel(rx, ry)
        if action == "long_press":
             long_press(sx, sy, idx)
        else:
             click(sx, sy, idx)
        time.sleep(5) # Diperlama jadi 5 detik (biar tidak "ketinggalan")
    
    print(f"✅ [{name}] Absen Sore Selesai.")
    
    # CATAT SEJARAH
    save_history_entry(name, 'sore')
    
    time.sleep(5)
    verify_attendance(idx)
    
    # CHAIN REACTION: Lanjut isi Aktivitas setelah absen sore (KHUSUS SUAMI)
    if name == 'Suami':
        trigger_activity(user)
    # Istri: Skip aktivitas, hanya absen saja

# Helper functions job_pagi_daily and job_sore_daily removed (replaced by dynamic lambda).

def get_schedule_rules(name, weekday):
    # Return tuple: (pagi_time_str, sore_time_str) or None if holiday
    # Format Time: "HH:MM"
    
    # === RULES PANCINGAN ===
    # Jalan 5 menit LEBIH AWAL dari jadwal utama
    if name == 'Pancingan':
         # Senin-Kamis: Sore 17:30 (Utama 17:35)
         if 0 <= weekday <= 3: return ("06:30", "17:30")
         # Jumat-Sabtu: Sore 14:00 (Utama 14:05)
         elif weekday >= 4: return ("06:30", "14:00")
         
    # === RULES ISTRI ===
    elif name == 'Istri':
        # Senin - Kamis (0-3): Pagi 06:35, Sore 17:35
        if 0 <= weekday <= 3:
            return ("06:35", "17:35")
        # Jumat (4): Sore 14:05
        elif weekday == 4:
            return ("06:35", "14:05")
        # Sabtu (5): Sore 14:05
        elif weekday == 5:
            return ("06:35", "14:05")
        # Minggu (6)
        else:
            return None # Libur

    # === RULES SUAMI ===
    elif name == 'Suami':
        # Senin - Kamis (0-3): Pagi 06:35, Sore 17:35
        if 0 <= weekday <= 3:
            return ("06:35", "17:35")
        # Jumat (4): Sore 14:05
        elif weekday == 4:
            return ("06:35", "14:05")
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
    def make_job_pagi(u): 
        def job(): 
            # Run Async (Parallel)
            print(f"🚀 [Job] Starting Thread Absen Pagi: {u['name']}")
            threading.Thread(target=absen_pagi, args=(u,)).start()
        return job
    
    def make_job_sore(u): 
        def job(): 
            # Run Async (Parallel)
            print(f"🚀 [Job] Starting Thread Absen Sore: {u['name']}")
            threading.Thread(target=absen_sore, args=(u,)).start()
        return job
    
    for user in target_users:
        name = user['name']
        
        print(f"\n📅 Jadwal {name}:")
        for day_idx in range(7):
            rules = get_schedule_rules(name, day_idx)
            if rules:
                pagi_t, sore_t = rules
                day_name = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"][day_idx]
                print(f"   - {day_name}: Pagi {pagi_t}, Sore {sore_t}")
                
                # Register berdasarkan hari - FRESH schedule object setiap kali
                if day_idx == 0:
                    schedule.every().monday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().monday.at(sore_t).do(make_job_sore(user))
                elif day_idx == 1:
                    schedule.every().tuesday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().tuesday.at(sore_t).do(make_job_sore(user))
                elif day_idx == 2:
                    schedule.every().wednesday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().wednesday.at(sore_t).do(make_job_sore(user))
                elif day_idx == 3:
                    schedule.every().thursday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().thursday.at(sore_t).do(make_job_sore(user))
                elif day_idx == 4:
                    schedule.every().friday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().friday.at(sore_t).do(make_job_sore(user))
                elif day_idx == 5:
                    schedule.every().saturday.at(pagi_t).do(make_job_pagi(user))
                    schedule.every().saturday.at(sore_t).do(make_job_sore(user))

    print("\n✅ Jadwal Terdaftar. Menunggu waktu eksekusi...")

    # 2. FITUR LATE ARRIVAL (Catch-Up)
    skrg = datetime.datetime.now()
    wd = skrg.weekday()
    
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
                     print(f"☕ [{user['name']}] Absen Pagi SUDAH SELESAI hari ini. (Skip Catch-Up)")
                 else:
                     print(f"⚠️ DETEKSI TELAT STARTUP ({user['name']})!!")
                     print(f"   Jadwal: {pagi_t_str}, Sekarang: {skrg.strftime('%H:%M')}")
                     print("   🚀 RUNNING CATCH-UP PAGI (Async)...")
                     threading.Thread(target=absen_pagi, args=(user,)).start()
            elif skrg >= batas_pagi:
                 # WINDOW PAGI SUDAH LEWAT (>= 11:00)
                 if not is_already_done(user['name'], 'pagi'):
                     print(f"⏰ [{user['name']}] Window Pagi sudah lewat ({skrg.strftime('%H:%M')} > 11:00).")
                     print(f"   📝 Auto-mark Pagi sebagai 'Done (Skipped)' agar history bersih.")
                     save_history_entry(user['name'], 'pagi')

            # --- CEK SORE ---
            h_s, m_s = map(int, sore_t_str.split(':'))
            target_sore = skrg.replace(hour=h_s, minute=m_s, second=0, microsecond=0)
            batas_sore = skrg.replace(hour=20, minute=0, second=0, microsecond=0)
            
            if target_sore < skrg < batas_sore:
                 # MASIH DALAM WINDOW SORE (sebelum 20:00)
                 if is_already_done(user['name'], 'sore'):
                     print(f"🍵 [{user['name']}] Absen Sore SUDAH SELESAI hari ini. (Skip Catch-Up)")
                 else:
                     print(f"⚠️ DETEKSI TELAT STARTUP ({user['name']})!!")
                     print(f"   Jadwal: {sore_t_str}, Sekarang: {skrg.strftime('%H:%M')}")
                     print("   🚀 RUNNING CATCH-UP SORE (Async)...")
                     threading.Thread(target=absen_sore, args=(user,)).start()
            elif skrg >= batas_sore:
                 # WINDOW SORE SUDAH LEWAT (>= 20:00)
                 if not is_already_done(user['name'], 'sore'):
                     print(f"⏰ [{user['name']}] Window Sore sudah lewat ({skrg.strftime('%H:%M')} > 20:00).")
                     print(f"   📝 Auto-mark Sore sebagai 'Done (Skipped)' agar history bersih.")
                     save_history_entry(user['name'], 'sore')
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 Script Dihentikan User.")
        time.sleep(2)

# Library Mode: No Main Function
# main() removed. Use Govem_Suami.py or Govem_Istri.py