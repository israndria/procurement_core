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

# Konfigurasi LDPlayer (Sama dengan V22)
LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")

PACKAGE_NAME = "go.id.tapinkab.govem"
CONFIG_FILE = "govem_aktivitas_config.ini" # Config khusus aktivitas

# Direktori Script Lama (Sumber Teks)
BASE_DIR_SUAMI = r"D:\Download\Aktivitas Govem\Aktivitas"
BASE_DIR_ISTRI = r"D:\Download\Aktivitas Govem\Aktivitas Istri" # Asumsi nama folder

# Global Config Storage
BUTTON_MAP = {} 

def run_command(command):
    try:
        startupinfo = None
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
    """
    Connect ke emulator via ADB.
    launch_if_needed: Jika False, tidak akan launch/bring to foreground emulator (BACKGROUND MODE)
                      Jika True, akan launch emulator jika belum running (foreground)
    """
    print(f"🔗 Menghubungkan Emulator {idx} (Background Mode: {not launch_if_needed})...")
    
    # BACKGROUND MODE: Skip ldconsole launch, langsung connect ADB
    if launch_if_needed:
        print("   📱 Launching emulator (foreground)...")
        run_command(f'"{LDCONSOLE}" launch --index {idx}')
    else:
        print("   🔇 Skip launch (background mode - assuming emulator already running)")
    
    # Auto-Discovery Port logic (Simplified from V22)
    possible_ports = [5554 + (idx*2), 5556, 5558, 5560, 5562, 5564]
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
             if f"emulator-{p}" in devices and "device" in devices:
                 detected_serial = f"emulator-{p}"
                 break
        if detected_serial: break
        time.sleep(1) # Dipercepat dari 2 detik
        
    if not detected_serial:
        # Fallback default
        detected_serial = f"127.0.0.1:{5554 + (idx*2)}"
        
    print(f"   ✅ Terhubung ke: {detected_serial}")
    return detected_serial

def adb_click(serial, x, y):
    run_command([ADB, "-s", serial, "shell", "input", "tap", str(x), str(y)])

def adb_input_text(serial, text, idx=0):
    """
    Input teks via ADB - sudah proven bekerja di background
    """
    # ADB input text (proven bekerja di background)
    escaped_text = text.replace(" ", "%s")
    escaped_text = escaped_text.replace("'", "")
    escaped_text = escaped_text.replace('"', "")
    escaped_text = escaped_text.replace("&", "")
    escaped_text = escaped_text.replace("(", "")
    escaped_text = escaped_text.replace(")", "")
    escaped_text = escaped_text.replace(";", "")
    
    run_command([ADB, "-s", serial, "shell", "input", "text", escaped_text])

# --- TEXT PARSER (Extract content from old Python scripts) ---
def parse_text_from_py(filepath):
    texts = []
    try:
        if not os.path.exists(filepath):
            print(f"❌ File tidak ditemukan: {filepath}")
            return []
            
        with open(filepath, 'r') as f:
            content = f.read()
            # Regex to find pyautogui.write("TEXT")
            matches = re.findall(r'pyautogui\.write\((?:"|\')(.+?)(?:"|\')\)', content)
            texts = matches
    except Exception as e:
        print(f"Error parsing {filepath}: {e}")
    return texts

def get_suami_activities(day_idx):
    """
    Ambil daftar aktivitas Suami berdasarkan hari.
    day_idx: 0=Senin, 1=Selasa, ..., 4=Jumat
    
    Senin-Kamis: 7 aktivitas (apel pagi, 5 telaah, apel sore)
    Jumat: 6 aktivitas (senam pagi, 5 telaah) - TIDAK ADA APEL SORE
    """
    # 1. Tentukan Source File
    source_file = ""
    is_jumat = (day_idx == 4)
    
    if is_jumat:
        source_file = os.path.join(BASE_DIR_SUAMI, "Aktivitas Govem (Jumat1).py")
    else: # Senin-Kamis
        source_file = os.path.join(BASE_DIR_SUAMI, "Aktivitas Govem (Rutinitas).py")
        
    print(f"📖 Membaca sumber teks: {os.path.basename(source_file)}")
    
    # 2. Ambil Teks
    activities = parse_text_from_py(source_file)
    
    # 3. Handle Jumat vs Senin-Kamis
    if is_jumat:
        # Jumat: 6 aktivitas saja, TIDAK PERLU tambah apel sore
        print(f"   📅 JUMAT: {len(activities)} aktivitas (senam pagi + telaah)")
    else:
        # Senin-Kamis: Cek apakah perlu tambah apel sore
        has_apel_sore = any("apel sore" in act.lower() for act in activities)
        if len(activities) > 0 and not has_apel_sore:
            activities.append("Melaksanakan apel sore")
            print(f"   ➕ Menambahkan 'Melaksanakan apel sore' sebagai aktivitas terakhir.")
        else:
            print(f"   ✅ {len(activities)} aktivitas (termasuk apel sore)")
    
    return activities, is_jumat  # Return juga flag is_jumat untuk Step 3 logic

# --- HELPER: DIRECT JSON REPLAY (BYPASS LDCONSOLE) ---
# Global cached serial untuk reuse antar macro calls
_CACHED_SERIAL = {}

def parse_json_and_replay(idx, record_filename_no_ext, serial=None):
    """
    Replay macro dari .record file via ADB input.
    serial: Jika None, akan connect otomatis. Jika disupply, akan reuse (lebih cepat).
    """
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
        
        # Use provided serial or get cached/new connection
        if serial:
            current_serial = serial
        elif idx in _CACHED_SERIAL:
            current_serial = _CACHED_SERIAL[idx]
        else:
            current_serial = connect_adb_smart(idx, launch_if_needed=False)
            _CACHED_SERIAL[idx] = current_serial
        
        # HARDCODE CALIBRATION (V22)
        w_res = 1600
        h_res = 900
        max_in_x = 19092 
        max_in_y = 10728

        last_timing = 0
        
        for op in ops:
            timing = op.get('timing', 0)
            op_id = op.get('operationId')
            
            # Hitung delay (dipercepat 50%)
            delay = (timing - last_timing) / 1000.0 * 0.5
            if delay > 0:
                time.sleep(delay)
            last_timing = timing
            
            if op_id == 'PutMultiTouch':
                points = op.get('points', [])
                if points and points[0].get('state') == 1:
                    raw_x = points[0].get('x')
                    raw_y = points[0].get('y')
                    
                    # Konversi Raw ke Pixel
                    real_x = int(raw_x / max_in_x * w_res)
                    real_y = int(raw_y / max_in_y * h_res)
                    
                    # Eksekusi ADB Click (SWIPE pendek)
                    run_command([ADB, "-s", current_serial, "shell", "input", "swipe", str(real_x), str(real_y), str(real_x), str(real_y), "100"])
                    
            elif op_id == 'Wait':
                 pass

        print("   ✅ Selesai")
        
    except Exception as e:
        print(f"   ❌ Gagal: {e}")

# Alias wrapper agar kompatibel dengan call code sebelumnya
def play_record_file(idx, record_filename_no_ext, serial=None):
    parse_json_and_replay(idx, record_filename_no_ext, serial)

# --- HYBRID RUNNER ---
def run_hybrid_automation(idx, background_mode=True):
    """
    Main automation runner.
    background_mode: True = tidak akan bring LDPlayer ke foreground (tetap minimize)
    """
    print(f"\n{'🔇' if background_mode else '📱'} Mode: {'BACKGROUND' if background_mode else 'FOREGROUND'}")
    
    # Connect sekali di awal, lalu reuse
    serial = connect_adb_smart(idx, launch_if_needed=(not background_mode))
    
    # Cache serial untuk reuse
    global _CACHED_SERIAL
    _CACHED_SERIAL[idx] = serial
    
    # 1. Identify User & Day
    is_suami = (idx == 0)
    wd = datetime.datetime.today().weekday()
    
    # 2. Load Activity List
    activity_texts = []
    is_jumat = False
    
    if is_suami:
        result = get_suami_activities(wd)
        activity_texts = result[0]
        is_jumat = result[1]
    else:
        # Logic Istri (Placeholder / Future)
        # Nanti disesuaikan mapping hari -> file istri
        print("ℹ Logic Istri belum disetujui detailnya. Skip.")
        return

    if not activity_texts:
        print("❌ Tidak ada daftar aktivitas yang ditemukan.")
        return

    print(f"\n🚀 MEMULAI PENGISIAN {len(activity_texts)} AKTIVITAS")
    if is_jumat:
        print("📅 Mode JUMAT: Step 3.2 akan digunakan (SKP aktivitas 6)")
    if background_mode:
        print("💡 LDPlayer bisa di-minimize, script tetap berjalan!")
    
    # 3. STEP 1: NAVIGASI (Hanya sekali di awal)
    print("\n[STEP 1] Navigasi ke Halaman Form...")
    play_record_file(idx, "Step 1 (Menuju Dashboard Isian)", serial)
    time.sleep(3)
    
    # 4. LOOP ITEMS
    for i, text in enumerate(activity_texts):
        print(f"\n📝 [{i+1}/{len(activity_texts)}] Mengisi: {text[:30]}...")
        
        # STEP 2: Focus Input
        play_record_file(idx, "Step 2 (Klik untuk supaya bisa mengetikcopas kata2)", serial)
        time.sleep(0.8)
        
        # INPUT TEKS (ADB - proven bekerja di background)
        print(f"   ⌨️ Mengetik...")
        clean_text = text.replace("'", "").replace('"', "")
        adb_input_text(serial, clean_text, idx)
        time.sleep(0.3)
        
        # Hide Keyboard
        run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(0.3)

        # --- STEP 3: DROPDOWN SKP (KONDISIONAL BERDASARKAN TEKS) ---
        # Step 3.2 HANYA untuk senam pagi (Jumat aktivitas 1) - perlu delay lebih lama karena ada scroll
        # Step 3.1 untuk semua aktivitas lainnya (termasuk telaah di Jumat)
        text_lower = text.lower()
        is_senam = "senam pagi" in text_lower
        is_apel = ("apel pagi" in text_lower or "apel sore" in text_lower)
        
        print("   📋 Memilih SKP...")
        if is_senam:
            # Senam Pagi (Jumat): Step 3.2 hanya buka dropdown, lalu scroll + klik aktivitas 6 via ADB
            play_record_file(idx, "Step 3.2 (Membuka dropdown SKP dan memilih aktivitas nomor 6)", serial)
            time.sleep(1)  # Tunggu dropdown terbuka
            
            # Scroll down dalam dropdown untuk melihat aktivitas 6
            # Koordinat scroll: dari tengah dropdown ke atas (swipe up = scroll down)
            # Posisi dropdown kira-kira di tengah layar, swipe dari Y bawah ke Y atas
            print("   📜 Scroll dropdown...")
            run_command([ADB, "-s", serial, "shell", "input", "swipe", "400", "700", "400", "400", "300"])
            time.sleep(0.8)
            
            # Klik aktivitas nomor 6 - koordinat sudah diverifikasi tepat
            print("   👆 Klik aktivitas 6...")
            run_command([ADB, "-s", serial, "shell", "input", "tap", "400", "790"])
            time.sleep(1)
        else:
            # Semua aktivitas lain: Step 3.1 (SKP aktivitas 4) - sudah termasuk buka dropdown + klik
            play_record_file(idx, "Step 3.1 (Membuka dropdown SKP dan memilih aktivitas nomor 4)", serial)
            time.sleep(1.5)
        
        # --- STEP 4: JENIS AKTIVITAS ---
        # Apel pagi/sore = Step 4.1 (Jenis 2)
        # Senam pagi dan aktivitas umum = Step 4 (Jenis 1)
        if is_apel:
            print("   ⭐ Mode Apel -> Step 4.1 (Jenis 2)")
            play_record_file(idx, "Step 4.1 (Memilih Jenis Aktivitas Nomor 2)", serial)
        else:
            print("   📄 Mode Umum/Senam -> Step 4 (Jenis 1)")
            play_record_file(idx, "Step 4 (Memilih Jenis Aktivitas Nomor 1)", serial)
        time.sleep(0.8)
        
        # STEP 5: Simpan
        play_record_file(idx, "Step 5 (Posting Aktivitas)", serial)
        print("   💾 Simpan...")
        time.sleep(2)
        
    print("\n✅ SEMUA AKTIVITAS SELESAI!")

def main():
    print("🤖 GOVEM HYBRID AUTOMATION (V23)")
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
