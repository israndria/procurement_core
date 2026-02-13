"""
V23_aktivitas_Istri.py - Pengisian Aktivitas Otomatis untuk Istri
=================================================================
- Background Mode (LDPlayer bisa minimize)
- ADB input text (proven bekerja)
- Aktivitas per hari berbeda (Senin-Sabtu)
- Reuse recording Suami + mapping koordinat
"""

import subprocess
import time
import datetime
import os
import sys
import json
import logging

# --- PATHS ---
LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")
RECORDS_DIR = os.path.join(LDPLAYER_PATH, r"vms\operationRecords")
PACKAGE_NAME = "go.id.tapinkab.govem"

# Index untuk Istri
ISTRI_IDX = 1

# Global cached serial
_CACHED_SERIAL = {}
 
# Setup Logging (agar tercatat di file log utama)
LOG_FILE = os.path.join(os.path.dirname(__file__), "govem_scheduler.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] [ISTRI] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("V23_Istri")

def log_info(msg):
    print(msg)
    logger.info(msg)

def log_error(msg):
    print(msg)
    logger.error(msg)
 
# Setup Logging (agar tercatat di file log utama)
LOG_FILE = os.path.join(os.path.dirname(__file__), "govem_scheduler.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] [ISTRI] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("V23_Istri")

def log_info(msg):
    print(msg)
    logger.info(msg)

def log_error(msg):
    print(msg)
    logger.error(msg)

# --- AKTIVITAS ISTRI PER HARI ---
# Format: {"teks": "...", "step3": "recording_name", "step4": "recording_name"}
# step3 = Recording untuk SKP (Istri 3, Istri 3.1, Istri 3.2, Istri 3.3)
# step4 = Recording untuk Jenis (Istri 4, Istri 4.1)

AKTIVITAS_ISTRI = {
    0: [  # SENIN
        {"teks": "membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan", "step3": "Istri 3", "step4": "Istri 4"},
        {"teks": "Menyusun kegiatan pembelajaran", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran bahasa Inggris", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran bahasa Indonesia", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Menyusun kegiatan evaluasi", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
    ],
    1: [  # SELASA
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
        {"teks": "Menyusun kegiatan pembelajaran", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran ilmu pengetahuan alam dan sosial", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran pendidikan Pancasila", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melakukan kegiatan evaluasi", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
    ],
    2: [  # RABU
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan", "step3": "Istri 3", "step4": "Istri 4"},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan", "step3": "Istri 3", "step4": "Istri 4"},
    ],
    3: [  # KAMIS
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
        {"teks": "Menyusun kegiatan pembelajaran", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran matematika", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Menyusun kegiatan evaluasi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Mengolah nilai peserta didik", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
    ],
    4: [  # JUMAT (6 aktivitas)
        {"teks": "Berpartisipasi dalam kegiatan Jumat Takwa", "step3": "Istri 3", "step4": "Istri 4.1"},
        {"teks": "Menyusun kegiatan pembelajaran", "step3": "Istri 3.1", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran muatan lokal baca tulis Al-Quran", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan pembelajaran kesenian", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan refleksi pembelajaran", "step3": "Istri 3.3", "step4": "Istri 4.1"},
    ],
    5: [  # SABTU (7 aktivitas)
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "step3": "Istri 3", "step4": "Istri 4"},
        {"teks": "Mempersiapkan kegiatan kokurikuler", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan kokurikuler", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan kokurikuler", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan kokurikuler", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Melaksanakan kegiatan kokurikuler", "step3": "Istri 3.2", "step4": "Istri 4.1"},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan", "step3": "Istri 3", "step4": "Istri 4"},
    ],
}

# --- KOORDINAT SKP (berdasarkan analisa pyautogui + mapping ke LDPlayer) ---
# Koordinat Y untuk dropdown SKP di LDPlayer (resolusi internal)
# SKP 1-4 visible, SKP 5-6 perlu scroll
SKP_COORDS = {
    1: (400, 550),   # SKP opsi 1
    2: (400, 620),   # SKP opsi 2 - Menyusun
    3: (400, 690),   # SKP opsi 3
    4: (400, 760),   # SKP opsi 4 - Melaksanakan pembelajaran
    5: (400, 820),   # SKP opsi 5 - Evaluasi (mungkin perlu scroll)
    6: (400, 790),   # SKP opsi 6 - Upacara/Membimbing (PASTI perlu scroll)
}

# Jenis aktivitas koordinat
JENIS_COORDS = {
    1: (400, 613),   # Jenis 1 - Umum
    2: (400, 680),   # Jenis 2 - Upacara
}

# --- HELPER FUNCTIONS ---
def run_command(command):
    creationflags = 0x08000000  # CREATE_NO_WINDOW
    if isinstance(command, list):
        result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
    else:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
    return result.stdout.strip()

# --- LDCONSOLE HELPER FUNCTIONS (INDEX BASED) ---
def ldconsole_adb(idx, command):
    """Kirim perintah ADB via LDConsole (Target Index Pasti Benar)"""
    # Format: ldconsole.exe adb --index 1 --command "shell input tap ..."
    # command arg harus di-quote? subprocess handle arg splitting.
    args = [LDCONSOLE, "adb", "--index", str(idx), "--command", command]
    return run_command(args)

def adb_click(idx, x, y):
    ldconsole_adb(idx, f"shell input tap {x} {y}")

def adb_swipe(idx, x1, y1, x2, y2, duration=300):
    ldconsole_adb(idx, f"shell input swipe {x1} {y1} {x2} {y2} {duration}")

def adb_input_text(idx, text):
    """Input teks via LDCONSOLE ADB - Index Based."""
    # Escape characters
    escaped_text = text.replace(" ", "%s")
    for char in ["'", '"', "&", "(", ")", ";", "<", ">", "|", "*", "\\"]:
        escaped_text = escaped_text.replace(char, "")
    
    ldconsole_adb(idx, f"shell input text {escaped_text}")

def play_record_file(idx, record_name):
    """Replay macro dari file recording via LDCONSOLE ADB Wrapper."""
    # NOTE: Serial argument removed, we use idx consistently
    serial = None # Backward compatibility variable unused
    
    filepath = os.path.join(RECORDS_DIR, record_name + ".record")
    
    if not os.path.exists(filepath):
        print(f"   ❌ Recording tidak ditemukan: {record_name}")
        return False
    
    print(f"   ▶️ Macro: {record_name[:40]}...")
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        operations = data.get("operations", [])
        
        # HARDCODE CALIBRATION (SAMA DENGAN V22/SUAMI)
        w_res = 1600
        h_res = 900
        max_in_x = 19092
        max_in_y = 10728
        
        prev_timing = 0
        
        for op in operations:
            op_id = op.get("operationId")
            timing = op.get("timing", 0)
            
            # Delay dipercepat 50%
            delay = (timing - prev_timing) / 1000.0 * 0.5
            if delay > 0:
                time.sleep(delay)
            prev_timing = timing
            
            if op_id == "PutMultiTouch":
                points = op.get("points", [])
                if points and points[0].get("state") == 1:
                    raw_x = points[0].get("x", 0)
                    raw_y = points[0].get("y", 0)
                    
                    # Konversi Raw ke Pixel (SAMA DENGAN SUAMI)
                    real_x = int(raw_x / max_in_x * w_res)
                    real_y = int(raw_y / max_in_y * h_res)
                    
                    # Eksekusi ADB Click via LDConsole Index
                    adb_click(idx, real_x, real_y)
        
        print("   ✅ Selesai")
        return True
    except Exception as e:
        print(f"   ❌ Error replay: {e}")
        return False

def click_skp_option(idx, skp_num):
    """Klik opsi SKP berdasarkan nomor."""
    # Buka dropdown SKP dulu
    print(f"   📋 Membuka dropdown SKP...")
    adb_click(idx, 400, 540)  # Koordinat dropdown SKP
    time.sleep(0.8)
    
    # Jika SKP 5 atau 6, perlu scroll
    if skp_num >= 5:
        print(f"   📜 Scroll untuk SKP {skp_num}...")
        adb_swipe(idx, 400, 700, 400, 400, 300)
        time.sleep(0.5)
    
    # Klik opsi SKP
    coords = SKP_COORDS.get(skp_num, (400, 620))
    print(f"   📋 Memilih SKP {skp_num} di ({coords[0]}, {coords[1]})")
    adb_click(idx, coords[0], coords[1])
    time.sleep(0.8)

def click_jenis_option(idx, jenis_num):
    """Klik opsi Jenis berdasarkan nomor."""
    # Buka dropdown Jenis
    print(f"   🎯 Membuka dropdown Jenis...")
    adb_click(idx, 400, 640)  # Koordinat dropdown Jenis
    time.sleep(0.8)
    
    # Klik opsi
    coords = JENIS_COORDS.get(jenis_num, (400, 613))
    print(f"   🎯 Memilih Jenis {jenis_num} di ({coords[0]}, {coords[1]})")
    adb_click(idx, coords[0], coords[1])
    time.sleep(0.5)

# --- MAIN AUTOMATION ---
def run_istri_automation(background_mode=True, override_hari=None):
    """Jalankan automation pengisian aktivitas untuk Istri."""
    idx = ISTRI_IDX
    
    print("\n" + "="*60)
    print("   🤖 V23 AKTIVITAS ISTRI (Background Mode)")
    print("="*60)
    
    # 1. Connect (Skip - We use LDCONSOLE Index)
    log_info(f"   🚀 Menggunakan LDConsole Index {idx}")
    
    # 2. Tentukan hari
    if override_hari is not None:
        hari_ini = override_hari
    else:
        hari_ini = datetime.datetime.now().weekday()
    nama_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"][hari_ini]
    
    # 2.5. RESTART APP (Force Clean State)
    log_info("🔄 Restarting App agar mulai dari Dashboard bersih...")
    run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(2)
    run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
    
    log_info("⏳ Menunggu aplikasi siap (25s)...")
    time.sleep(25)
    
    # 3. Navigasi ke form (Step 1)
    
    # Minggu tidak ada aktivitas
    if hari_ini == 6:
        print(f"\n📅 Hari ini {nama_hari} - tidak ada aktivitas")
        return True
    
    # 3. Ambil aktivitas hari ini
    aktivitas_list = AKTIVITAS_ISTRI.get(hari_ini, [])
    if not aktivitas_list:
        print(f"❌ Tidak ada data aktivitas untuk {nama_hari}")
        return False
    
    print(f"\n📅 Hari: {nama_hari}")
    print(f"📋 Jumlah aktivitas: {len(aktivitas_list)}")
    
    # 4. Navigasi ke form (Step 1)
    print("\n📱 Step 1: Navigasi ke form aktivitas...")
    play_record_file(idx, "Step 1 (Menuju Dashboard Isian)")
    time.sleep(3)
    
    # 5. Loop aktivitas
    for i, akt in enumerate(aktivitas_list):
        teks = akt["teks"]
        step3_rec = akt["step3"]  # Recording untuk SKP
        step4_rec = akt["step4"]  # Recording untuk Jenis
        
        log_info(f"\n📝 [{i+1}/{len(aktivitas_list)}] {teks[:40]}...")
        # print(f"   Recording: {step3_rec} + {step4_rec}")
        
        # Step 2: Klik field input
        print("   ⌨️ Step 2: Klik field input...")
        play_record_file(idx, "Step 2 (Klik untuk supaya bisa mengetikcopas kata2)")
        time.sleep(0.8)
        
        # Input teks
        print("   ⌨️ Mengetik teks...")
        clean_text = teks.replace("'", "").replace('"', "")
        adb_input_text(idx, clean_text)
        time.sleep(0.3)
        
        # Hide keyboard
        ldconsole_adb(idx, "shell input keyevent 111")
        time.sleep(0.3)
        
        # Step 3: Pilih SKP (pakai recording Istri)
        print(f"   📋 Step 3: {step3_rec}...")
        play_record_file(idx, step3_rec)
        time.sleep(1)
        
        # Step 4: Pilih Jenis (pakai recording Istri)
        print(f"   🎯 Step 4: {step4_rec}...")
        play_record_file(idx, step4_rec)
        time.sleep(0.8)
        
        # Step 5: Simpan
        print("   💾 Step 5: Menyimpan...")
        play_record_file(idx, "Step 5 (Posting Aktivitas)")
        time.sleep(2)
    
    print("\n" + "="*60)
    log_info(f"   ✅ SELESAI! {len(aktivitas_list)} aktivitas terisi")
    print("="*60)
    
    return True

# --- ENTRY POINT ---
def main():
    print("\n" + "="*50)
    print("   🤖 V23_aktivitas_Istri")
    print("="*50)
    print()
    print("Mode:")
    print("1. Jalankan otomatis (hari ini)")
    print("2. Test hari tertentu")
    print("0. Keluar")
    print()
    
    choice = input("Pilihan: ").strip()
    
    if choice == '1':
        run_istri_automation()
    elif choice == '2':
        print("\nPilih hari:")
        print("0=Senin, 1=Selasa, 2=Rabu, 3=Kamis, 4=Jumat, 5=Sabtu")
        hari = int(input("Hari: ").strip())
        
        # Jalankan dengan override hari
        run_istri_automation(override_hari=hari)
    elif choice == '0':
        return
    
    input("\nTekan ENTER untuk keluar...")

if __name__ == "__main__":
    main()
