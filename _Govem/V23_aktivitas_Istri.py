"""
V23_aktivitas_Istri.py - Pengisian Aktivitas Otomatis untuk Istri
=================================================================
v1.2 — Rewritten to mirror Suami's proven structure
- Uses ADB -s serial (NOT ldconsole_adb) — same as Suami
- Same coordinates as Suami (same Govem app UI)
- Same timing, same flow, same helpers
- Aktivitas per hari berbeda (Senin-Sabtu)
"""

import os
import time
import subprocess
import datetime
import sys
import logging

# --- PATHS ---
LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")
PACKAGE_NAME = "go.id.tapinkab.govem"

# --- DRY RUN FLAG ---
DRY_RUN = "--dry-run" in sys.argv

# Index untuk Istri
ISTRI_IDX = 1

# Setup Logging
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
    logger.info(msg)

def log_error(msg):
    logger.error(msg)

# =============================================
# KOORDINAT UI GOVEM ISTRI (dari uiautomator dump)
# Berbeda dari Suami! Modal popup punya offset berbeda.
# =============================================
# Dropdown SKP
SKP_DROPDOWN_XY = (500, 318)
SKP_COORDS = {
    # bounds dari uiautomator: pakai center Y
    1: (800, 518),   # [0,469]-[1600,568] Melaksanakan tugas yang diberikan atasan
    2: (800, 617),   # [0,568]-[1600,667] Merencanakan pembelajaran
    3: (800, 716),   # [0,667]-[1600,766] Melaksanakan pembelajaran
    4: (800, 815),   # [0,766]-[1600,865] Menilai hasil pembelajaran
}

# Dropdown Jenis
JENIS_DROPDOWN_XY = (1130, 318)
JENIS_COORDS = {
    1: (800, 508),   # [0,475]-[1600,542] Aktifitas
    2: (800, 592),   # [0,559]-[1600,626] Apel/Shif/Piket/Lainnya
}

# Navigasi & Form
STEP1_TAP_1 = (1153, 769)  # Klik menu aktivitas
STEP1_TAP_2 = (500, 331)   # Klik buat aktivitas harian
STEP2_INPUT_XY = (246, 456) # Klik field input teks
STEP5_SAVE_XY = (795, 897)  # Tombol simpan/posting

# =============================================
# AKTIVITAS ISTRI PER HARI
# Format: {"teks": "...", "skp": N, "jenis": N}
# skp: nomor opsi di dropdown SKP (1-4)
# jenis: 1=Aktifitas, 2=Apel/Shif/Piket
# =============================================
AKTIVITAS_ISTRI = {
    0: [  # SENIN (7 aktivitas)
        {"teks": "membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan", "skp": 1, "jenis": 2},
        {"teks": "Menyusun kegiatan pembelajaran", "skp": 2, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran bahasa Inggris", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran bahasa Indonesia", "skp": 3, "jenis": 1},
        {"teks": "Menyusun kegiatan evaluasi", "skp": 2, "jenis": 1},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "skp": 4, "jenis": 1},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "skp": 1, "jenis": 2},
    ],
    1: [  # SELASA (7 aktivitas)
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "skp": 1, "jenis": 2},
        {"teks": "Menyusun kegiatan pembelajaran", "skp": 2, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran ilmu pengetahuan alam dan sosial", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "skp": 4, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran pendidikan Pancasila", "skp": 3, "jenis": 1},
        {"teks": "Melakukan kegiatan evaluasi", "skp": 4, "jenis": 1},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "skp": 1, "jenis": 2},
    ],
    2: [  # RABU (7 aktivitas)
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan", "skp": 1, "jenis": 2},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "skp": 2, "jenis": 1},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "skp": 2, "jenis": 1},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "skp": 2, "jenis": 1},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "skp": 2, "jenis": 1},
        {"teks": "Menambah wawasan yang relevan dengan tugas sebagai guru", "skp": 2, "jenis": 1},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan", "skp": 1, "jenis": 2},
    ],
    3: [  # KAMIS (7 aktivitas)
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "skp": 1, "jenis": 2},
        {"teks": "Menyusun kegiatan pembelajaran", "skp": 2, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran matematika", "skp": 3, "jenis": 1},
        {"teks": "Menyusun kegiatan evaluasi pembelajaran", "skp": 4, "jenis": 1},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "skp": 4, "jenis": 1},
        {"teks": "Mengolah nilai peserta didik", "skp": 4, "jenis": 1},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan pembelajaran", "skp": 1, "jenis": 2},
    ],
    4: [  # JUMAT (6 aktivitas)
        {"teks": "Berpartisipasi dalam kegiatan Jumat Takwa", "skp": 1, "jenis": 1},
        {"teks": "Menyusun kegiatan pembelajaran", "skp": 2, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran muatan lokal baca tulis Al-Quran", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan pembelajaran kesenian", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan evaluasi pembelajaran", "skp": 4, "jenis": 1},
        {"teks": "Melaksanakan kegiatan refleksi pembelajaran", "skp": 4, "jenis": 1},
    ],
    5: [  # SABTU (7 aktivitas)
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk memulai kegiatan pembelajaran", "skp": 1, "jenis": 2},
        {"teks": "Mempersiapkan kegiatan kokurikuler", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan kokurikuler", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan kokurikuler", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan kokurikuler", "skp": 3, "jenis": 1},
        {"teks": "Melaksanakan kegiatan kokurikuler", "skp": 3, "jenis": 1},
        {"teks": "Membimbing peserta didik baris berbaris dan berdoa untuk mengakhiri kegiatan", "skp": 1, "jenis": 2},
    ],
}

# =============================================
# HELPER FUNCTIONS (carbon copy dari Suami)
# =============================================
def run_command(command):
    if DRY_RUN:
        return ""
    try:
        creationflags = 0x08000000 if os.name == 'nt' else 0
        if isinstance(command, list):
            result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
        else:
            result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
        return result.stdout.strip()
    except Exception as e:
        log_error(f"Error executing command: {e}")
        return ""

_CACHED_SERIAL = {}

def connect_adb_smart(idx, launch_if_needed=False):
    if DRY_RUN:
        log_info(f"[DRY RUN] Skip ADB connect (emulator {idx})")
        return "dry-run-serial"

    log_info(f"Menghubungkan Emulator {idx}...")

    if launch_if_needed:
        log_info("Launching emulator...")
        run_command(f'"{LDCONSOLE}" launch --index {idx}')
        # Loop minimize sampai emulator running
        for attempt in range(30):
            time.sleep(1)
            run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
            s = run_command(f'"{LDCONSOLE}" list2')
            for line in s.splitlines():
                parts = line.split(",")
                if len(parts) > 4 and parts[0] == str(idx) and parts[4] == "1":
                    log_info(f"Emulator running setelah {attempt+1}s, minimize sent!")
                    break
            else:
                continue
            break
        # Extra minimize untuk memastikan
        for _ in range(3):
            time.sleep(1)
            run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')

    possible_ports = [5554 + (idx*2), 5556, 5558, 5560, 5562, 5564]
    detected_serial = None

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
        if detected_serial:
            break
        time.sleep(1)

    if not detected_serial:
        detected_serial = f"127.0.0.1:{5554 + (idx*2)}"

    log_info(f"Terhubung ke: {detected_serial}")
    return detected_serial

def adb_click(serial, x, y):
    run_command([ADB, "-s", serial, "shell", "input", "tap", str(x), str(y)])

def adb_input_text(serial, text, idx=0):
    if DRY_RUN:
        return
    escaped_text = text.replace(" ", "%s")
    for char in ["'", '"', "&", "(", ")", ";", "<", ">", "|", "*", "\\"]:
        escaped_text = escaped_text.replace(char, "")
    run_command([ADB, "-s", serial, "shell", "input", "text", escaped_text])

def take_screenshot(serial, idx):
    if DRY_RUN:
        log_info("[DRY RUN] Skip screenshot")
        return None
    try:
        import tempfile
        screenshot_path = os.path.join(tempfile.gettempdir(), f"govem_aktivitas_istri_{idx}.png")
        result = subprocess.run(
            [ADB, "-s", serial, "exec-out", "screencap", "-p"],
            capture_output=True, timeout=10,
            creationflags=0x08000000 if os.name == 'nt' else 0
        )
        if result.returncode == 0 and result.stdout:
            with open(screenshot_path, 'wb') as f:
                f.write(result.stdout)
            log_info(f"Screenshot disimpan: {screenshot_path}")
            return screenshot_path
    except Exception as e:
        log_error(f"Gagal screenshot: {e}")
    return None

def show_notification(message, screenshot_path=None):
    try:
        from ctypes import windll
        windll.user32.MessageBeep(0x00000040)
    except:
        pass
    try:
        ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$textNodes = $template.GetElementsByTagName("text")
$textNodes.Item(0).AppendChild($template.CreateTextNode("Govem Bot (Istri)")) | Out-Null
$textNodes.Item(1).AppendChild($template.CreateTextNode("{message}")) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($template)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Govem Bot").Show($toast)
'''
        subprocess.run(["powershell", "-Command", ps_script],
                       capture_output=True, timeout=10,
                       creationflags=0x08000000 if os.name == 'nt' else 0)
        log_info(f"Notifikasi terkirim: {message}")
    except Exception as e:
        log_error(f"Notifikasi gagal: {e}")

    if screenshot_path and os.path.exists(screenshot_path):
        try:
            os.startfile(screenshot_path)
        except:
            pass

def minimize_emulator(idx):
    try:
        run_command(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
    except:
        pass

# =============================================
# MAIN AUTOMATION (mirror Suami's run_hybrid_automation)
# =============================================
def run_istri_automation(background_mode=True, override_hari=None):
    idx = ISTRI_IDX

    if DRY_RUN:
        log_info("=== DRY RUN MODE === (tidak ada ADB yang dieksekusi)")

    # AUTO-MINIMIZE INSTANT (sebelum apapun)
    minimize_emulator(idx)

    # Connect ADB (same as Suami)
    serial = connect_adb_smart(idx, launch_if_needed=(not background_mode))
    _CACHED_SERIAL[idx] = serial

    # Tentukan hari
    if override_hari is not None:
        hari_ini = override_hari
    else:
        hari_ini = datetime.datetime.now().weekday()
    nama_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"][hari_ini]

    if hari_ini == 6:
        log_info(f"Hari ini {nama_hari} - tidak ada aktivitas")
        return True

    # Ambil aktivitas hari ini
    aktivitas_list = AKTIVITAS_ISTRI.get(hari_ini, [])
    if not aktivitas_list:
        log_error(f"Tidak ada data aktivitas untuk {nama_hari}")
        return False

    log_info(f"Hari: {nama_hari}, Jumlah aktivitas: {len(aktivitas_list)}")
    for i, act in enumerate(aktivitas_list):
        log_info(f"  {i+1}. [SKP{act['skp']} J{act['jenis']}] {act['teks'][:70]}")

    # RESTART APP (Force Clean State)
    log_info("Restarting App...")
    run_command(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE_NAME}')
    time.sleep(2)
    run_command(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE_NAME}')
    log_info("Menunggu aplikasi siap (25s)...")
    time.sleep(25)

    if DRY_RUN:
        log_info("[DRY RUN] Preview selesai.")
        return True

    log_info(f"MEMULAI PENGISIAN {len(aktivitas_list)} AKTIVITAS")

    # LOOP ITEMS
    for i, act in enumerate(aktivitas_list):
        text = act["teks"]
        skp_num = act["skp"]
        jenis_num = act["jenis"]

        log_info(f"[{i+1}/{len(aktivitas_list)}] Mengisi: {text[:50]}...")

        # STEP 1: NAVIGASI ke Form (HANYA iterasi pertama)
        # Setelah save, form auto-reset — langsung isi tanpa navigasi ulang
        if i == 0:
            log_info("  Navigasi Dashboard → Form...")
            adb_click(serial, STEP1_TAP_1[0], STEP1_TAP_1[1])
            time.sleep(2)
            adb_click(serial, STEP1_TAP_2[0], STEP1_TAP_2[1])
            time.sleep(3)
        else:
            log_info("  (Form auto-reset, langsung isi)")
            time.sleep(1)

        # STEP 2: Focus Input (direct ADB tap)
        adb_click(serial, STEP2_INPUT_XY[0], STEP2_INPUT_XY[1])
        time.sleep(0.8)

        # CLEAR FIELD: triple-tap to select all, then delete
        for _ in range(2):
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "123"])  # END
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "KEYCODE_MOVE_HOME"])
            run_command(f'"{ADB}" -s {serial} shell input keyevent --longpress 112 123')  # Shift+END (select all)
            time.sleep(0.1)
            run_command([ADB, "-s", serial, "shell", "input", "keyevent", "67"])  # DEL
            time.sleep(0.1)
        time.sleep(0.5)

        # INPUT TEKS
        log_info(f"  Mengetik...")
        clean_text = text.replace("'", "").replace('"', "")
        adb_input_text(serial, clean_text, idx)
        time.sleep(1)

        # Hide Keyboard
        run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(0.3)

        # STEP 3: Pilih SKP (same coords as Suami)
        log_info(f"  Memilih SKP {skp_num}...")
        adb_click(serial, SKP_DROPDOWN_XY[0], SKP_DROPDOWN_XY[1])
        time.sleep(0.8)
        if skp_num >= 5:
            # Perlu scroll (unlikely for Istri, but keep for safety)
            run_command([ADB, "-s", serial, "shell", "input", "swipe", "400", "700", "400", "400", "300"])
            time.sleep(0.8)
            adb_click(serial, 400, 790)
        else:
            coords = SKP_COORDS[skp_num]
            adb_click(serial, coords[0], coords[1])
        time.sleep(0.8)

        # STEP 4: Pilih Jenis (same coords as Suami)
        jenis_label = "Apel" if jenis_num == 2 else "Aktifitas"
        log_info(f"  Memilih Jenis {jenis_num} ({jenis_label})...")
        adb_click(serial, JENIS_DROPDOWN_XY[0], JENIS_DROPDOWN_XY[1])
        time.sleep(0.8)
        jenis_coords = JENIS_COORDS[jenis_num]
        adb_click(serial, jenis_coords[0], jenis_coords[1])
        time.sleep(0.8)

        # STEP 5: Simpan (same as Suami)
        log_info("  Simpan & Tunggu (12s)...")
        adb_click(serial, STEP5_SAVE_XY[0], STEP5_SAVE_XY[1])
        time.sleep(12)

    log_info(f"SELESAI! {len(aktivitas_list)} aktivitas terisi")
    return True

# --- ENTRY POINT ---
def main():
    print("V23_aktivitas_Istri v1.2 (Mirror Suami)")
    if DRY_RUN:
        print("[DRY RUN MODE]")

    print("1. Jalankan otomatis (hari ini)")
    print("2. Test hari tertentu")
    print("0. Keluar")

    choice = input("Pilihan: ").strip()
    if choice == '1':
        run_istri_automation(background_mode=False)
    elif choice == '2':
        print("0=Senin, 1=Selasa, 2=Rabu, 3=Kamis, 4=Jumat, 5=Sabtu")
        hari = int(input("Hari: ").strip())
        run_istri_automation(background_mode=False, override_hari=hari)

if __name__ == "__main__":
    main()
