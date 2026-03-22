"""
Debug Script v2: Form auto-reset setelah save
===============================================
Step 1 hanya untuk iterasi pertama.
Iterasi 2+ langsung Step 2-5 (form sudah bersih).
Screenshot di setiap titik kunci.

Jalankan: python -X utf8 _debug_step_by_step.py
"""
import os, sys, time, subprocess, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")
PACKAGE_NAME = "go.id.tapinkab.govem"
IDX = 0  # Suami

SS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots", "debug")
os.makedirs(SS_DIR, exist_ok=True)

# Koordinat Suami
STEP1_TAP_1 = (1153, 769)  # menu aktivitas (dari dashboard)
STEP1_TAP_2 = (500, 331)   # buat aktivitas harian
STEP2_INPUT_XY = (246, 456)
STEP5_SAVE_XY = (795, 897)
SKP_DROPDOWN_XY = (500, 318)
JENIS_DROPDOWN_XY = (1130, 318)

## KOORDINAT DARI UIAUTOMATOR DUMP (resolusi 1600x900)
SKP_COORDS = {
    1: (800, 518),   # bounds [0,469][1600,568]
    2: (800, 617),   # bounds [0,568][1600,667]
    3: (800, 716),   # bounds [0,667][1600,766]
    4: (800, 815),   # bounds [0,766][1600,865]
}
JENIS_COORDS = {
    1: (800, 508),   # Aktifitas — bounds [0,475][1600,542]
    2: (800, 592),   # Apel/Shif/Piket — bounds [0,559][1600,626]
}

# 3 aktivitas untuk test
AKTIVITAS = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Menelaah dan menganalisa dokumen peraturan", "skp": 4, "jenis": 1},
    {"teks": "Melaksanakan proses pemilihan penyedia", "skp": 1, "jenis": 1},
]

serial = None

def run_cmd(command):
    flags = 0x08000000 if os.name == 'nt' else 0
    if isinstance(command, list):
        result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=flags)
    else:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=flags)
    return result.stdout.strip()

def tap(x, y):
    """Gunakan swipe 150ms (sama seperti Govem_Engine) untuk reliabilitas."""
    run_cmd([ADB, "-s", serial, "shell", "input", "swipe", str(x), str(y), str(x), str(y), "150"])

def screenshot(label):
    ts = datetime.datetime.now().strftime("%H%M%S")
    path = os.path.join(SS_DIR, f"{ts}_{label}.png")
    result = subprocess.run(
        [ADB, "-s", serial, "exec-out", "screencap", "-p"],
        capture_output=True, timeout=10,
        creationflags=0x08000000 if os.name == 'nt' else 0
    )
    if result.returncode == 0 and result.stdout:
        with open(path, 'wb') as f:
            f.write(result.stdout)
        print(f"  📸 {label}: {os.path.basename(path)}")
    return path

def input_text(text):
    escaped = text.replace(" ", "%s").replace("'","").replace('"',"").replace("&","").replace("(","").replace(")","").replace(";","")
    run_cmd([ADB, "-s", serial, "shell", "input", "text", escaped])

def hide_keyboard():
    run_cmd([ADB, "-s", serial, "shell", "input", "keyevent", "111"])  # ESCAPE
    time.sleep(0.5)

# --- Connect ---
print("🔌 Connecting ADB...")
for p in [5554, 5556, 5558, 5560]:
    run_cmd(f'"{ADB}" connect 127.0.0.1:{p}')
devices = run_cmd(f'"{ADB}" devices')
for p in [5554, 5556, 5558, 5560]:
    if f"emulator-{p}" in devices:
        serial = f"emulator-{p}"
        break
if not serial:
    serial = "emulator-5554"
print(f"  Serial: {serial}")

# --- Restart app ---
print("\n🔄 Restarting app...")
run_cmd(f'"{LDCONSOLE}" killapp --index {IDX} --packagename {PACKAGE_NAME}')
time.sleep(2)
run_cmd(f'"{LDCONSOLE}" runapp --index {IDX} --packagename {PACKAGE_NAME}')
print("  Waiting 20s for app to load...")
time.sleep(20)
screenshot("00_dashboard")

# --- Step 1: Navigasi dari dashboard ke form (HANYA SEKALI) ---
print("\n📱 [Step 1] Dashboard → Form (hanya sekali)")
tap(*STEP1_TAP_1)
time.sleep(2)
screenshot("01_menu_aktivitas")
tap(*STEP1_TAP_2)
time.sleep(3)
screenshot("02_form_buat_aktivitas")

# --- Loop aktivitas ---
for i, akt in enumerate(AKTIVITAS):
    teks = akt["teks"]
    skp = akt["skp"]
    jenis = akt["jenis"]
    jenis_label = "Apel" if jenis == 2 else "Aktifitas"

    print(f"\n{'='*60}")
    print(f"📝 AKTIVITAS {i+1}/{len(AKTIVITAS)}: {teks[:40]}... (SKP{skp}, {jenis_label})")
    print(f"{'='*60}")

    if i > 0:
        # Form sudah auto-reset setelah save sebelumnya
        # TIDAK perlu BACK atau Step 1
        print("  (Form auto-reset, langsung isi)")
        time.sleep(1)
        screenshot(f"akt{i+1}_00_form_reset")

    # Step 2: Klik input field & ketik
    print(f"  [Step 2] Input teks...")
    tap(*STEP2_INPUT_XY)
    time.sleep(1)
    input_text(teks)
    time.sleep(1)
    hide_keyboard()
    time.sleep(0.5)
    screenshot(f"akt{i+1}_01_after_type")

    # Step 3: SKP dropdown
    print(f"  [Step 3] SKP → opsi {skp}...")
    tap(*SKP_DROPDOWN_XY)
    time.sleep(1.5)  # Extra wait for dropdown animation
    screenshot(f"akt{i+1}_02_skp_open")
    tap(*SKP_COORDS[skp])
    time.sleep(1.5)
    screenshot(f"akt{i+1}_03_skp_selected")

    # Step 4: Jenis dropdown
    print(f"  [Step 4] Jenis → opsi {jenis} ({jenis_label})...")
    tap(*JENIS_DROPDOWN_XY)
    time.sleep(1.5)  # Extra wait for dropdown animation
    screenshot(f"akt{i+1}_04_jenis_open")
    tap(*JENIS_COORDS[jenis])
    time.sleep(1.5)
    screenshot(f"akt{i+1}_05_jenis_selected")

    # Step 5: Save
    print(f"  [Step 5] SAVE...")
    tap(*STEP5_SAVE_XY)
    time.sleep(10)  # Wait for save + form reset
    screenshot(f"akt{i+1}_06_after_save")

print(f"\n✅ Debug selesai! {len(AKTIVITAS)} aktivitas di-test.")
print(f"   Cek screenshots: {SS_DIR}")
