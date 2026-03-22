"""
Debug: Isi 7 aktivitas Suami (Selasa) dari halaman menu Aktivitas.
User sudah klik tombol kuning → sudah di halaman "Aktivitas".
"""
import os, sys, time, subprocess, datetime

ADB = r"D:\LDPlayer\LDPlayer9\adb.exe"
SERIAL = "emulator-5554"

SS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots", "debug")
os.makedirs(SS_DIR, exist_ok=True)

# Koordinat dari uiautomator dump (1600x900)
STEP1_TAP_2 = (500, 331)   # "Buat Aktifitas harian" dari menu Aktivitas
INPUT_XY = (246, 456)
SAVE_XY = (795, 897)
SKP_DD = (500, 318)
JENIS_DD = (1130, 318)
SKP = {1: (800, 518), 2: (800, 617), 3: (800, 716), 4: (800, 815)}
JENIS = {1: (800, 508), 2: (800, 592)}

# Template Rutinitas (Selasa = 7 aktivitas)
AKTIVITAS = [
    {"teks": "melaksanakan apel pagi", "skp": 4, "jenis": 2},
    {"teks": "Menelaah dan menganalisa dokumen peraturan presiden no 46 tahun 2025 tentang perubahan kedua atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah SE Kep LKPP No. 1 2025 tentang Penjelasan Atas Pelaksanaan Perpres No. 46 2025 tentang Perubahan Kedua Atas Perpres No. 16 2018 Pada Masa Transisi", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Presiden No. 12 tahun 2021 tentang perubahan pertama atas peraturan presiden no 16 tahun 2018 tentang pengadaan barang/jasa pemerintah", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 12 Tahun 2021 tentang Pedoman Pelaksanaan Pengadaan Barang/Jasa Pemerintah Melalui Penyedia", "skp": 4, "jenis": 1},
    {"teks": "Menelaah dan menganalisa Peraturan Lembaga LKPP Nomor 3 Tahun 2021 tentang Pedoman Swakelola", "skp": 4, "jenis": 1},
    {"teks": "melaksanakan apel sore", "skp": 4, "jenis": 2},
]

def run_cmd(command):
    flags = 0x08000000 if os.name == 'nt' else 0
    if isinstance(command, list):
        return subprocess.run(command, capture_output=True, text=True, creationflags=flags).stdout.strip()
    return subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=flags).stdout.strip()

def tap(x, y):
    run_cmd([ADB, "-s", SERIAL, "shell", "input", "swipe", str(x), str(y), str(x), str(y), "150"])

def ss(label):
    ts = datetime.datetime.now().strftime("%H%M%S")
    path = os.path.join(SS_DIR, f"{ts}_{label}.png")
    r = subprocess.run([ADB, "-s", SERIAL, "exec-out", "screencap", "-p"],
                       capture_output=True, timeout=10, creationflags=0x08000000 if os.name=='nt' else 0)
    if r.returncode == 0 and r.stdout:
        with open(path, 'wb') as f: f.write(r.stdout)
        print(f"  📸 {label}")

def input_text(text):
    escaped = text.replace(" ", "%s").replace("'","").replace('"',"").replace("&","").replace("(","").replace(")","").replace(";","")
    run_cmd([ADB, "-s", SERIAL, "shell", "input", "text", escaped])

# Mulai dari halaman menu "Aktivitas" → klik "Buat Aktifitas harian"
print("📱 Klik 'Buat Aktifitas harian'...")
tap(*STEP1_TAP_2)
time.sleep(3)
ss("00_form_awal")

for i, akt in enumerate(AKTIVITAS):
    teks = akt["teks"]
    skp = akt["skp"]
    jenis = akt["jenis"]
    jlabel = "Apel" if jenis == 2 else "Aktifitas"

    print(f"\n{'='*50}")
    print(f"📝 [{i+1}/7] {teks[:50]}... (SKP{skp}, {jlabel})")

    if i > 0:
        time.sleep(1)  # Form auto-reset

    # Input teks
    tap(*INPUT_XY)
    time.sleep(0.8)
    input_text(teks)
    time.sleep(0.8)
    run_cmd([ADB, "-s", SERIAL, "shell", "input", "keyevent", "111"])  # hide keyboard
    time.sleep(0.5)

    # SKP
    tap(*SKP_DD)
    time.sleep(1.2)
    tap(*SKP[skp])
    time.sleep(1.2)

    # Jenis
    tap(*JENIS_DD)
    time.sleep(1.2)
    tap(*JENIS[jenis])
    time.sleep(1.2)

    # Save
    tap(*SAVE_XY)
    time.sleep(8)
    ss(f"akt{i+1}_saved")
    print(f"  ✅ Saved!")

print(f"\n🎉 Selesai! 7 aktivitas terisi.")
