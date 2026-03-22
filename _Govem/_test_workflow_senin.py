"""
Test Workflow Senin — Simulasi full flow pagi + sore
Verifikasi: auto-minimize semua emulator, aktivitas Istri (pagi), aktivitas Suami (sore),
            batch screenshot + 1 notifikasi per sesi.

CATATAN: Ini TIDAK menjalankan absen sungguhan (skip GPS + klik absen).
         Fokus test: launch/minimize, isi aktivitas, screenshot, notifikasi.
"""
import os, sys, time, subprocess

# Setup path agar bisa import dari _Govem
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
LDCONSOLE = os.path.join(LDPLAYER_PATH, "ldconsole.exe")
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")
PACKAGE = "go.id.tapinkab.govem"

def cmd(command):
    flags = 0x08000000 if os.name == 'nt' else 0
    if isinstance(command, list):
        r = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=flags)
    else:
        r = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=flags)
    return r.stdout.strip()

def is_running(idx):
    s = cmd(f'"{LDCONSOLE}" list2')
    for line in s.splitlines():
        parts = line.split(",")
        if len(parts) > 4 and parts[0] == str(idx) and parts[4] == "1":
            return True
    return False

def launch_and_minimize(idx, name):
    """Launch emulator + auto-minimize (sama seperti launch_emulator di Engine)."""
    print(f"\n  [{name}] Launching emulator {idx}...")
    cmd(f'"{LDCONSOLE}" launch --index {idx}')

    for attempt in range(30):
        cmd(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
        boot = cmd(f'"{LDCONSOLE}" adb --index {idx} --command "shell getprop sys.boot_completed"')
        if "1" in boot:
            print(f"  [{name}] Ready setelah {attempt*2}s, minimized!")
            for _ in range(3):
                time.sleep(1)
                cmd(f'"{LDCONSOLE}" sortWnd --index {idx} --minimize')
            return True
        time.sleep(2)
        if attempt % 5 == 0:
            print(f"  [{name}] Waiting boot ({attempt*2}s)...")

    print(f"  [{name}] WARNING: Boot timeout!")
    return False

def get_serial(idx):
    devices = cmd(f'"{ADB}" devices')
    for p in [5554 + (idx*2), 5556, 5558, 5560]:
        if f"emulator-{p}" in devices:
            return f"emulator-{p}"
        if f"127.0.0.1:{p}" in devices:
            return f"127.0.0.1:{p}"
    # Try connect
    port = 5554 + (idx*2)
    cmd(f'"{ADB}" connect 127.0.0.1:{port}')
    time.sleep(2)
    return f"127.0.0.1:{port}"

def screenshot(idx, serial, name):
    """Restart app → dashboard depan → screenshot."""
    import tempfile
    # Restart app agar kembali ke dashboard depan
    cmd(f'"{LDCONSOLE}" killapp --index {idx} --packagename {PACKAGE}')
    time.sleep(2)
    cmd(f'"{LDCONSOLE}" runapp --index {idx} --packagename {PACKAGE}')
    print(f"  [{name}] Restart app → dashboard (10s)...")
    time.sleep(10)

    path = os.path.join(tempfile.gettempdir(), f"govem_final_{name.lower()}.png")
    r = subprocess.run([ADB, "-s", serial, "exec-out", "screencap", "-p"],
                       capture_output=True, timeout=10,
                       creationflags=0x08000000 if os.name == 'nt' else 0)
    if r.returncode == 0 and r.stdout:
        with open(path, 'wb') as f:
            f.write(r.stdout)
        print(f"  [{name}] Screenshot: {path}")
        return path
    return None

def send_notification(job_type, screenshot_paths):
    names = [os.path.basename(p).replace("govem_final_", "").replace(".png", "").capitalize()
             for p in screenshot_paths if p]
    summary = ", ".join(names) if names else "Selesai"
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
    print(f"  Notifikasi: {job_type} — {summary}")
    for p in screenshot_paths:
        if p and os.path.exists(p):
            os.startfile(p)

def test_aktivitas_istri(serial):
    """Test 2 aktivitas Istri (Senin: SKP1+Apel, SKP2+Aktifitas)."""
    import V23_aktivitas_Istri as v23i
    # Override: hanya 2 item pertama dari Senin
    items = v23i.AKTIVITAS_ISTRI[0][:2]
    print(f"\n  [Istri] Testing {len(items)} aktivitas...")

    # Restart app
    cmd(f'"{LDCONSOLE}" killapp --index 1 --packagename {PACKAGE}')
    time.sleep(2)
    cmd(f'"{LDCONSOLE}" runapp --index 1 --packagename {PACKAGE}')
    print("  [Istri] Waiting app ready (25s)...")
    time.sleep(25)

    # Step 1: Navigate
    v23i.adb_click(serial, *v23i.STEP1_TAP_1)
    time.sleep(2)
    v23i.adb_click(serial, *v23i.STEP1_TAP_2)
    time.sleep(3)

    for i, act in enumerate(items):
        print(f"  [Istri] [{i+1}/{len(items)}] {act['teks'][:50]}... SKP{act['skp']} J{act['jenis']}")

        # Step 2: Focus input
        v23i.adb_click(serial, *v23i.STEP2_INPUT_XY)
        time.sleep(0.8)

        # Clear
        for _ in range(2):
            v23i.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "123"])
            v23i.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "KEYCODE_MOVE_HOME"])
            v23i.run_command(f'"{ADB}" -s {serial} shell input keyevent --longpress 112 123')
            time.sleep(0.1)
            v23i.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "67"])
            time.sleep(0.1)
        time.sleep(0.5)

        # Type
        v23i.adb_input_text(serial, act["teks"].replace("'", "").replace('"', ""))
        time.sleep(1)
        v23i.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(0.5)

        # Step 3: SKP
        v23i.adb_click(serial, *v23i.SKP_DROPDOWN_XY)
        time.sleep(1.5)
        v23i.adb_click(serial, *v23i.SKP_COORDS[act["skp"]])
        time.sleep(1.5)

        # Step 4: Jenis
        v23i.adb_click(serial, *v23i.JENIS_DROPDOWN_XY)
        time.sleep(1.5)
        v23i.adb_click(serial, *v23i.JENIS_COORDS[act["jenis"]])
        time.sleep(1.5)

        # Step 5: Simpan
        v23i.adb_click(serial, *v23i.STEP5_SAVE_XY)
        time.sleep(12)

    print(f"  [Istri] {len(items)} aktivitas terisi!")
    return True

def test_aktivitas_suami(serial):
    """Test 2 aktivitas Suami (Senin)."""
    import V23_aktivitas_Suami as v23s
    import importlib
    importlib.reload(v23s)

    # Restart app
    cmd(f'"{LDCONSOLE}" killapp --index 0 --packagename {PACKAGE}')
    time.sleep(2)
    cmd(f'"{LDCONSOLE}" runapp --index 0 --packagename {PACKAGE}')
    print("  [Suami] Waiting app ready (25s)...")
    time.sleep(25)

    # Panggil run_hybrid_automation dengan override 2 item saja
    # Terlalu complex — langsung mirror logic Istri test
    print(f"\n  [Suami] Testing 2 aktivitas...")

    # Step 1: Navigate
    v23s.adb_click(serial, *v23s.STEP1_TAP_1)
    time.sleep(2)
    v23s.adb_click(serial, *v23s.STEP1_TAP_2)
    time.sleep(3)

    test_items = [
        {"teks": "Test Suami aktivitas 1", "skp": 1, "jenis": 2},
        {"teks": "Test Suami aktivitas 2", "skp": 3, "jenis": 1},
    ]

    for i, act in enumerate(test_items):
        print(f"  [Suami] [{i+1}/{len(test_items)}] SKP{act['skp']} J{act['jenis']}")

        v23s.adb_click(serial, *v23s.STEP2_INPUT_XY)
        time.sleep(0.8)

        # Clear
        for _ in range(2):
            v23s.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "123"])
            v23s.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "KEYCODE_MOVE_HOME"])
            v23s.run_command(f'"{ADB}" -s {serial} shell input keyevent --longpress 112 123')
            time.sleep(0.1)
            v23s.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "67"])
            time.sleep(0.1)
        time.sleep(0.5)

        v23s.adb_input_text(serial, act["teks"])
        time.sleep(1)
        v23s.run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(0.5)

        # SKP (Suami coords)
        v23s.adb_click(serial, *v23s.SKP_DROPDOWN_XY)
        time.sleep(1.5)
        v23s.adb_click(serial, *v23s.SKP_COORDS[act["skp"]])
        time.sleep(1.5)

        # Jenis (Suami coords)
        v23s.adb_click(serial, *v23s.JENIS_DROPDOWN_XY)
        time.sleep(1.5)
        v23s.adb_click(serial, *v23s.JENIS_COORDS[act["jenis"]])
        time.sleep(1.5)

        # Simpan
        v23s.adb_click(serial, *v23s.STEP5_SAVE_XY)
        time.sleep(12)

    print(f"  [Suami] {len(test_items)} aktivitas terisi!")
    return True

def kill_all():
    for idx in [0, 1, 2]:
        cmd(f'"{LDCONSOLE}" quit --index {idx}')
    time.sleep(3)

def main():
    print("=" * 60)
    print("  WORKFLOW TEST: Simulasi Hari Senin")
    print("  Auto-minimize + Aktivitas + Batch Screenshot + Notifikasi")
    print("=" * 60)

    # Kill semua emulator dulu
    print("\n[PREP] Mematikan semua emulator...")
    kill_all()

    # ===================================================
    # SESI PAGI: Pancingan → Suami (absen only) → Istri (absen + aktivitas)
    # ===================================================
    print("\n" + "=" * 60)
    print("  SESI PAGI (Simulasi)")
    print("=" * 60)

    # 1. Pancingan — launch + minimize only (no absen)
    print("\n[PAGI 1/3] Pancingan — Launch + Auto-Minimize")
    launch_and_minimize(2, "Pancingan")

    # 2. Suami — launch + minimize (skip absen, just screenshot dashboard)
    print("\n[PAGI 2/3] Suami — Launch + Auto-Minimize")
    launch_and_minimize(0, "Suami")
    serial_suami = get_serial(0)
    print(f"  [Suami] ADB: {serial_suami}")

    # 3. Istri — launch + minimize + 2 test aktivitas
    print("\n[PAGI 3/3] Istri — Launch + Auto-Minimize + 2 Aktivitas")
    launch_and_minimize(1, "Istri")
    serial_istri = get_serial(1)
    print(f"  [Istri] ADB: {serial_istri}")
    test_aktivitas_istri(serial_istri)

    # Batch screenshot + notifikasi PAGI
    print("\n[PAGI] Batch Screenshot (restart app → dashboard depan)...")
    pagi_ss = []
    ss = screenshot(0, serial_suami, "Suami")
    if ss: pagi_ss.append(ss)
    ss = screenshot(1, serial_istri, "Istri")
    if ss: pagi_ss.append(ss)
    send_notification("Absen Pagi", pagi_ss)

    input("\n>>> Cek hasil PAGI. Tekan ENTER untuk lanjut ke SORE...")

    # ===================================================
    # SESI SORE: Suami (absen + aktivitas) → Istri (absen only)
    # ===================================================
    print("\n" + "=" * 60)
    print("  SESI SORE (Simulasi)")
    print("=" * 60)

    # Suami — 2 test aktivitas
    print("\n[SORE 1/2] Suami — 2 Aktivitas")
    test_aktivitas_suami(serial_suami)

    # Istri — skip (hanya absen, kita screenshot saja)
    print("\n[SORE 2/2] Istri — Screenshot dashboard saja")

    # Batch screenshot + notifikasi SORE
    print("\n[SORE] Batch Screenshot (restart app → dashboard depan)...")
    sore_ss = []
    ss = screenshot(0, serial_suami, "Suami")
    if ss: sore_ss.append(ss)
    ss = screenshot(1, serial_istri, "Istri")
    if ss: sore_ss.append(ss)
    send_notification("Absen Sore", sore_ss)

    print("\n" + "=" * 60)
    print("  TEST SELESAI!")
    print("=" * 60)

if __name__ == "__main__":
    main()
