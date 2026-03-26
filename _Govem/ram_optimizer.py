"""
RAM Optimizer untuk Govem (LDPlayer)
====================================
Memastikan RAM cukup sebelum launch emulator.
Khusus untuk laptop 8 GB RAM.

Strategi:
1. Flush Windows Standby RAM (via PowerShell)
2. Kill proses berat (opsional, configurable)
3. Cek RAM tersedia, tunggu jika kurang
4. Set LDPlayer config ke RAM minimal (1024 MB per instance)
"""

import os
import subprocess
import time
import logging
import json
import ctypes

logger = logging.getLogger(__name__)

# === KONFIGURASI ===
# Minimal RAM (MB) yang harus tersedia sebelum launch emulator
MIN_FREE_RAM_MB = 2000  # 2 GB minimal harus free

# Proses yang BOLEH di-kill untuk menghemat RAM (lowercase)
# Tambahkan/hapus sesuai kebutuhan
KILLABLE_PROCESSES = [
    # Browser (pemakan RAM terbesar)
    "chrome.exe",
    "msedge.exe",
    "firefox.exe",
    # IDE/Editor (opsional — uncomment jika mau)
    # "code.exe",           # VSCode
    # "devenv.exe",         # Visual Studio
    # Discord
    "discord.exe",
    "update.exe",           # Discord updater
    # Lain-lain
    "teams.exe",
    "slack.exe",
    "spotify.exe",
]

# Proses yang TIDAK BOLEH di-kill (safety list)
PROTECTED_PROCESSES = [
    "explorer.exe",
    "pythonw.exe",      # Govem Tray sendiri
    "python.exe",       # Script kita
    "ldplayer.exe",
    "dnplayer.exe",
    "svchost.exe",
    "csrss.exe",
    "winlogon.exe",
    "system",
]

# LDPlayer config path
LDPLAYER_VMS_DIR = r"D:\LDPlayer\LDPlayer9\vms"

def get_available_ram_mb():
    """Ambil RAM yang tersedia (free + standby) dalam MB."""
    try:
        # Pakai PowerShell untuk akurasi (termasuk standby)
        cmd = 'powershell -Command "(Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True,
                               creationflags=0x08000000, timeout=10)
        free_kb = int(result.stdout.strip())
        return free_kb // 1024  # Convert KB to MB
    except Exception as e:
        logger.warning(f"Gagal cek RAM: {e}")
        # Fallback via ctypes
        try:
            kernel32 = ctypes.windll.kernel32
            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]
            mem = MEMORYSTATUSEX()
            mem.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            kernel32.GlobalMemoryStatusEx(ctypes.byref(mem))
            return mem.ullAvailPhys // (1024 * 1024)
        except:
            return 0

def get_ram_usage_percent():
    """Return persentase RAM yang terpakai."""
    try:
        kernel32 = ctypes.windll.kernel32
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        mem = MEMORYSTATUSEX()
        mem.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        kernel32.GlobalMemoryStatusEx(ctypes.byref(mem))
        return mem.dwMemoryLoad
    except:
        return 0

def flush_standby_ram():
    """Flush Windows Standby List (bebaskan cached RAM).
    Butuh admin privilege — jika gagal, skip tanpa error.
    """
    logger.info("🧹 Flushing standby RAM...")
    try:
        # Metode 1: Via PowerShell (tidak perlu tool external)
        ps_script = '''
        # Clear Working Sets (force apps to release unused pages)
        Get-Process | Where-Object {$_.WorkingSet64 -gt 100MB} | ForEach-Object {
            try {
                $handle = $_.Handle
                # Trim working set
                [System.Diagnostics.Process]::GetProcessById($_.Id).MinWorkingSet = [IntPtr]::new(204800)
            } catch {}
        }
        # Force GC pada .NET processes
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        '''
        subprocess.run(["powershell", "-Command", ps_script],
                       capture_output=True, timeout=15,
                       creationflags=0x08000000)
        logger.info("✅ Standby RAM flushed")
    except Exception as e:
        logger.warning(f"⚠️ Flush standby gagal (mungkin perlu admin): {e}")

def kill_heavy_processes(dry_run=False):
    """Kill proses-proses berat yang ada di KILLABLE_PROCESSES.
    Return: list nama proses yang di-kill.
    """
    killed = []

    try:
        # Ambil daftar proses + RAM usage
        cmd = 'powershell -Command "Get-Process | Select-Object Name, Id, @{N=\'MemMB\';E={[math]::Round($_.WorkingSet64/1MB)}} | Sort-Object MemMB -Descending | ConvertTo-Json"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True,
                               creationflags=0x08000000, timeout=15)

        processes = json.loads(result.stdout)
        if not isinstance(processes, list):
            processes = [processes]

        for proc in processes:
            proc_name = (proc.get("Name", "") + ".exe").lower()
            mem_mb = proc.get("MemMB", 0)
            pid = proc.get("Id", 0)

            if proc_name in [p.lower() for p in KILLABLE_PROCESSES] and mem_mb > 50:
                if dry_run:
                    logger.info(f"   [DRY RUN] Akan kill: {proc_name} (PID {pid}, {mem_mb} MB)")
                else:
                    logger.info(f"   🔪 Killing {proc_name} (PID {pid}, {mem_mb} MB)")
                    try:
                        subprocess.run(f'taskkill /F /PID {pid}', shell=True,
                                      capture_output=True, creationflags=0x08000000, timeout=5)
                        killed.append(f"{proc_name} ({mem_mb}MB)")
                    except:
                        pass
    except Exception as e:
        logger.warning(f"⚠️ Gagal scan proses: {e}")

    return killed

def optimize_ldplayer_ram(emulator_index, ram_mb=1024, cpu_cores=2):
    """Set RAM allocation LDPlayer instance ke nilai minimal.
    Default: 1024 MB RAM, 2 CPU cores (cukup untuk Govem yang ringan).

    PENTING: Emulator HARUS dalam keadaan MATI saat mengubah config!
    Config path: vms/config/leidian{index}.config
    Fields: advancedSettings.memorySize, advancedSettings.cpuCount
    """
    config_file = os.path.join(LDPLAYER_VMS_DIR, "config", f"leidian{emulator_index}.config")

    if not os.path.exists(config_file):
        logger.warning(f"⚠️ Config emulator {emulator_index} tidak ditemukan: {config_file}")
        return False

    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)

        changed = False

        # Optimize RAM
        current_ram = config.get("advancedSettings.memorySize", 0)
        if current_ram > ram_mb:
            config["advancedSettings.memorySize"] = ram_mb
            logger.info(f"   [Emu {emulator_index}] RAM: {current_ram} → {ram_mb} MB")
            changed = True
        else:
            logger.info(f"   [Emu {emulator_index}] RAM sudah optimal: {current_ram} MB")

        # Optimize CPU
        current_cpu = config.get("advancedSettings.cpuCount", 4)
        if current_cpu > cpu_cores:
            config["advancedSettings.cpuCount"] = cpu_cores
            logger.info(f"   [Emu {emulator_index}] CPU: {current_cpu} → {cpu_cores} cores")
            changed = True

        if changed:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            logger.info(f"   ✅ [Emu {emulator_index}] Config updated!")

        return True

    except Exception as e:
        logger.warning(f"⚠️ [Emu {emulator_index}] Gagal optimasi config: {e}")
        return False

def ensure_ram_available(min_mb=None, max_wait_seconds=60, aggressive=False):
    """
    Pastikan RAM cukup sebelum launch emulator.

    Flow:
    1. Cek RAM → jika cukup, return True
    2. Flush standby RAM
    3. Cek lagi → jika cukup, return True
    4. Jika aggressive=True, kill proses berat
    5. Tunggu sampai max_wait_seconds

    Return: True jika RAM cukup, False jika timeout
    """
    if min_mb is None:
        min_mb = MIN_FREE_RAM_MB

    # Step 1: Cek awal
    free = get_available_ram_mb()
    usage_pct = get_ram_usage_percent()
    logger.info(f"📊 RAM: {free} MB free ({usage_pct}% used)")

    if free >= min_mb:
        logger.info(f"✅ RAM cukup ({free} MB >= {min_mb} MB)")
        return True

    logger.warning(f"⚠️ RAM kurang! {free} MB < {min_mb} MB target")

    # Step 2: Flush standby
    flush_standby_ram()
    time.sleep(2)

    free = get_available_ram_mb()
    if free >= min_mb:
        logger.info(f"✅ RAM cukup setelah flush ({free} MB)")
        return True

    # Step 3: Kill proses berat (jika aggressive)
    if aggressive:
        logger.info("🔪 Mode agresif: kill proses berat...")
        killed = kill_heavy_processes()
        if killed:
            logger.info(f"   Killed: {', '.join(killed)}")
            time.sleep(3)
            free = get_available_ram_mb()
            if free >= min_mb:
                logger.info(f"✅ RAM cukup setelah kill ({free} MB)")
                return True

    # Step 4: Tunggu (mungkin proses lain selesai)
    logger.info(f"⏳ Menunggu RAM tersedia (max {max_wait_seconds}s)...")
    waited = 0
    while waited < max_wait_seconds:
        time.sleep(5)
        waited += 5
        free = get_available_ram_mb()
        if free >= min_mb:
            logger.info(f"✅ RAM cukup setelah {waited}s ({free} MB)")
            return True
        if waited % 15 == 0:
            logger.info(f"   ... RAM: {free} MB (target: {min_mb} MB, waited: {waited}s)")

    # Timeout — lanjut saja (hope for the best)
    logger.warning(f"⚠️ RAM timeout! Lanjut dengan {free} MB (target: {min_mb} MB)")
    return False

def get_ram_report():
    """Return string laporan RAM untuk notifikasi/logging."""
    free = get_available_ram_mb()
    pct = get_ram_usage_percent()

    status = "[OK]" if free > 2000 else "[LOW]" if free > 1000 else "[CRITICAL]"
    return f"{status} RAM: {free} MB free ({pct}% used)"

# === STANDALONE MODE ===
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s')

    print("=" * 50)
    print("  RAM Optimizer untuk Govem (LDPlayer)")
    print("=" * 50)

    # Laporan awal
    print(f"\n{get_ram_report()}")

    # Scan proses berat
    print("\nProses yang bisa di-kill:")
    kill_heavy_processes(dry_run=True)

    # Optimasi LDPlayer config
    print("\nOptimasi LDPlayer config:")
    for idx in range(3):
        optimize_ldplayer_ram(idx, ram_mb=1024)

    # Flush
    print("\nFlushing RAM...")
    before = get_available_ram_mb()
    flush_standby_ram()
    time.sleep(2)
    after = get_available_ram_mb()
    print(f"   Sebelum: {before} MB -> Sesudah: {after} MB (freed: {after - before} MB)")

    print(f"\n{get_ram_report()}")
    print("\nSelesai! Jalankan lagi kapan saja untuk membersihkan RAM.")
