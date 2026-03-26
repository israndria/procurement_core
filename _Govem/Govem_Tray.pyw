"""
Govem Scheduler System Tray
===========================
Menampilkan icon di system tray (pojok kanan bawah) untuk menunjukkan
status scheduler sedang aktif.

Fitur:
- Icon hijau = Scheduler aktif
- Klik kanan untuk menu (Exit, Status, dll)
- Tooltip menunjukkan info scheduler
"""

import os
import sys
import subprocess
import threading
import time
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item

# Path ke script utama
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GOVEM_ENGINE = os.path.join(SCRIPT_DIR, "Govem_Engine.py")
PYTHON_EXE = sys.executable

# Fix pythonw: stdout/stderr = None → redirect ke devnull agar print() tidak crash
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')

# Status global
scheduler_running = False
scheduler_thread = None

def create_image(color='green'):
    """Buat icon sederhana dengan warna tertentu."""
    width = 64
    height = 64
    image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # Background circle
    if color == 'green':
        fill_color = (50, 205, 50, 255)  # Lime green
    elif color == 'yellow':
        fill_color = (255, 215, 0, 255)  # Gold
    elif color == 'red':
        fill_color = (220, 20, 60, 255)  # Crimson
    else:
        fill_color = (128, 128, 128, 255)  # Gray
    
    # Draw filled circle
    draw.ellipse([4, 4, width-4, height-4], fill=fill_color, outline=(0, 0, 0, 255), width=2)
    
    # Draw "G" letter
    draw.text((20, 15), "G", fill=(255, 255, 255, 255))
    
    return image

def run_scheduler(reset=False):
    """Jalankan Govem_Engine.py di background."""
    global scheduler_running
    
    try:
        # Import dan jalankan scheduler
        sys.path.insert(0, SCRIPT_DIR)
        import Govem_Engine
        
        # Set flag
        scheduler_running = True
        
        # Jalankan scheduler (ini blocking)
        target_users = Govem_Engine.USERS
        Govem_Engine.run_scheduler(target_users, force_reset=reset)
        
    except Exception as e:
        print(f"Error: {e}")
        scheduler_running = False

def start_scheduler(icon, item):
    """Menu: Start scheduler (Normal)."""
    global scheduler_thread, scheduler_running
    
    if not scheduler_running:
        scheduler_thread = threading.Thread(target=run_scheduler, args=(False,), daemon=True)
        scheduler_thread.start()
        icon.icon = create_image('green')
        icon.notify("Scheduler Started", "Govem Scheduler")

def start_scheduler_force(icon, item):
    """Menu: Start scheduler + Reset History Hari Ini."""
    global scheduler_thread, scheduler_running
    
    if scheduler_running:
        # Jika sedang jalan, stop dulu (user harus stop manual sebenarnya, tapi kita notify)
        icon.notify("Harap STOP dulu sebelum Force Run!", "Govem Scheduler")
        return

    # Start dengan reset=True
    scheduler_thread = threading.Thread(target=run_scheduler, args=(True,), daemon=True)
    scheduler_thread.start()
    icon.icon = create_image('yellow') # Kuning = Warning/Special mode
    icon.notify("Scheduler FORCE START!\nHistory hari ini dihapus & Catch-Up dijalankan.", "Govem Scheduler")

def get_status(icon, item):
    """Menu: Show status."""
    if scheduler_running:
        status = "✅ Scheduler AKTIF"
    else:
        status = "❌ Scheduler TIDAK AKTIF"
    
    # Tampilkan status per-akun
    import Govem_Engine
    akun_info = []
    for u_name in ['Suami', 'Istri', 'Pancingan']:
        st = "✅" if Govem_Engine.is_user_enabled(u_name) else "⏸️"
        akun_info.append(f"{u_name}: {st}")
    
    detail = " | ".join(akun_info)
    icon.notify(f"{status}\n{detail}", "Govem Scheduler Status")

def on_exit(icon, item):
    """Menu: Exit & Stop Scheduler - Hentikan semua proses."""
    global scheduler_running
    scheduler_running = False
    icon.stop()

# ... (remove_autostart, tools launcher tetap sama)

def remove_autostart(icon, item):
    """Menu: Remove autostart (tidak akan auto-start saat Windows boot)."""
    import os
    startup_folder = os.path.join(os.environ['APPDATA'], 
        'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    vbs_file = os.path.join(startup_folder, 'Govem_Scheduler.vbs')
    
    if os.path.exists(vbs_file):
        os.remove(vbs_file)
        icon.notify("Autostart dihapus!\nScheduler tidak akan jalan otomatis saat Windows startup.", 
                   "Govem Scheduler")
    else:
        icon.notify("Tidak ada autostart yang perlu dihapus.", "Govem Scheduler")

# --- LAUNCHER V19, V20, V21 (TANPA CONSOLE POPUP) ---
def launch_v19(icon, item):
    """Menu: Subscribe Jadwal SPSE (V19)."""
    bat_path = r"D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\V19_Scheduler.bat"
    # Gunakan shell=True dan CREATE_NO_WINDOW untuk menghindari popup
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    subprocess.Popen(['cmd', '/c', 'start', '', bat_path], startupinfo=startupinfo)
    icon.notify("V19 Scheduler diluncurkan!", "Govem Tools")

def launch_v20(icon, item):
    """Menu: Scrap Data Tender/Non Tender (V20)."""
    bat_path = r"D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\V20_Scrapper.bat"
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    subprocess.Popen(['cmd', '/c', 'start', '', bat_path], startupinfo=startupinfo)
    icon.notify("V20 Scrapper diluncurkan!", "Govem Tools")

def launch_v21(icon, item):
    """Menu: Scrap Produk Katalog 6 (V21)."""
    bat_path = r"D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\V21_Katalog.bat"
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    subprocess.Popen(['cmd', '/c', 'start', '', bat_path], startupinfo=startupinfo)
    icon.notify("V21 Katalog diluncurkan!", "Govem Tools")

# --- MANUAL TRIGGER HANDLERS ---
def run_manual_engine(mode, user_idx):
    """Helper untuk menjalankan task di thread terpisah."""
    try:
        sys.path.insert(0, SCRIPT_DIR)
        import Govem_Engine
        user = Govem_Engine.USERS[user_idx]

        if mode == 'pagi':
            Govem_Engine.absen_pagi(user)
        elif mode == 'sore':
            Govem_Engine.absen_sore(user)
        elif mode == 'aktivitas':
            # Manual trigger: skip_nav=True (user sudah di form aktivitas)
            if user_idx == 0:
                Govem_Engine.trigger_activity(user, skip_nav=True)
            else:
                Govem_Engine.trigger_activity_istri(user, skip_nav=True)
        elif mode == 'pancingan':
             # Pancingan logic (absen_pagi/sore is same: Launch Only)
             Govem_Engine.absen_pagi(user)
                
                
    except Exception as e:
        print(f"Error executing manual task: {e}")

def manual_pagi_suami(icon, item):
    icon.notify("🚀 Memulai Absen Pagi SUAMI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('pagi', 0), daemon=True).start()

def manual_pagi_istri(icon, item):
    icon.notify("🚀 Memulai Absen Pagi ISTRI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('pagi', 1), daemon=True).start()

def manual_sore_suami(icon, item):
    icon.notify("🚀 Memulai Absen Sore SUAMI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('sore', 0), daemon=True).start()

def manual_sore_istri(icon, item):
    icon.notify("🚀 Memulai Absen Sore ISTRI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('sore', 1), daemon=True).start()

def manual_akt_suami(icon, item):
    icon.notify("🚀 Memulai Aktivitas SUAMI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('aktivitas', 0), daemon=True).start()

def manual_akt_istri(icon, item):
    icon.notify("🚀 Memulai Aktivitas ISTRI...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('aktivitas', 1), daemon=True).start()

def manual_pancingan(icon, item):
    icon.notify("🎣 Membuka Akun Pancingan...", "Manual Trigger")
    threading.Thread(target=lambda: run_manual_engine('pancingan', 2), daemon=True).start()


def setup_tray():
    """Setup system tray icon."""
    global scheduler_thread, scheduler_running
    
    # Buat icon
    image = create_image('green')
    
    # Submenu Tools
    tools_menu = pystray.Menu(
        item('📅 Subscribe Jadwal SPSE (V19)', launch_v19),
        item('📊 Scrap Tender/Non Tender (V20)', launch_v20),
        item('🛒 Scrap Katalog 6 (V21)', launch_v21),
    )
    
    # Submenu Manual Trigger (NEW)
    manual_menu = pystray.Menu(
        item('☀️ Absen Pagi - SUAMI', manual_pagi_suami),
        item('☀️ Absen Pagi - ISTRI', manual_pagi_istri),
        item('🎣 Mulai Pancingan (Test/Manual)', manual_pancingan),
        pystray.Menu.SEPARATOR,
        item('🌙 Absen Sore - SUAMI', manual_sore_suami),
        item('🌙 Absen Sore - ISTRI', manual_sore_istri),
        pystray.Menu.SEPARATOR,
        item('📝 Isi Aktivitas - SUAMI', manual_akt_suami),
        item('📝 Isi Aktivitas - ISTRI', manual_akt_istri),
    )
    
    def optimize_ram(icon, item):
        """Menu: Bersihkan RAM."""
        try:
            sys.path.insert(0, SCRIPT_DIR)
            from ram_optimizer import flush_standby_ram, get_ram_report, kill_heavy_processes
            before_report = get_ram_report()
            flush_standby_ram()
            killed = kill_heavy_processes()
            import time as _t; _t.sleep(2)
            after_report = get_ram_report()
            killed_info = f"\nKilled: {', '.join(killed)}" if killed else ""
            icon.notify(f"Sebelum: {before_report}\nSesudah: {after_report}{killed_info}", "RAM Optimizer")
        except Exception as e:
            icon.notify(f"Error: {e}", "RAM Optimizer")

    def check_ram_status(icon, item):
        """Menu: Cek status RAM."""
        try:
            sys.path.insert(0, SCRIPT_DIR)
            from ram_optimizer import get_ram_report
            icon.notify(get_ram_report(), "RAM Status")
        except Exception as e:
            icon.notify(f"Error: {e}", "RAM Status")

    # Submenu Settings
    settings_menu = pystray.Menu(
        item('🧹 Bersihkan RAM', optimize_ram),
        item('📊 Cek RAM', check_ram_status),
        pystray.Menu.SEPARATOR,
        item('🗑️ Hapus Autostart (Windows Boot)', remove_autostart),
    )
    
    # --- Toggle Per-Akun (NEW FEATURE) ---
    # Handler Factory: buat handler unik per user agar closure benar
    def make_toggle_handler(user_name):
        def handler(icon, menu_item):
            import Govem_Engine
            is_aktif = Govem_Engine.toggle_user(user_name)
            status = "AKTIF ✅" if is_aktif else "NONAKTIF ⏸️"
            icon.notify(f"{user_name}: {status}", "Toggle Akun")
        return handler
    
    # Checked callback: centang jika user AKTIF (tidak di-disable)
    def make_checked_callback(user_name):
        def is_checked(menu_item):
            import Govem_Engine
            return Govem_Engine.is_user_enabled(user_name)
        return is_checked
    
    toggle_menu = pystray.Menu(
        item('👨 Suami', make_toggle_handler('Suami'), checked=make_checked_callback('Suami')),
        item('👩 Istri', make_toggle_handler('Istri'), checked=make_checked_callback('Istri')),
        item('🎣 Pancingan', make_toggle_handler('Pancingan'), checked=make_checked_callback('Pancingan')),
    )
    
    # Menu items - LEBIH JELAS
    menu = pystray.Menu(
        item('▶️ Manual Trigger', manual_menu), # NEW FEATURE
        item('🧰 Tools Lainnya', tools_menu),
        pystray.Menu.SEPARATOR,
        item('📊 Cek Status', get_status),
        item('▶️ Start Normal', start_scheduler),
        item('♻️ Start & FORCE RUN (Reset Hari Ini)', start_scheduler_force),
        pystray.Menu.SEPARATOR,
        item('🔀 Aktifkan/Nonaktifkan Akun', toggle_menu),  # PER-USER TOGGLE
        item('⚙️ Pengaturan', settings_menu),
        item('🛑 STOP & Exit', on_exit)
    )
    
    # Buat tray icon
    icon = pystray.Icon(
        "govem_scheduler",
        image,
        "Govem Scheduler - Aktif",
        menu
    )
    
    # Auto-start scheduler
    def auto_start():
        time.sleep(2)  # Tunggu icon muncul
        run_scheduler()
    
    scheduler_thread = threading.Thread(target=auto_start, daemon=True)
    scheduler_thread.start()
    
    # Run icon (blocking)
    icon.run()

if __name__ == "__main__":
    setup_tray()
