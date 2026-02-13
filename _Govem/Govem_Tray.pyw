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
    
    icon.notify(status, "Govem Scheduler Status")

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
            if user_idx == 0:
                Govem_Engine.trigger_activity(user) # Suami
            else:
                Govem_Engine.trigger_activity_istri(user) # Istri
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
    
    # Submenu Settings
    settings_menu = pystray.Menu(
        item('🗑️ Hapus Autostart (Windows Boot)', remove_autostart),
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
