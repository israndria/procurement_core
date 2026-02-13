"""
GOVEM - Test Manual Pengisian Aktivitas
Untuk testing tanpa perlu absensi terlebih dahulu

FITUR:
- Background Mode: LDPlayer bisa diminimize saat script berjalan
- Delay dipercepat untuk eksekusi lebih cepat
"""
import sys
import os
import datetime

# Import dari V23_Aktivitas_Auto
sys.path.insert(0, os.path.dirname(__file__))
from V23_aktivitas_Suami import run_hybrid_automation

def print_banner():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("=" * 50)
    print("   🧪 TEST MANUAL - PENGISIAN AKTIVITAS")
    print("=" * 50)
    print()

def test_menu():
    print_banner()
    
    print("⚠️  PENTING:")
    print("   Pastikan LDPlayer sudah:")
    print("   1. Running (boleh minimize)")
    print("   2. Sudah LOGIN ke aplikasi GOVEM")
    print("   3. Berada di HALAMAN AWAL (Dashboard)")
    print()
    print("   💡 BACKGROUND MODE: LDPlayer bisa diminimize!")
    print("   Script tetap berjalan di background.")
    print()
    
    print("Pilih Emulator untuk Test:")
    print("1. 🧑 Suami (Emulator 0) - Background Mode")
    print("2. 👩 Istri (Emulator 1) - Background Mode")
    print("3. 🧑 Suami - Foreground Mode (bawa ke depan)")
    print("0. Keluar")
    print()
    
    choice = input("Pilihan: ").strip()
    
    if choice == '0':
        print("Keluar...")
        return
    
    idx = -1
    background_mode = True
    
    if choice == '1':
        idx = 0
        user_name = "Suami"
    elif choice == '2':
        idx = 1
        user_name = "Istri"
    elif choice == '3':
        idx = 0
        user_name = "Suami"
        background_mode = False
    else:
        print("❌ Pilihan tidak valid!")
        input("Tekan ENTER untuk kembali...")
        return test_menu()
    
    print()
    print(f"📌 Target: {user_name} (Emulator {idx})")
    print(f"📌 Mode: {'BACKGROUND (minimize OK)' if background_mode else 'FOREGROUND'}")
    print()
    confirm = input(f"Jalankan test? (y/n): ").strip().lower()
    
    if confirm == 'y':
        print()
        print("=" * 50)
        print(f"🚀 MEMULAI TEST: {user_name.upper()}")
        print("=" * 50)
        print()
        
        # Jalankan automation
        run_hybrid_automation(idx, background_mode=background_mode)
        
        print()
        print("=" * 50)
        print("✅ TEST SELESAI!")
        print("=" * 50)
        print()
        input("Tekan ENTER untuk keluar...")
    else:
        print("Test dibatalkan.")
        input("Tekan ENTER untuk kembali...")
        return test_menu()

if __name__ == "__main__":
    test_menu()
