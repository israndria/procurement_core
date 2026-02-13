import sys
import os
import time

# Tambahkan path ke _Govem agar bisa import module engine
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from Govem_Engine import USERS, absen_pagi, absen_sore

def main():
    print("==========================================")
    print("   🚑 MODE DARURAT / MANUAL TRIGGER 🚑")
    print("==========================================")
    print("Script ini memaksa fungsi absensi jalan TANPA scheduler.")
    print("Pastikan LDPlayer sudah siap (atau biarkan script launch).")
    print("")
    print("PILIH MENU:")
    print("1. Absen PAGI - SUAMI")
    print("2. Absen PAGI - ISTRI")
    print("3. Absen PAGI - KEDUANYA (Suami lalu Istri)")
    print("5. Absen SORE - ISTRI")
    print("6. Aktivitas Saja - SUAMI (Senin-Jumat)")
    print("7. Aktivitas Saja - ISTRI (Senin-Sabtu)")
    print("0. Batal")
    
    choice = input("\nMasukkan pilihan (0-7): ").strip()
    
    if choice == '1':
        user = USERS[0] # Suami
        print(f"\n🚀 Menjalankan Absen PAGI untuk {user['name']}...")
        absen_pagi(user)
        
    elif choice == '2':
        user = USERS[1] # Istri
        print(f"\n🚀 Menjalankan Absen PAGI untuk {user['name']}...")
        absen_pagi(user)
        
    elif choice == '3':
        u_suami = USERS[0]
        u_istri = USERS[1]
        print(f"\n🚀 Menjalankan Absen PAGI untuk {u_suami['name']}...")
        absen_pagi(u_suami)
        print("\n⏳ Jeda 5 detik...")
        time.sleep(5)
        print(f"\n🚀 Menjalankan Absen PAGI untuk {u_istri['name']}...")
        absen_pagi(u_istri)
        
    elif choice == '4':
        user = USERS[0] # Suami
        print(f"\n🚀 Menjalankan Absen SORE untuk {user['name']}...")
        absen_sore(user)
        
    elif choice == '5':
        user = USERS[1] # Istri
        print(f"\n🚀 Menjalankan Absen SORE untuk {user['name']}...")
        absen_sore(user)

    elif choice == '6':
        from Govem_Engine import trigger_activity
        user = USERS[0]
        print(f"\n🚀 Menjalankan Aktivitas Harian untuk {user['name']}...")
        trigger_activity(user)

    elif choice == '7':
        from Govem_Engine import trigger_activity_istri
        user = USERS[1]
        print(f"\n🚀 Menjalankan Aktivitas Harian untuk {user['name']}...")
        trigger_activity_istri(user)
        
    elif choice == '0':
        print("Batal.")
    else:
        print("Pilihan tidak valid.")

    print("\n✅ Selesai.")
    time.sleep(5)

if __name__ == "__main__":
    main()
