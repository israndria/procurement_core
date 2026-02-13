import Govem_Engine as engine
import sys
import time
import os

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("====================================")
    print("   🤖 GOVEM AUTO: SYSTEM ISTRI")
    print("====================================")
    
    # Cek Argument --auto
    if len(sys.argv) > 1 and sys.argv[1] == "--auto":
        print("\n🚀 MODE AUTO-STARTUP TERDETEKSI!")
        print("Langsung menjalankan Scheduler Istri...")
        time.sleep(3)
        istri = engine.USERS[1]
        engine.run_scheduler([istri])
        return

    while True:
        print("\nMenu Istri:")
        print("1. ☀️ Manual Absen Pagi")
        print("2. 🌙 Manual Absen Sore")
        print("3. 📍 Set Lokasi Kantor")
        print("4. 🕒 Start Scheduler (Otomatis)")
        print("5. 🔧 Tools (Kalibrasi / Import)")
        print("6. 🚑 Diagnosa System")
        print("0. Keluar")
        
        choice = input("Pilihan: ")
        istri = engine.USERS[1]
        
        if choice == '1':
            engine.absen_pagi(istri)
        elif choice == '2':
            engine.absen_sore(istri)
        elif choice == '3':
            print("\n--- SETTING LOKASI ISTRI ---")
            lat = input("Masukkan Latitude: ")
            lng = input("Masukkan Longitude: ")
            engine.save_config("LOCATION_1", "latitude", lat)
            engine.save_config("LOCATION_1", "longitude", lng)
            print("✅ Tersimpan.")
            
        elif choice == '4':
            engine.run_scheduler([istri])
        elif choice == '5':
            engine.calibration_mode() # Shared tools
        elif choice == '6':
            engine.run_diagnostics()
        elif choice == '0':
            break

if __name__ == "__main__":
    main()
