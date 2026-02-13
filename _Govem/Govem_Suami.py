import Govem_Engine as engine
import sys
import time
import os

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("====================================")
    print("   🤖 GOVEM AUTO: SYSTEM SUAMI")
    print("====================================")
    
    # Cek Argument --auto
    if len(sys.argv) > 1 and sys.argv[1] == "--auto":
        print("\n🚀 MODE AUTO-STARTUP TERDETEKSI!")
        print("Langsung menjalankan Scheduler Suami...")
        time.sleep(3)
        suami = engine.USERS[0]
        engine.run_scheduler([suami])
        return

    while True:
        print("\nMenu Suami:")
        print("1. ☀️ Manual Absen Pagi")
        print("2. 🌙 Manual Absen Sore")
        print("3. 📍 Set Lokasi Kantor")
        print("4. 🕒 Start Scheduler (Otomatis)")
        print("5. 🔧 Tools (Kalibrasi / Import)")
        print("6. 🚑 Diagnosa System")
        print("0. Keluar")
        
        choice = input("Pilihan: ")
        suami = engine.USERS[0]
        
        if choice == '1':
            engine.absen_pagi(suami)
        elif choice == '2':
            engine.absen_sore(suami)
        elif choice == '3':
            engine.set_location(suami) # Will ask for input handled in engine or wrapper? Engine set_location asks input.
            # Engine set_location actually handles input internally if not hardcoded.
            # Let's check Govem_Engine.set_location logic. 
            # It uses "LOCATION_{idx}". It has logic to read user input if manual? 
            # Actually engine.calibration_mode has menu 7 for location.
            # Let's call engine.calibration_mode() -> Option 7? Or create wrapper?
            # engine.set_location(user) reads from config and applies. It does NOT ask input.
            # We need a wizard to SET it.
            print("\n--- SETTING LOKASI SUAMI ---")
            lat = input("Masukkan Latitude: ")
            lng = input("Masukkan Longitude: ")
            engine.save_config("LOCATION_0", "latitude", lat)
            engine.save_config("LOCATION_0", "longitude", lng)
            print("✅ Tersimpan.")
            
        elif choice == '4':
            engine.run_scheduler([suami])
        elif choice == '5':
            engine.calibration_mode() # Shared tools
        elif choice == '6':
            engine.run_diagnostics()
        elif choice == '0':
            break

if __name__ == "__main__":
    main()
