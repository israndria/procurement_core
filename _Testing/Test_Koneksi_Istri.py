"""
TEST DIAGNOSA: Koneksi LDPlayer Istri
"""
import subprocess
import time
import os

ADB = r"D:\LDPlayer\LDPlayer9\adb.exe"
LDCONSOLE = r"D:\LDPlayer\LDPlayer9\ldconsole.exe"

def run_command(command):
    creationflags = 0x08000000
    if isinstance(command, list):
        result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
    else:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
    return result.stdout.strip()

def main():
    os.system('cls')
    print("="*60)
    print("   DIAGNOSA: Koneksi LDPlayer Istri")
    print("="*60)
    
    # 1. List semua emulator
    print("\n📋 1. Daftar semua emulator LDPlayer:")
    list_result = run_command([LDCONSOLE, "list2"])
    print(f"   {list_result}")
    
    # 2. Connect ke semua port
    print("\n🔌 2. Connect ADB ke semua port...")
    ports = [5555, 5556, 5557, 5558, 5559, 5560]
    for port in ports:
        result = run_command([ADB, "connect", f"127.0.0.1:{port}"])
        if "connected" in result.lower() or "already" in result.lower():
            print(f"   ✅ Port {port}: {result}")
        else:
            print(f"   ❌ Port {port}: {result}")
    
    time.sleep(1)
    
    # 3. List connected devices
    print("\n📱 3. Devices terhubung:")
    devices_result = run_command([ADB, "devices"])
    print(f"   {devices_result}")
    
    lines = devices_result.split('\n')
    serials = []
    for line in lines:
        if "127.0.0.1" in line and "device" in line:
            serial = line.split()[0]
            serials.append(serial)
            print(f"   ✅ Found: {serial}")
    
    # 4. Tentukan serial Istri
    print("\n🎯 4. Menentukan serial untuk Istri (Index 1):")
    if len(serials) > 1:
        istri_serial = serials[1]
        print(f"   ✅ Istri serial: {istri_serial}")
    elif len(serials) == 1:
        istri_serial = serials[0]
        print(f"   ⚠️ Hanya 1 device, pakai: {istri_serial}")
    else:
        print("   ❌ Tidak ada device!")
        return
    
    # 5. Test tap sederhana
    print(f"\n🖱️ 5. Test tap di {istri_serial}...")
    input("   Pastikan LDPlayer Istri sudah running. Tekan ENTER...")
    
    print("   Tap ke (400, 400)...")
    run_command([ADB, "-s", istri_serial, "shell", "input", "tap", "400", "400"])
    print("   ✅ Tap selesai!")
    
    print("\n📋 Apakah tap berhasil di LDPlayer Istri?")
    print("   Jika ya, berarti koneksi OK.")
    print("   Jika tidak, mungkin serial salah atau emulator tidak responsif.")

if __name__ == "__main__":
    main()
    input("\nTekan ENTER untuk keluar...")
