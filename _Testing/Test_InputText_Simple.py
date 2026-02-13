"""
TEST SEDERHANA: Sama persis dengan cara V23_Aktivitas_Auto mengirim text
"""
import subprocess
import time
import os

ADB = r"D:\LDPlayer\LDPlayer9\adb.exe"

def run_command(command):
    creationflags = 0x08000000  # CREATE_NO_WINDOW
    if isinstance(command, list):
        result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
    else:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
    return result.stdout.strip()

def connect_and_get_serial():
    # Connect ke ports
    ports = [5555, 5556, 5557, 5558, 5559, 5560]
    for port in ports:
        run_command([ADB, "connect", f"127.0.0.1:{port}"])
    
    time.sleep(0.5)
    
    result = run_command([ADB, "devices"])
    for line in result.split('\n'):
        if "127.0.0.1" in line and "device" in line:
            return line.split()[0]
    return None

def adb_input_text(serial, text):
    """Sama persis dengan V23_Aktivitas_Auto.adb_input_text"""
    escaped_text = text.replace(" ", "%s")
    escaped_text = escaped_text.replace("'", "")
    escaped_text = escaped_text.replace('"', "")
    escaped_text = escaped_text.replace("&", "")
    escaped_text = escaped_text.replace("(", "")
    escaped_text = escaped_text.replace(")", "")
    escaped_text = escaped_text.replace(";", "")
    
    run_command([ADB, "-s", serial, "shell", "input", "text", escaped_text])

def main():
    os.system('cls')
    print("=" * 60)
    print("   TEST SEDERHANA: Input Text sama seperti V23")
    print("=" * 60)
    print()
    
    print("PERSIAPAN:")
    print("1. LDPlayer running dan login GOVEM")
    print("2. Buka halaman form aktivitas")
    print("3. KLIK text field (cursor aktif)")
    print("4. MINIMIZE LDPlayer")
    print()
    
    input("Tekan ENTER ketika siap...")
    
    serial = connect_and_get_serial()
    if not serial:
        print("❌ Tidak ada device!")
        return
    
    print(f"\n✅ Serial: {serial}")
    
    # Test text
    test_text = "Menelaah dokumen dan berkas pengadaan barang jasa"
    
    print(f"\n📤 Mengirim: '{test_text}'")
    print("   (Tunggu 10 detik)")
    
    time.sleep(3)  # Kasih waktu user minimize
    
    adb_input_text(serial, test_text)
    
    print("\n✅ Selesai mengirim!")
    print()
    print("Buka LDPlayer dan cek:")
    print("  - Apakah teks terketik di text field?")
    print()
    
    result = input("Teks terketik? (y/n): ").strip().lower()
    if result == 'y':
        print("\n🎉 SUKSES! ADB input text bekerja di background!")
    else:
        print("\n❓ Gagal - mungkin cursor tidak aktif atau minimize terlalu cepat")

if __name__ == "__main__":
    main()
    input("\nTekan ENTER untuk keluar...")
