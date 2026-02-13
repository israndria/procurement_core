"""
TEST: ldconsole inputtext SAJA (tanpa macro)
Untuk menguji apakah ldconsole inputtext bekerja di background
CATATAN: User harus sudah klik dulu pada text field!
"""
import subprocess
import time
import os

LDCONSOLE = r"D:\LDPlayer\LDPlayer9\ldconsole.exe"
ADB = r"D:\LDPlayer\LDPlayer9\adb.exe"

def run_command(command):
    creationflags = 0x08000000
    if isinstance(command, list):
        result = subprocess.run(command, capture_output=True, text=True, shell=False, creationflags=creationflags)
    else:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, creationflags=creationflags)
    return result.stdout.strip()

def main():
    os.system('cls')
    print("=" * 60)
    print("   TEST: ldconsole inputtext (TANPA MACRO)")
    print("=" * 60)
    print()
    print("PERSIAPAN:")
    print("1. LDPlayer running dan login GOVEM")
    print("2. Buka halaman form aktivitas")
    print("3. ⚠️ KLIK pada text field (cursor HARUS aktif)")
    print("4. MINIMIZE LDPlayer")
    print()
    
    input("Tekan ENTER ketika cursor SUDAH AKTIF di text field...")
    
    test_text = "TEST_LDCONSOLE_BERHASIL_BACKGROUND"
    idx = 0  # Suami
    
    print(f"\n📤 Mengirim via ldconsole inputtext: '{test_text}'")
    print("   (LDPlayer harus tetap minimize)")
    
    time.sleep(2)  # Beri waktu user
    
    start = time.time()
    run_command([LDCONSOLE, "inputtext", "--index", str(idx), "--text", test_text])
    elapsed = time.time() - start
    
    print(f"   ⏱️ Waktu: {elapsed:.3f} detik")
    print()
    print("SELESAI! Cek LDPlayer:")
    print("  1. Apakah LDPlayer tetap minimize?")
    print("  2. Apakah teks terketik di text field?")
    print()
    
    window_ok = input("LDPlayer tetap minimize? (y/n): ").strip().lower()
    text_ok = input("Teks terketik? (y/n): ").strip().lower()
    
    if window_ok == 'y' and text_ok == 'y':
        print("\n🎉 SUKSES! ldconsole inputtext bekerja di background!")
    elif window_ok == 'y' and text_ok == 'n':
        print("\n❌ ldconsole: Window OK tapi teks tidak terketik")
    else:
        print("\n❌ Gagal")

if __name__ == "__main__":
    main()
    input("\nTekan ENTER untuk keluar...")
