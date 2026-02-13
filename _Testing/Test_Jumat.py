"""
TEST KHUSUS: Aktivitas JUMAT (Override Hari)
Untuk testing aktivitas Jumat di hari lain
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from V23_aktivitas_Suami import (
    connect_adb_smart, play_record_file, adb_input_text, 
    run_command, ADB, get_suami_activities, _CACHED_SERIAL
)
import time

def test_jumat():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("=" * 50)
    print("   🧪 TEST AKTIVITAS JUMAT (Suami)")
    print("=" * 50)
    print()
    print("⚠️ PENTING:")
    print("   1. LDPlayer sudah running (boleh minimize)")
    print("   2. Sudah LOGIN ke GOVEM")
    print("   3. Berada di HALAMAN AWAL (Dashboard)")
    print()
    print("📅 Override: Hari = JUMAT (day_idx = 4)")
    print("📋 Aktivitas: 6 item (senam pagi + 5 telaah)")
    print("📋 Step 3: Step 3.2 (SKP aktivitas 6)")
    print()
    
    confirm = input("Jalankan test Jumat? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Test dibatalkan.")
        return
    
    idx = 0  # Suami
    
    print("\n🔗 Connecting...")
    serial = connect_adb_smart(idx, launch_if_needed=False)
    _CACHED_SERIAL[idx] = serial
    
    # Force Jumat (day_idx = 4)
    result = get_suami_activities(4)  # 4 = Jumat
    activity_texts = result[0]
    is_jumat = result[1]
    
    print(f"\n🚀 MEMULAI PENGISIAN {len(activity_texts)} AKTIVITAS")
    print("📅 Mode JUMAT: Step 3.2 akan digunakan")
    print("💡 LDPlayer bisa di-minimize!")
    
    # STEP 1
    print("\n[STEP 1] Navigasi ke Form...")
    play_record_file(idx, "Step 1 (Menuju Dashboard Isian)", serial)
    time.sleep(3)
    
    for i, text in enumerate(activity_texts):
        print(f"\n📝 [{i+1}/{len(activity_texts)}] Mengisi: {text[:30]}...")
        
        # Step 2
        play_record_file(idx, "Step 2 (Klik untuk supaya bisa mengetikcopas kata2)", serial)
        time.sleep(0.8)
        
        # Input text (ADB - proven bekerja di background)
        print(f"   ⌨️ Mengetik...")
        clean_text = text.replace("'", "").replace('"', "")
        adb_input_text(serial, clean_text, idx)
        time.sleep(0.3)
        
        # Hide keyboard
        run_command([ADB, "-s", serial, "shell", "input", "keyevent", "111"])
        time.sleep(0.3)
        
        # Step 3: Senam pagi = Step 3.2 (buka dropdown) + scroll + klik aktivitas 6
        text_lower = text.lower()
        is_senam = "senam pagi" in text_lower
        
        if is_senam:
            print("   📋 Membuka dropdown SKP (Step 3.2)...")
            play_record_file(idx, "Step 3.2 (Membuka dropdown SKP dan memilih aktivitas nomor 6)", serial)
            time.sleep(1)
            
            # Scroll down dalam dropdown
            print("   📜 Scroll dropdown...")
            run_command([ADB, "-s", serial, "shell", "input", "swipe", "400", "700", "400", "400", "300"])
            time.sleep(0.8)
            
            # === DEBUG: Screenshot SEBELUM klik dengan penanda ===
            click_x, click_y = 400, 790
            debug_path = r"D:\debug_klik_aktivitas6.png"
            
            print(f"   📸 Screenshot dengan penanda di ({click_x}, {click_y})...")
            try:
                import subprocess
                # Ambil screenshot langsung ke stdout dan simpan
                adb_path = r"D:\LDPlayer\LDPlayer9\adb.exe"
                result = subprocess.run(
                    [adb_path, "-s", serial, "exec-out", "screencap", "-p"],
                    capture_output=True
                )
                if result.returncode == 0 and len(result.stdout) > 0:
                    with open(debug_path, 'wb') as f:
                        f.write(result.stdout)
                    
                    # Gambar lingkaran merah di koordinat klik
                    from PIL import Image, ImageDraw
                    img = Image.open(debug_path)
                    draw = ImageDraw.Draw(img)
                    # Lingkaran merah dengan radius 30
                    r = 30
                    draw.ellipse([click_x-r, click_y-r, click_x+r, click_y+r], outline="red", width=5)
                    # Cross/silang
                    draw.line([click_x-r, click_y, click_x+r, click_y], fill="red", width=3)
                    draw.line([click_x, click_y-r, click_x, click_y+r], fill="red", width=3)
                    # Teks koordinat
                    draw.text((click_x+35, click_y-10), f"KLIK ({click_x},{click_y})", fill="red")
                    img.save(debug_path)
                    print(f"   ✅ Screenshot: {debug_path}")
                else:
                    print(f"   ⚠️ Screenshot gagal")
            except Exception as e:
                print(f"   ⚠️ Error screenshot: {e}")
            
            # Klik aktivitas nomor 6
            print(f"   👆 Klik aktivitas 6 di ({click_x}, {click_y})...")
            run_command([ADB, "-s", serial, "shell", "input", "tap", str(click_x), str(click_y)])
            time.sleep(1)
        else:
            print("   📋 Memilih SKP (Step 3.1 - Umum)...")
            play_record_file(idx, "Step 3.1 (Membuka dropdown SKP dan memilih aktivitas nomor 4)", serial)
            time.sleep(1.5)
        
        # Step 4: Senam dan aktivitas umum = Step 4 (Jenis 1)
        # Tidak ada apel di Jumat, jadi semua pakai Step 4
        print("   📄 Mode Umum/Senam -> Step 4 (Jenis 1)")
        play_record_file(idx, "Step 4 (Memilih Jenis Aktivitas Nomor 1)", serial)
        time.sleep(0.8)
        
        # Step 5
        play_record_file(idx, "Step 5 (Posting Aktivitas)", serial)
        print("   💾 Simpan...")
        time.sleep(2)
    
    print("\n✅ TEST JUMAT SELESAI!")
    input("\nTekan ENTER untuk keluar...")

if __name__ == "__main__":
    test_jumat()
