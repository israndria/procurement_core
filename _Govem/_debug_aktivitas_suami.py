"""
Debug Script: Aktivitas Suami ONLY
====================================
Fokus test: V23_aktivitas_Suami mengisi 7 aktivitas dengan benar.
Emulator Suami harus sudah menyala dan app Govem sudah terbuka.

Jalankan: python -X utf8 _debug_aktivitas_suami.py
"""
import os, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import Govem_Engine as engine

SUAMI = {'name': 'Suami', 'index': 0, 'gps': True, 'days': [0], 'port': None}

def run():
    print("=" * 60)
    print("  DEBUG: AKTIVITAS SUAMI ONLY")
    print("  Launch emulator -> buka app -> isi 7 aktivitas")
    print("=" * 60)

    idx = SUAMI['index']

    # 1. Launch emulator
    print("\n[1] Launching emulator Suami...")
    engine.launch_emulator(idx)

    # 2. Buka app
    print("\n[2] Membuka app Govem...")
    if not engine.open_app(idx):
        print("❌ Gagal buka app. Abort.")
        return

    # 3. Trigger aktivitas (ini yang mau ditest)
    print("\n[3] Menjalankan V23_aktivitas_Suami...")
    engine.trigger_activity(SUAMI)

    print("\n" + "=" * 60)
    print("  SELESAI! Cek app Govem untuk verifikasi.")
    print("=" * 60)

if __name__ == "__main__":
    run()
