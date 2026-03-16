"""
Debug Script: Sesi SORE Hari Senin
===================================
Flow: Pancingan -> Suami (absen + aktivitas) -> Istri (absen only)
Test: auto-minimize, pancingan deferred kill, aktivitas Suami, dashboard screenshot, batch notifikasi

PENTING: Jalankan SETELAH _debug_pagi_senin.py selesai dan hasilnya OK.
         Emulator Suami & Istri harus masih menyala dari sesi pagi.

Jalankan: python -X utf8 _debug_sore_senin.py
"""
import os, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import Govem_Engine as engine

USERS_SENIN = [
    {'name': 'Pancingan', 'index': 2, 'gps': False, 'days': [0], 'port': None},
    {'name': 'Suami', 'index': 0, 'gps': True, 'days': [0], 'port': None},
    {'name': 'Istri', 'index': 1, 'gps': True, 'days': [0], 'port': None},
]

LAUNCH_ORDER = ['Pancingan', 'Suami', 'Istri']


def run_sore():
    print("=" * 60)
    print("  DEBUG: SESI SORE SENIN")
    print("  Pancingan -> Suami (+ aktivitas) -> Istri (absen only)")
    print("  Pancingan deferred kill + batch screenshot + notifikasi")
    print("=" * 60)

    sorted_users = sorted(USERS_SENIN,
                          key=lambda u: LAUNCH_ORDER.index(u['name']) if u['name'] in LAUNCH_ORDER else 99)

    screenshots = []
    pancingan_index = None

    for u in sorted_users:
        name = u['name']
        print(f"\n{'='*40}")
        print(f"  Memulai: {name}")
        print(f"{'='*40}")

        try:
            engine.absen_sore(u)
            print(f"  [OK] {name} selesai absen sore")

            if name == 'Pancingan':
                pancingan_index = u['index']
                print(f"  [Pancingan] Menunggu emulator berikutnya boot sebelum kill...")
            else:
                # Kill Pancingan setelah emulator berikutnya selesai boot
                if pancingan_index is not None:
                    time.sleep(3)
                    print(f"  [Pancingan] Auto-kill (iklan sudah terserap)")
                    engine.run_command(f'"{engine.LDCONSOLE}" quit --index {pancingan_index}')
                    pancingan_index = None

                ss = engine._take_dashboard_screenshot(u)
                if ss:
                    screenshots.append(ss)
                    print(f"  [{name}] Screenshot: {ss}")

        except Exception as e:
            print(f"  [ERROR] {name} GAGAL: {e}")
            import traceback
            traceback.print_exc()

    # Safety: kill Pancingan jika belum
    if pancingan_index is not None:
        print(f"\n  [Pancingan] Auto-kill (cleanup)")
        engine.run_command(f'"{engine.LDCONSOLE}" quit --index {pancingan_index}')

    # Batch notification
    if screenshots:
        print(f"\n[NOTIFIKASI] Mengirim batch notification ({len(screenshots)} screenshot)...")
        engine._send_batch_notification("Absen Sore", screenshots)
    else:
        print("\n[WARNING] Tidak ada screenshot yang berhasil diambil.")

    print("\n" + "=" * 60)
    print("  SESI SORE SELESAI!")
    print("=" * 60)


if __name__ == "__main__":
    run_sore()
