"""
Debug Script: Sesi PAGI Hari Senin
===================================
Flow: Pancingan -> Suami (absen) -> Istri (absen + aktivitas)
Test: auto-minimize instan, pancingan deferred kill via callback,
      screenshot ke folder timestamps, batch notifikasi

Jalankan: python -X utf8 _debug_pagi_senin.py
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


def kill_all():
    """Matikan semua emulator sebelum mulai."""
    print("\n[PREP] Mematikan semua emulator...")
    for idx in [0, 1, 2]:
        engine.run_command(f'"{engine.LDCONSOLE}" quit --index {idx}')
    time.sleep(3)
    print("[PREP] Semua emulator mati.\n")


def run_pagi():
    print("=" * 60)
    print("  DEBUG: SESI PAGI SENIN")
    print("  Pancingan -> Suami -> Istri (+ aktivitas)")
    print("  Callback kill Pancingan + screenshot folder + notifikasi")
    print("=" * 60)

    kill_all()

    sorted_users = sorted(USERS_SENIN,
                          key=lambda u: LAUNCH_ORDER.index(u['name']) if u['name'] in LAUNCH_ORDER else 99)

    screenshots = []
    pancingan_index = None

    def _kill_pancingan_callback():
        """Callback: kill Pancingan 3s setelah emulator berikutnya boot."""
        nonlocal pancingan_index
        if pancingan_index is not None:
            time.sleep(3)
            print(f"  [Pancingan] Auto-kill via callback (iklan sudah terserap)")
            engine.run_command(f'"{engine.LDCONSOLE}" quit --index {pancingan_index}')
            pancingan_index = None

    for u in sorted_users:
        name = u['name']
        print(f"\n{'='*40}")
        print(f"  Memulai: {name}")
        print(f"{'='*40}")

        try:
            # Inject callback ke user berikutnya setelah Pancingan
            if name != 'Pancingan' and pancingan_index is not None:
                u['_on_boot_callback'] = _kill_pancingan_callback

            engine.absen_pagi(u)
            print(f"  [OK] {name} selesai absen pagi")

            if name == 'Pancingan':
                pancingan_index = u['index']
                print(f"  [Pancingan] Index {pancingan_index} disimpan, menunggu callback kill...")
            else:
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
        engine._send_batch_notification("Absen Pagi", screenshots)
    else:
        print("\n[WARNING] Tidak ada screenshot yang berhasil diambil.")

    print("\n" + "=" * 60)
    print("  SESI PAGI SELESAI!")
    print("=" * 60)


if __name__ == "__main__":
    run_pagi()
