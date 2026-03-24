"""
INJECT ALL — Batch inject VBA ke semua file .xlsm di bawah @ POKJA 2026
========================================================================
Scan rekursif, skip file yang sudah punya ModWordLink versi terbaru.

Cara pakai:
  python inject_all.py           # scan semua, tanya konfirmasi
  python inject_all.py --force   # inject semua tanpa tanya
"""
import os
import sys
import glob

from config import POKJA_ROOT
from inject_buttons import inject_buttons


def find_all_xlsm():
    """Cari semua .xlsm di bawah POKJA_ROOT (rekursif), skip ~$ temp files."""
    pattern = os.path.join(POKJA_ROOT, "**", "*.xlsm")
    files = glob.glob(pattern, recursive=True)
    # Filter: skip temp files dan file di V19_Scheduler (bukan paket)
    return [
        f for f in files
        if not os.path.basename(f).startswith("~$")
        and "V19_Scheduler" not in f
    ]


def main():
    force = "--force" in sys.argv

    files = find_all_xlsm()

    if not files:
        print("Tidak ada file .xlsm ditemukan.")
        return

    print(f"Ditemukan {len(files)} file .xlsm:\n")
    for i, f in enumerate(files, 1):
        rel = os.path.relpath(f, POKJA_ROOT)
        print(f"  {i}. {rel}")

    if not force:
        print()
        jawab = input("Inject VBA ke semua file di atas? (y/n): ").strip().lower()
        if jawab != "y":
            print("[BATAL]")
            return

    print("\n" + "=" * 60)
    success = 0
    fail = 0
    for f in files:
        rel = os.path.relpath(f, POKJA_ROOT)
        print(f"\n{'='*60}")
        print(f"[{success + fail + 1}/{len(files)}] {rel}")
        print(f"{'='*60}")
        try:
            inject_buttons(f)
            success += 1
        except Exception as e:
            print(f"[ERROR] {e}")
            fail += 1

    print(f"\n{'='*60}")
    print(f"  SELESAI: {success} berhasil, {fail} gagal dari {len(files)} file")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
