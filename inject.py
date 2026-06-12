"""
inject.py — Entry point tunggal untuk inject VBA ke semua tipe workbook POKJA.

Auto-detect tipe workbook dari nama file, dispatch ke script yang tepat:
  - BAPLJKK / BAPLPK  → inject_pl.py  (Pengadaan Langsung JKK/PK)
  - lainnya            → inject_buttons.py (Tender PK)

Usage:
    python inject.py "path/to/file.xlsm"          # inject satu file
    python inject.py                               # inject semua xlsm di POKJA root
    python inject.py --type pl                     # paksa mode PL
    python inject.py --type tender                 # paksa mode Tender
    python inject.py --list                        # tampilkan file yang akan di-inject tanpa eksekusi

Kenapa satu entry point:
    inject_buttons.py = untuk Tender PK
    inject_pl.py      = untuk PL JKK/PK
    Salah pilih → tombol salah, modul salah, bug sulit dilacak.
    inject.py eliminasi pilihan — selalu auto-detect.
"""
import os
import sys
import glob
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()


def detect_type(filepath: str) -> str:
    """Detect tipe workbook dari nama file. Return 'pl' atau 'tender'."""
    fname = os.path.basename(filepath).upper()
    if "BAPLJKK" in fname or "BAPLPK" in fname:
        return "pl"
    return "tender"


def inject_one(filepath: str, force_type: str = None, dry_run: bool = False) -> bool:
    """Inject satu file. Return True jika sukses."""
    filepath = os.path.abspath(filepath)

    if not os.path.isfile(filepath):
        print(f"[ERROR] File tidak ditemukan: {filepath}")
        return False

    fname = os.path.basename(filepath)
    if fname.startswith("~$"):
        print(f"[SKIP] Temp file: {fname}")
        return True

    tipe = force_type or detect_type(filepath)

    if dry_run:
        print(f"  [{tipe.upper()}] {fname}")
        return True

    print(f"\n{'='*60}")
    print(f"  Tipe : {tipe.upper()}")
    print(f"  File : {fname}")
    print(f"{'='*60}")

    if tipe == "pl":
        from inject_pl import inject_pl
        inject_pl(filepath)
    else:
        from inject_buttons import inject_buttons
        inject_buttons(filepath)

    return True


def find_all_xlsm() -> list:
    """Cari semua .xlsm di POKJA root, skip temp + V19_Scheduler."""
    try:
        from config import POKJA_ROOT
    except ImportError:
        POKJA_ROOT = str(SCRIPT_DIR.parent.parent)

    pattern = os.path.join(POKJA_ROOT, "**", "*.xlsm")
    files = glob.glob(pattern, recursive=True)
    return [
        f for f in files
        if not os.path.basename(f).startswith("~$")
        and "V19_Scheduler" not in f
        and "WPy64" not in f
        and "Backup" not in os.path.basename(f)
    ]


def main():
    args = sys.argv[1:]

    # Parse flags
    force_type = None
    dry_run = False
    list_only = False
    files_explicit = []

    i = 0
    while i < len(args):
        a = args[i]
        if a == "--type" and i + 1 < len(args):
            force_type = args[i + 1].lower()
            if force_type not in ("pl", "tender"):
                print(f"[ERROR] --type harus 'pl' atau 'tender', bukan '{force_type}'")
                sys.exit(1)
            i += 2
        elif a == "--list":
            list_only = True
            i += 1
        elif a == "--dry-run":
            dry_run = True
            i += 1
        elif a.endswith(".xlsm"):
            files_explicit.append(a)
            i += 1
        else:
            # Coba sebagai path file meskipun tidak berakhiran .xlsm
            files_explicit.append(a)
            i += 1

    # Tentukan target
    if files_explicit:
        targets = files_explicit
    else:
        targets = find_all_xlsm()

    if not targets:
        print("Tidak ada file .xlsm ditemukan.")
        sys.exit(0)

    # List mode
    if list_only or dry_run:
        print(f"\nFile yang akan di-inject ({len(targets)} file):\n")
        for f in targets:
            tipe = force_type or detect_type(f)
            fname = os.path.basename(f)
            rel = os.path.relpath(f, str(SCRIPT_DIR.parent.parent)) if not files_explicit else fname
            print(f"  [{tipe.upper():6}] {rel}")
        if list_only:
            return

    # Inject
    ok, fail = 0, 0
    for f in targets:
        try:
            inject_one(f, force_type=force_type, dry_run=dry_run)
            ok += 1
        except Exception as e:
            print(f"[ERROR] {os.path.basename(f)}: {e}")
            fail += 1

    if not dry_run:
        print(f"\n{'='*60}")
        print(f"  Selesai: {ok} OK, {fail} GAGAL")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
