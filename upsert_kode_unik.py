"""
upsert_kode_unik.py — Dipanggil dari VBA ModKodeUnik setelah generate kode unik.
Usage: python upsert_kode_unik.py <kode_tender> <kode_unik>
"""
import sys
import os
import pathlib

def main():
    if len(sys.argv) < 3:
        print("Usage: upsert_kode_unik.py <kode_tender> <kode_unik>")
        sys.exit(1)

    kode_tender = sys.argv[1].strip()
    kode_unik   = sys.argv[2].strip()

    if not kode_tender or not kode_unik:
        sys.exit(0)

    # Tambah Asisten_Pokja ke path untuk config
    base = pathlib.Path(__file__).parent
    asisten_dir = base.parent.parent / "Asisten_Pokja"
    if asisten_dir.exists():
        sys.path.insert(0, str(asisten_dir))
    else:
        # fallback: satu level di atas WPy64
        asisten_dir = base.parent / "Asisten_Pokja"
        sys.path.insert(0, str(asisten_dir))

    try:
        from config import sb
        client = sb()
        client.table("draft_paket").update({"kode_unik": kode_unik}).eq("kode_tender", kode_tender).execute()
        print(f"OK: kode_unik '{kode_unik}' disimpan untuk {kode_tender}")
    except Exception as e:
        print(f"Error upsert Supabase: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
