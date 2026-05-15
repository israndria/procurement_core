"""
upsert_kode_unik_pl.py — Dipanggil dari VBA ModKodeUnikPL setelah generate kode unik PL.
Usage: python upsert_kode_unik_pl.py <kode_paket> <kode_unik>
"""
import sys
import pathlib


def main():
    if len(sys.argv) < 3:
        print("Usage: upsert_kode_unik_pl.py <kode_paket> <kode_unik>")
        sys.exit(1)

    kode_paket = sys.argv[1].strip()
    kode_unik  = sys.argv[2].strip()

    if not kode_paket or not kode_unik:
        sys.exit(0)

    base = pathlib.Path(__file__).parent
    asisten_dir = base.parent.parent / "Asisten_Pokja"
    if not asisten_dir.exists():
        asisten_dir = base.parent / "Asisten_Pokja"
    sys.path.insert(0, str(asisten_dir))

    try:
        from config import sb
        client = sb()
        client.table("draft_paket_pl").update({"kode_unik": kode_unik}).eq("kode_paket", kode_paket).execute()
        print(f"OK: kode_unik '{kode_unik}' disimpan untuk {kode_paket}")
    except Exception as e:
        print(f"Error upsert Supabase: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
