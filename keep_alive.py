import os
from supabase import create_client

# Ambil rahasia dari Environment Variable GitHub
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

try:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError("URL atau Key belum disetting di GitHub Secrets!")

    # Koneksi
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    
    # Lakukan Query Ringan (Select 1 baris saja)
    response = supabase.table("tender").select("kode_tender").limit(1).execute()
    
    print("✅ Jantung berdetak! Database berhasil dicolek.")
    print(f"Data sample: {response.data}")

except Exception as e:
    print(f"❌ Gagal mencolek database: {e}")
    # Kita exit dengan error supaya GitHub memberi notifikasi email kalau gagal
    exit(1)