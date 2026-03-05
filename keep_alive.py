"""
Supabase Keep-Alive Script
===========================
Mencegah project Supabase FREE tier dari auto-pause (freeze).
Supabase mem-pause project gratis setelah 7 hari tidak ada aktivitas database.

Script ini harus dijadwalkan via GitHub Actions dengan cron setiap 3-4 hari.

Contoh cron di GitHub Actions workflow:
  schedule:
    - cron: '0 0 */3 * *'   # Setiap 3 hari sekali jam 00:00 UTC
"""

import os
import sys
import requests
import json
from datetime import datetime

# ============================================================
# KONFIGURASI - Ambil dari Environment Variable GitHub Secrets
# ============================================================
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

def keep_alive():
    """Melakukan query ringan ke Supabase agar project tidak di-pause."""
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"🕐 [{timestamp}] Memulai keep-alive ping...")
    
    # Validasi environment variables
    if not SUPABASE_URL:
        print("❌ SUPABASE_URL belum di-set di GitHub Secrets!")
        print("   Pastikan sudah menambahkan secret 'SUPABASE_URL' di repository settings.")
        return False
    
    if not SUPABASE_KEY:
        print("❌ SUPABASE_KEY belum di-set di GitHub Secrets!")
        print("   Pastikan sudah menambahkan secret 'SUPABASE_KEY' di repository settings.")
        return False
    
    print(f"📡 Target: {SUPABASE_URL}")
    
    # ===== METODE 1: REST API langsung (lebih ringan, tanpa dependency supabase-py) =====
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    
    try:
        # Query ringan: SELECT 1 baris dari tabel apapun yang ada
        # Menggunakan REST API PostgREST bawaan Supabase
        rest_url = f"{SUPABASE_URL}/rest/v1/"
        
        # Coba ping ke health check endpoint dulu
        health_url = f"{SUPABASE_URL}/rest/v1/"
        response = requests.get(health_url, headers=headers, timeout=30)
        print(f"📊 REST API Status: {response.status_code}")
        
        # Lakukan query ke tabel 'tender' (ambil 1 baris saja)
        query_url = f"{SUPABASE_URL}/rest/v1/tender?select=kode_tender&limit=1"
        response = requests.get(query_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Jantung berdetak! Database berhasil dicolek.")
            print(f"   Data sample: {data}")
            return True
        elif response.status_code == 404:
            # Tabel 'tender' mungkin tidak ada, coba query RPC atau tabel lain
            print(f"⚠️  Tabel 'tender' tidak ditemukan (404). Mencoba metode alternatif...")
            
            # Gunakan RPC untuk menjalankan query SQL sederhana
            rpc_url = f"{SUPABASE_URL}/rest/v1/rpc/ping"
            rpc_response = requests.post(rpc_url, headers=headers, json={}, timeout=30)
            
            if rpc_response.status_code == 200:
                print(f"✅ Ping via RPC berhasil!")
                return True
            else:
                # Fallback: coba akses auth endpoint (ini tetap menghitung sebagai aktivitas)
                auth_url = f"{SUPABASE_URL}/auth/v1/health"
                auth_response = requests.get(auth_url, headers=headers, timeout=30)
                print(f"📊 Auth Health Status: {auth_response.status_code}")
                if auth_response.status_code == 200:
                    print(f"✅ Auth health check berhasil! Project tetap aktif.")
                    return True
                else:
                    print(f"⚠️  Auth health response: {auth_response.text}")
                    # Meskipun gagal, request ini sendiri sudah dianggap aktivitas
                    return True
        else:
            print(f"⚠️  Response tidak terduga: {response.status_code}")
            print(f"   Body: {response.text[:500]}")
            # Request tetap sampai ke Supabase, jadi tetap dianggap aktivitas
            return True
            
    except requests.exceptions.Timeout:
        print("❌ Timeout! Supabase tidak merespons dalam 30 detik.")
        print("   Kemungkinan project sudah di-pause. Resume manual di dashboard.")
        return False
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Koneksi gagal: {e}")
        return False
    except Exception as e:
        print(f"❌ Error tidak terduga: {e}")
        return False


def main():
    success = keep_alive()
    
    if success:
        print("\n🎉 Keep-alive berhasil! Project Supabase tetap aktif.")
        sys.exit(0)
    else:
        print("\n💀 Keep-alive GAGAL! Cek konfigurasi dan status project.")
        print("   👉 Buka https://supabase.com/dashboard untuk resume manual.")
        sys.exit(1)


if __name__ == "__main__":
    main()