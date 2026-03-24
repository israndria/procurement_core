import os
import re
import subprocess
import threading

def get_scripts(directory):
    """
    Memindai direktori untuk file .py dan mengembalikannya sebagai list.
    Mengabaikan file sistem atau file __init__.py.
    """
    if not os.path.exists(directory):
        return []
    
    scripts = []
    for f in os.listdir(directory):
        if f.endswith(".py") and f != "__init__.py" and "Mencari Koordinat" not in f:
             scripts.append(f)
    return sorted(scripts)

def format_display_name(filename):
    """
    Mengubah nama file menjadi nama tampilan yang lebih bersih.
    Contoh: "Aktivitas Govem (Jumat1).py" -> "Jumat (1)"
    """
    name = filename.replace(".py", "")
    
    # Hapus prefix umum
    name = name.replace("Aktivitas Govem", "").replace("Aktivitas Istri", "").replace("Aktivitas", "")
    
    # Bersihkan tanda kurung berlebih atau spasi
    name = name.strip()
    
    # Jika hasil kosong, kembalikan nama asli
    if not name:
        return filename.replace(".py", "")
        
    # Khusus untuk request user: Jumat1 -> Jumat
    name = name.replace("Jumat1", "Jumat")
    
    # Hapus tanda kurung di awal/akhir jika ada
    if name.startswith("(") and name.endswith(")"):
        name = name[1:-1]
        
    return name.strip()

def extract_description(file_path):
    """
    Mencoba mengekstrak deskripsi aktivitas dari perintah pyautogui.write() di dalam file.
    Mengembalikan list string deskripsi yang ditemukan.
    """
    descriptions = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Mencari pola pyautogui.write("...") atau pyautogui.write('...')
            matches = re.findall(r'pyautogui\.write\((?:"|\')(.+?)(?:"|\')\)', content)
            descriptions.extend(matches)
    except Exception as e:
        descriptions.append(f"Gagal membaca deskripsi: {e}")
    
    if not descriptions:
        return ["Tidak ada deskripsi spesifik ditemukan (tidak ada perintah pyautogui.write)."]
    return descriptions

    return descriptions

def open_chrome_profile(profile_directory):
    """
    Membuka Google Chrome dengan profil tertentu.
    profile_directory: nama folder profil (misal: "Default" atau "Profile 5")
    """
    try:
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            return False, "Chrome.exe tidak ditemukan di lokasi standar."
            
        # Command line arguments untuk membuka profil
        # --profile-directory="Profile 5"
        cmd = [chrome_path, f"--profile-directory={profile_directory}"]
        
        # Jalankan Chrome (detached process)
        subprocess.Popen(cmd)
        
        # Beri waktu sebentar agar window aktif
        import time
        time.sleep(1) # Tunggu 1 detik biar fokus pindah
        
        return True, f"Membuka Chrome profil: {profile_directory}"
    except Exception as e:
        return False, f"Gagal membuka browser: {e}"

def run_script_process(script_path):
    """
    Menjalankan script python menggunakan subprocess.
    """
    try:
        # Menggunakan python yang sama dengan environment saat ini atau 'python' global
        # Kita asumsikan 'python' ada di PATH atau gunakan sys.executable
        import sys
        python_executable = sys.executable 
        
        # Siapkan file log
        log_file = os.path.join(os.path.dirname(script_path), "debug_log.txt")
        
        # Jalankan sebagai proses terpisah dengan redirect output
        with open(log_file, "w") as f:
            subprocess.Popen(
                [python_executable, script_path], 
                cwd=os.path.dirname(script_path),
                stdout=f,
                stderr=f
            )
            
        return True, f"Skrip dijalankan. Cek {log_file} jika tidak ada reaksi."
    except Exception as e:
        return False, f"Gagal menjalankan skrip: {e}"
