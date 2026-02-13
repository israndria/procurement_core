import os
import subprocess
import time

LDPLAYER_PATH = r"D:\LDPlayer\LDPlayer9"
ADB = os.path.join(LDPLAYER_PATH, "adb.exe")

def run(cmd):
    print(f"Exec: {cmd}")
    subprocess.run(cmd, shell=True)

print("--- TEST ADB CONNECTION ---")
run(f'"{ADB}" devices')

print("\n--- TEST INPUT TEXT (Pastikan Kursor di Kolom Teks) ---")
print("Akan mengetik 'halo' dalam 5 detik...")
time.sleep(5)
run(f'"{ADB}" -s emulator-5554 shell input text "halo"')
run(f'"{ADB}" -s 127.0.0.1:5554 shell input text "halo_ip"')

print("\nDone.")
