import pyautogui
import keyboard
import time

print("Arahkan kursor dan tekan tombol 's' untuk menyimpan koordinat. Tekan 'esc' untuk keluar.")

saved_coords = []

while True:
    if keyboard.is_pressed('s'):
        x, y = pyautogui.position()
        saved_coords.append((x, y))
        print(f"Disimpan: (x={x}, y={y})")
        time.sleep(0.5)  # Delay agar tidak mencatat dobel

    elif keyboard.is_pressed('esc'):
        print("Selesai. Semua koordinat yang disimpan:")
        for coord in saved_coords:
            print(coord)
        break