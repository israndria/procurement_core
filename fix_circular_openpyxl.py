"""Fix circular reference langsung via openpyxl — tanpa buka Excel."""
import openpyxl, shutil, os

filepath = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.4.xlsm"
backup  = filepath + ".bak"

# Backup dulu
shutil.copy2(filepath, backup)
print(f"Backup: {backup}")

wb = openpyxl.load_workbook(filepath, keep_vba=True)

# Fix 1: Sheet 6 W29 — hapus formula circular → value 0
ws6 = wb["6. BA KLARIF SKP ALAT"]
ws6["W29"] = 0
print(f"Fix W29: {ws6['W29'].value}")

# Fix 2: "0. Input BA" C27/C28/D27/E27 — hapus formula circular → kosong
ws0 = wb["0. Input BA"]
for coord in ["C27", "D27", "E27", "C28"]:
    ws0[coord] = ""
    print(f"Fix {coord}: dikosongkan")

wb.save(filepath)
print("Saved — buka Excel sekarang, tidak ada circular reference lagi.")
