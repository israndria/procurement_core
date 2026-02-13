import json
import os
import glob

PATH = r"D:\LDPlayer\LDPlayer9\vms\operationRecords"

def parse(filename):
    fpath = os.path.join(PATH, filename)
    if not os.path.exists(fpath): return "File Not Found"
    
    with open(fpath, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    ops = data.get('operations', [])
    for op in ops:
         if op.get('operationId') == 'PutMultiTouch':
             points = op.get('points', [])
             if points and points[0].get('state') == 1:
                 return f"X={points[0].get('x')}, Y={points[0].get('y')}"
    return "No Click Found"

files = [
    "Memilih Menu Aktivitas.record",
    "Buat Aktivitas Harian.record",
    "Memilih Jenis.record",
    "Memilih SKP.record",
    "Untuk Mengetik Isian SKP.record",
    "posting.record"
]

print("ANALISA KOORDINAT REKAMAN USER:")
for f in files:
    coord = parse(f)
    print(f"File {f}: {coord}")
