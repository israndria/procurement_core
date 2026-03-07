"""
Analisa pola kode unik surat dari 18 folder POKJA 2025.
Mengekstrak: Nama Paket, Kode Unik, Nomor Surat dari setiap file Excel.
"""
import win32com.client
import pythoncom
import os
import json
import re
import glob

BASE_DIR = r"D:\Dokumen\3 @ POKJA 2025\@ POKJA 2025"

def find_excel_files(base_dir):
    """Cari semua file .xlsm di folder-folder Pokja"""
    results = []
    for item in os.listdir(base_dir):
        full_path = os.path.join(base_dir, item)
        if os.path.isdir(full_path) and "Pokja" in item:
            # Cari file .xlsm di folder ini
            for f in os.listdir(full_path):
                if f.endswith(".xlsm") and "BA PK" in f and not f.startswith("~"):
                    results.append({
                        "folder": item,
                        "file": os.path.join(full_path, f)
                    })
    return results

def extract_data(wb, folder_name):
    """Ekstrak nama paket, kode unik, dan nomor surat dari workbook"""
    data = {"folder": folder_name}
    
    # 1. Nama Paket dari "1. Input Data" sheet
    try:
        ws = wb.Sheets("1. Input Data")
        
        # Nama paket biasa di F6
        nama_paket = ""
        for row in [6, 5, 7]:
            for col in [6, 5]:
                val = ws.Cells(row, col).Value
                if val and len(str(val)) > 10 and not str(val).startswith("="):
                    if any(kw in str(val).lower() for kw in ["perbaikan", "peningkatan", "pembangunan", "rehabilitasi", "jalan", "gedung", "jembatan", "drainase", "gorong", "irigasi", "saluran"]):
                        nama_paket = str(val)
                        break
            if nama_paket:
                break
        
        if not nama_paket:
            # Fallback: cari di range yang lebih luas
            for row in range(1, 20):
                for col in range(4, 9):
                    val = ws.Cells(row, col).Value
                    if val and len(str(val)) > 15:
                        if any(kw in str(val).lower() for kw in ["perbaikan", "peningkatan", "pembangunan", "rehabilitasi"]):
                            nama_paket = str(val)
                            break
                if nama_paket:
                    break
        
        data["nama_paket"] = nama_paket
        
        # Nomor SPPH atau kode unik dari F38 atau G38
        for row in [38, 37, 39]:
            for col in [6, 7, 5]:
                val = ws.Cells(row, col).Value
                if val and ("POKJA" in str(val).upper() or "SPPH" in str(val).upper() or "PJ" in str(val)):
                    data["kode_surat_input"] = str(val)
                    break
            if "kode_surat_input" in data:
                break
    except:
        pass
    
    # 2. Cari kode unik dari sheet-sheet BA (nomor surat)
    nomor_surat_sheets = [
        "9. BA Penetapan Pemenang",
        "15. BA Timpang", 
        "7. BA Klarifikasi HS",
        "3. BA Pembukaan Penawaran",
        "10. BA Klarifikasi Non Nego (2)",
    ]
    
    kode_unik_candidates = []
    
    for sname in nomor_surat_sheets:
        try:
            ws = wb.Sheets(sname)
            ur = ws.UsedRange
            for row in range(1, min(ur.Rows.Count + 1, 25)):
                for col in range(1, min(ur.Columns.Count + 1, 25)):
                    cell = ws.Cells(row, col)
                    val = cell.Value
                    formula = ""
                    try:
                        if cell.HasFormula:
                            formula = cell.Formula
                    except:
                        pass
                    
                    # Cari pattern nomor surat
                    text = str(val) if val else ""
                    if formula:
                        text = text + " |FORMULA| " + formula
                    
                    if "Nomor" in text and ("POKJA" in text.upper() or "UKPBJ" in text.upper()):
                        # Ekstrak kode unik dari nomor surat
                        # Pattern: 000.3.3/0XX/T/KODE_UNIK/POKJAXXX/UKPBJ/YYYY
                        match = re.search(r'/T/([^/]+)/POKJA', text)
                        if match:
                            kode_unik_candidates.append({
                                "sheet": sname,
                                "cell": cell.Address,
                                "kode_unik": match.group(1),
                                "full_text": text[:200]
                            })
                        
                        # Juga cari di formula
                        if formula:
                            match = re.search(r'/T/([^/]+)/POKJA', formula)
                            if match:
                                kode_unik_candidates.append({
                                    "sheet": sname,
                                    "cell": cell.Address,
                                    "kode_unik": match.group(1),
                                    "full_text": formula[:200],
                                    "source": "formula"
                                })
        except:
            pass
    
    # Juga cek langsung di cell yang berisi nomor surat (bukan formula)
    try:
        ws = wb.Sheets("9. BA Penetapan Pemenang")
        for row in range(5, 20):
            for col in range(1, 25):
                val = ws.Cells(row, col).Value
                if val and "/T/" in str(val) and "POKJA" in str(val).upper():
                    match = re.search(r'/T/([^/]+)/POKJA', str(val))
                    if match:
                        kode_unik_candidates.append({
                            "sheet": "9. BA Penetapan Pemenang",
                            "cell": ws.Cells(row, col).Address,
                            "kode_unik": match.group(1),
                            "full_text": str(val)[:200],
                            "source": "value"
                        })
    except:
        pass
    
    # Deduplicate
    seen = set()
    unique_candidates = []
    for c in kode_unik_candidates:
        if c["kode_unik"] not in seen:
            seen.add(c["kode_unik"])
            unique_candidates.append(c)
    
    data["kode_unik_candidates"] = unique_candidates
    
    # Simpel: ambil kode unik pertama
    if unique_candidates:
        data["kode_unik"] = unique_candidates[0]["kode_unik"]
    
    return data


def main():
    print("=" * 70)
    print("ANALISA POLA KODE UNIK SURAT - 18 POKJA 2025")
    print("=" * 70)
    
    files = find_excel_files(BASE_DIR)
    print(f"\nDitemukan {len(files)} file Excel:\n")
    for f in files:
        print(f"  {f['folder']}: {os.path.basename(f['file'])}")
    
    pythoncom.CoInitialize()
    excel = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        try:
            excel.Visible = False
        except:
            pass
        try:
            excel.DisplayAlerts = False
        except:
            pass
        
        all_data = []
        
        for idx, f in enumerate(files):
            print(f"\n--- [{idx+1}/{len(files)}] {f['folder']} ---")
            
            try:
                wb = excel.Workbooks.Open(os.path.abspath(f["file"]), ReadOnly=True)
                data = extract_data(wb, f["folder"])
                all_data.append(data)
                
                print(f"  Nama Paket: {data.get('nama_paket', 'N/A')[:80]}")
                print(f"  Kode Unik: {data.get('kode_unik', 'N/A')}")
                if data.get("kode_surat_input"):
                    print(f"  Kode Surat (input): {data['kode_surat_input'][:80]}")
                
                wb.Close(SaveChanges=False)
            except Exception as e:
                print(f"  [ERROR] {e}")
        
        # Ringkasan pola
        print(f"\n{'=' * 70}")
        print("RINGKASAN POLA PENAMAAN KODE UNIK")
        print(f"{'=' * 70}\n")
        
        print(f"{'No':<4} {'Nama Paket':<55} {'Kode Unik':<25}")
        print("-" * 84)
        
        for i, d in enumerate(all_data):
            nama = d.get("nama_paket", "N/A")[:53]
            kode = d.get("kode_unik", "N/A")
            print(f"{i+1:<4} {nama:<55} {kode:<25}")
        
        # Simpan ke JSON untuk analisa lanjut
        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kode_unik_analysis.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)
        print(f"\n[OK] Detail disimpan ke: {output_path}")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
