"""Analisa detail error #REF! dan struktur data untuk planning improvement."""
import win32com.client
import pythoncom
import os

def analyze_details(filepath):
    filepath = os.path.abspath(filepath)
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        
        # 1. Detail #REF! errors
        print("=" * 60)
        print("1. DETAIL ERROR #REF!")
        print("=" * 60)
        
        error_sheets = ["data_tender", "rincian_tender", "10. Dengan Nego Agit"]
        for sname in error_sheets:
            ws = None
            try:
                ws = wb.Sheets(sname)
            except:
                print(f"\n  [{sname}] TIDAK DITEMUKAN")
                continue
            
            print(f"\n  [{sname}]")
            ur = ws.UsedRange
            print(f"  Used Range: {ur.Address} ({ur.Rows.Count}x{ur.Columns.Count})")
            
            error_count = 0
            for row in range(1, min(ur.Rows.Count + 1, 60)):
                for col in range(1, min(ur.Columns.Count + 1, 30)):
                    cell = ws.Cells(row, col)
                    try:
                        if cell.HasFormula:
                            formula = cell.Formula
                            if "#REF!" in formula:
                                print(f"    {cell.Address}: {formula}")
                                error_count += 1
                    except:
                        pass
            print(f"  Total #REF! formulas: {error_count}")
        
        # 2. Struktur "1. Input Data" untuk protection & form
        print(f"\n{'=' * 60}")
        print("2. STRUKTUR '1. Input Data' (untuk Form & Protection)")
        print("=" * 60)
        
        ws = wb.Sheets("1. Input Data")
        ur = ws.UsedRange
        print(f"  Used Range: {ur.Address}")
        
        print("\n  Cell dengan data input (kolom E-G, baris 1-70):")
        for row in range(1, min(71, ur.Rows.Count + 1)):
            for col in range(4, 9):  # D-H
                cell = ws.Cells(row, col)
                val = cell.Value
                if val is not None and str(val).strip() != "":
                    has_formula = cell.HasFormula
                    label_cell = ws.Cells(row, col - 1)
                    label = label_cell.Value if label_cell.Value else ""
                    marker = " [FORMULA]" if has_formula else " [INPUT]"
                    if not has_formula:
                        print(f"    Row {row}, Col {col} ({cell.Address}): {str(val)[:50]}{marker} | Label: {str(label)[:30]}")
        
        # 3. Info sheet "7.2 Dengan Nego" untuk conditional formatting
        print(f"\n{'=' * 60}")
        print("3. KOLOM '% Terhadap HPS' di '7.2 Dengan Nego'")
        print("=" * 60)
        
        ws = wb.Sheets("7.2 Dengan Nego")
        # Cari kolom % Terhadap HPS
        for col in range(1, 25):
            val = ws.Cells(6, col).Value
            if val and "Terhadap" in str(val):
                print(f"  Kolom header: {ws.Cells(6, col).Address} = {val}")
                print(f"  Sample data:")
                for row in range(8, 35):
                    cv = ws.Cells(row, col).Value
                    if cv is not None:
                        print(f"    Row {row}: {cv}")
                break
            val7 = ws.Cells(7, col).Value
            if val7 and "Terhadap" in str(val7):
                print(f"  Kolom header: {ws.Cells(7, col).Address} = {val7}")
                break
        
        # 4. Info kolom Ket/Timpang
        print(f"\n  Kolom Keterangan/Timpang:")
        for col in range(1, 25):
            for hrow in [6, 7]:
                val = ws.Cells(hrow, col).Value
                if val and ("Ket" in str(val) or "Timpang" in str(val)):
                    print(f"  {ws.Cells(hrow, col).Address} (Col {col}): {val}")
        
        # 5. Nomor surat pattern
        print(f"\n{'=' * 60}")
        print("4. PATTERN NOMOR SURAT")
        print("=" * 60)
        
        surat_sheets = ["1. Surat Undangan", "9. BA Penetapan Pemenang", 
                        "11. Surat Pengantar Hasil", "15. BA Timpang"]
        for sname in surat_sheets:
            try:
                ws = wb.Sheets(sname)
                ur = ws.UsedRange
                for row in range(1, min(ur.Rows.Count + 1, 30)):
                    for col in range(1, min(ur.Columns.Count + 1, 25)):
                        cell = ws.Cells(row, col)
                        if cell.HasFormula and "Nomor" in str(cell.Formula):
                            print(f"  [{sname}] {cell.Address}: {cell.Formula}")
                        elif cell.Value and "Nomor" in str(cell.Value)[:20]:
                            print(f"  [{sname}] {cell.Address}: {cell.Value}")
            except:
                pass
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    analyze_details(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
