"""
Cek sheet Excel mana yang berisi data untuk merge fields.
"""
import win32com.client
import pythoncom
import os

def check_sheets(filepath):
    filepath = os.path.abspath(filepath)
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        try: excel.Visible = False
        except: pass
        try: excel.DisplayAlerts = False
        except: pass
        
        wb = excel.Workbooks.Open(filepath, ReadOnly=True)
        
        # Cek setiap sheet: nama range / named range yang match merge fields
        print("=== NAMED RANGES ===")
        for i in range(1, wb.Names.Count + 1):
            try:
                n = wb.Names(i)
                print(f"  {n.Name} -> {n.RefersTo}")
            except:
                pass
        
        # Cek header di setiap sheet (baris 1-6) untuk mencocokkan merge fields
        print("\n=== SHEET HEADERS (baris 1-6) ===")
        for idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(idx)
            name = ws.Name
            ur = ws.UsedRange
            if ur.Rows.Count < 2:
                continue
            
            headers = []
            for row in range(1, min(7, ur.Rows.Count + 1)):
                for col in range(1, min(20, ur.Columns.Count + 1)):
                    val = ws.Cells(row, col).Value
                    if val and isinstance(val, str) and len(val) > 2:
                        headers.append(f"R{row}C{col}: {val[:50]}")
            
            if headers:
                print(f"\n  [{name}]")
                for h in headers[:10]:
                    print(f"    {h}")
        
        wb.Close(SaveChanges=False)
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
    check_sheets(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v.1.xlsm")
