"""
Diagnosa masalah tombol di Excel:
1. Cek VBA modules yang ada
2. Cek shapes/buttons yang ada di tiap sheet
3. Cek apakah macro bisa dijalankan
"""
import win32com.client
import pythoncom
import os

def diagnose(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Diagnosing: {filepath}")
    print(f"File exists: {os.path.exists(filepath)}")
    print(f"File size: {os.path.getsize(filepath):,} bytes")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath, ReadOnly=True)
        
        # 1. Cek VBA modules
        print("\n=== VBA MODULES ===")
        try:
            vb = wb.VBProject
            for comp in vb.VBComponents:
                print(f"  [{comp.Type}] {comp.Name} ({comp.CodeModule.CountOfLines} lines)")
                # Show first few lines of ModWordLink
                if comp.Name == "ModWordLink":
                    code = comp.CodeModule
                    print(f"    --- First 5 lines ---")
                    for i in range(1, min(6, code.CountOfLines + 1)):
                        print(f"    {i}: {code.Lines(i, 1)}")
                    print(f"    --- Last 5 lines ---")
                    for i in range(max(1, code.CountOfLines - 4), code.CountOfLines + 1):
                        print(f"    {i}: {code.Lines(i, 1)}")
        except Exception as e:
            print(f"  [ERROR] Cannot access VBProject: {e}")
            print(f"  -> Check: File > Options > Trust Center > Trust Center Settings > Macro Settings")
            print(f"  -> Enable 'Trust access to the VBA project object model'")
        
        # 2. Cek shapes di tiap sheet
        print("\n=== SHAPES/BUTTONS ===")
        target_sheets = ["1. Input Data", "database_reviu", "database_dokpil"]
        for sname in target_sheets:
            try:
                ws = wb.Sheets(sname)
                print(f"\n  [{sname}] Shapes: {ws.Shapes.Count}")
                for i in range(1, ws.Shapes.Count + 1):
                    shp = ws.Shapes(i)
                    try:
                        action = shp.OnAction
                    except:
                        action = "(none)"
                    try:
                        text = shp.TextFrame2.TextRange.Text
                    except:
                        text = "(no text)"
                    print(f"    {shp.Name}: type={shp.Type}, action='{action}', text='{text}'")
                    print(f"      pos: top={shp.Top:.0f}, left={shp.Left:.0f}, w={shp.Width:.0f}, h={shp.Height:.0f}")
            except Exception as e:
                print(f"  [{sname}] ERROR: {e}")
        
        # 3. Cek path
        print(f"\n=== PATHS ===")
        print(f"  ThisWorkbook.Path: {wb.Path}")
        print(f"  Word files in same folder:")
        word_files = [
            "1. Full Dokumen BA PK v.1.docx",
            "2. Isi Reviu PK v.1.docx",
            "3. Dokpil Full PK v.1.docx",
        ]
        for wf in word_files:
            fp = os.path.join(wb.Path, wf)
            exists = os.path.exists(fp)
            print(f"    {'[OK]' if exists else '[MISSING]'} {wf}")
        
        wb.Close(SaveChanges=False)
        
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    diagnose(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v.1.xlsm")
