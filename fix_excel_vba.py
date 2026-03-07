"""
Script untuk menghapus kode VBA dari ThisWorkbook di file Excel .xlsm
"""
import win32com.client
import pythoncom
import sys
import os

def remove_vba_from_thisworkbook(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Membuka file: {filepath}")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        print("File berhasil dibuka (macro dinonaktifkan)")
        
        vb_project = wb.VBProject
        
        for comp in vb_project.VBComponents:
            if comp.Name == "ThisWorkbook":
                line_count = comp.CodeModule.CountOfLines
                if line_count > 0:
                    comp.CodeModule.DeleteLines(1, line_count)
                    print(f"[OK] Berhasil menghapus {line_count} baris kode VBA dari ThisWorkbook")
                else:
                    print("ThisWorkbook sudah kosong")
                break
        
        # Juga cek dan hapus dari semua Sheet modules
        for comp in vb_project.VBComponents:
            if comp.Type == 100:  # vbext_ct_Document (Sheet modules)
                line_count = comp.CodeModule.CountOfLines
                if line_count > 0:
                    comp.CodeModule.DeleteLines(1, line_count)
                    print(f"[OK] Juga menghapus {line_count} baris kode dari {comp.Name}")
        
        wb.Save()
        print("[OK] File berhasil disimpan!")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        print("\nJika error 'Programmatic access to Visual Basic Project is not trusted':")
        print("1. Buka Excel (buat file baru kosong)")
        print("2. File > Options > Trust Center > Trust Center Settings")
        print("3. Macro Settings > centang 'Trust access to the VBA project object model'")
        print("4. OK > OK, lalu tutup Excel dan jalankan script ini lagi")
        sys.exit(1)
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    target = r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026.xlsm"
    remove_vba_from_thisworkbook(target)
