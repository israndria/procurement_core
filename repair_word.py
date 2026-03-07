"""Repair corrupted Word files using Word's built-in repair."""
import win32com.client
import pythoncom
import os
import time

def repair_word_files(folder):
    pythoncom.CoInitialize()
    wdApp = None
    
    try:
        wdApp = win32com.client.DispatchEx("Word.Application")
        wdApp.Visible = False
        wdApp.DisplayAlerts = 0
        
        word_files = [
            "1. Full Dokumen BA PK v.1.docx",
            "2. Isi Reviu PK v.1.docx",
            "3. Dokpil Full PK v.1.docx",
        ]
        
        for wf in word_files:
            wp = os.path.join(folder, wf)
            if not os.path.exists(wp):
                print(f"  [SKIP] {wf}")
                continue
            
            print(f"  Repairing: {wf}...")
            try:
                # Open with repair mode
                doc = wdApp.Documents.Open(
                    FileName=wp,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    OpenAndRepair=True
                )
                
                # Save back (Word creates clean file)
                doc.Save()
                doc.Close(SaveChanges=False)
                print(f"    [OK] Repaired!")
                
            except Exception as e:
                print(f"    [ERROR] {e}")
                # Try alternative: open normally
                try:
                    doc = wdApp.Documents.Open(wp, False, False, False)
                    doc.SaveAs2(wp, FileFormat=12)  # 12 = wdFormatDocumentDefault (.docx)
                    doc.Close(False)
                    print(f"    [OK] Saved via SaveAs2")
                except Exception as e2:
                    print(f"    [FAIL] {e2}")
        
        time.sleep(1)
        wdApp.Quit()
        
    except Exception as e:
        print(f"[ERROR] {e}")
        if wdApp:
            try: wdApp.Quit()
            except: pass
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    folder = r"D:\Dokumen\@ POKJA 2026\Paket Experiment"
    print("Repairing Word files via Word OpenAndRepair...")
    repair_word_files(folder)
    print("\nDone!")
