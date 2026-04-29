"""Fix circular reference + inject ulang modules."""
import win32com.client, pythoncom, time, os, sys

filepath = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.4.xlsm"
script_dir = os.path.dirname(os.path.abspath(__file__))

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(filepath)

    # Fix 1: Sheet 6 W29 — hapus formula circular
    ws6 = wb.Sheets("6. BA KLARIF SKP ALAT")
    ws6.Unprotect()
    ws6.Range("W29").Value = 0
    print("Fix W29: value=0")

    # Fix 2: "0. Input BA" C27/C28 — hapus formula circular
    ws0 = wb.Sheets("0. Input BA")
    ws0.Unprotect()
    ws0.Cells(27, 3).Value = ""
    ws0.Cells(27, 4).Value = ""
    ws0.Cells(27, 5).Value = ""
    ws0.Cells(28, 3).Value = ""
    print("Fix C27/C28: dikosongkan")

    wb.Save()
    print("Saved")
    time.sleep(1)
    wb.Close(SaveChanges=False)
    print("Done - circular reference diperbaiki")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback; traceback.print_exc()
finally:
    try: excel.Quit()
    except: pass
    pythoncom.CoUninitialize()
