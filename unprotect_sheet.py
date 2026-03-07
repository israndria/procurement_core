"""Buka gembok Excel Sheet '1. Input Data' selamanya."""
import win32com.client
import os

def unprotect_excel(filepath, sheet_name, password):
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(filepath)
        try:
            ws = wb.Sheets(sheet_name)
            ws.Unprotect(Password=password)
            print(f"Sukses unprotect sheet '{sheet_name}'.")
        except Exception as e:
            print(f"Gagal unprotect: {e}")
        wb.Save()
        wb.Close(SaveChanges=True)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        xl.Quit()

if __name__ == "__main__":
    fp = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.xlsm"
    unprotect_excel(fp, "1. Input Data", "pokja2026")
