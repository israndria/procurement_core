"""Debug: cek apakah Content Control terbaca di dokumen hasil copy."""
import win32com.client
import pythoncom
import sys

DOCM_PATH = r"D:\Dokumen\@ POKJA 2026\1. Pokja 086 - Perbaikan - Peningkatan Jalan Desa Hatungun Rt.06, Kab. Tapin\2. Isi Reviu PK v2.docm"

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOCM_PATH, False, False, False)
        print(f"Dibuka: {doc.Name}")
        print(f"Jumlah Content Control: {doc.ContentControls.Count}")
        print()
        for cc in doc.ContentControls:
            teks = (cc.Range.Text or '')[:50].strip()
            print(f"  Tag='{cc.Tag}' | Title='{cc.Title}' | Teks='{teks}'")
        doc.Close(SaveChanges=False)
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        try: word.Quit()
        except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
