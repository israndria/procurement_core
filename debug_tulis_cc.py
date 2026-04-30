"""Debug: coba tulis ke CC rekomen_1 di experiment, lalu baca balik."""
import win32com.client
import pythoncom
import sys

DOCM_PATH = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\2. Isi Reviu PK v2.docm"
TAG = "rekomen_1"
TEKS_TEST = "INI TEKS TEST DARI PYTHON"

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOCM_PATH, False, False, False)
        print(f"Dibuka: {doc.Name} | CC count: {doc.ContentControls.Count}")

        cc_target = None
        for cc in doc.ContentControls:
            if cc.Tag == TAG:
                cc_target = cc
                break

        if cc_target is None:
            print(f"[ERROR] CC tag='{TAG}' tidak ditemukan!")
            doc.Close(SaveChanges=False)
            return

        print(f"Sebelum: Tag='{cc_target.Tag}' | LockCC={cc_target.LockContentControl} | LockContents={cc_target.LockContents}")
        print(f"  Teks saat ini: '{(cc_target.Range.Text or '')[:80].strip()}'")

        # Coba unlock
        cc_target.LockContentControl = False
        cc_target.LockContents = False
        print(f"Setelah unlock: LockCC={cc_target.LockContentControl} | LockContents={cc_target.LockContents}")

        # Coba tulis
        try:
            cc_target.Range.Text = TEKS_TEST
            print(f"[OK] Range.Text berhasil ditulis")
        except Exception as e:
            print(f"[FAIL] Range.Text gagal: {e}")
            try:
                cc_target.Range.Delete()
                cc_target.Range.InsertAfter(TEKS_TEST)
                print(f"[OK] InsertAfter berhasil")
            except Exception as e2:
                print(f"[FAIL] InsertAfter juga gagal: {e2}")

        # Baca balik
        print(f"Teks setelah tulis: '{(cc_target.Range.Text or '')[:80].strip()}'")

        doc.Save()
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
