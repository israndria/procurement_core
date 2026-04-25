"""
Inject ModJawabanReviu ke 2. Isi Reviu PK v2.docm + tambah tombol Simpan/Load.
"""
import win32com.client
import pythoncom
import os
import tempfile
import time
from dotenv import load_dotenv
import pathlib

def inject_reviu(docm_path=None):
    script_dir = os.path.dirname(os.path.abspath(__file__))

    if docm_path is None:
        docm_path = os.path.join(
            os.path.dirname(os.path.dirname(script_dir)),
            "Paket Experiment", "2. Isi Reviu PK v2.docm"
        )
    docm_path = os.path.abspath(docm_path)
    print(f"Injecting to: {docm_path}")

    # Load secret Supabase
    env_path = pathlib.Path(script_dir) / "secret_supabase.env"
    load_dotenv(env_path)
    sb_url = os.environ.get("SUPABASE_URL", "").strip('"')
    sb_key = os.environ.get("SUPABASE_KEY", "").strip('"')

    bas_path = os.path.join(script_dir, "ModJawabanReviu.bas")
    with open(bas_path, "r", encoding="utf-8") as f:
        content = f.read()
    content = content.replace("%%SUPABASE_URL%%", sb_url)
    content = content.replace("%%SUPABASE_KEY%%", sb_key)

    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False, mode="w", encoding="utf-8")
    tmp.write(content)
    tmp.close()

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(docm_path, False, False, False)
        vbp = doc.VBProject

        # Hapus modul lama jika ada
        for comp in list(vbp.VBComponents):
            if comp.Name == "ModJawabanReviu":
                vbp.VBComponents.Remove(comp)
                print("  ModJawabanReviu lama dihapus")
                break

        # Import modul baru
        imported = vbp.VBComponents.Import(tmp.name)
        print(f"  [OK] Imported: {imported.Name} ({imported.CodeModule.CountOfLines} baris)")

        # Tambah tombol via CustomDocumentProperties tidak supported di Word secara visual
        # Pakai pendekatan: tambah macro ke ThisDocument untuk bind ke shortcut / toolbar
        # Tapi lebih praktis: tambah ke ribbon via Quick Access — atau cukup macro saja
        # User bisa run via Alt+F8 atau kita tambah tombol via Shapes di header

        # Tambah 2 tombol Shape di halaman pertama (floating, pojok kanan atas)
        secs = doc.Sections(1)
        # Pakai InlineShape tidak bisa dipasang macro — pakai Shape di header
        # Lebih reliabel: tambah ke body paragraf pertama sebagai floating shape

        # Word tidak support Shape.OnAction seperti Excel
        # Macro diakses via Alt+F8 atau tombol di Excel (ModWordLink)
        print("  [INFO] Tombol tidak diinjek — akses via Alt+F8 atau tombol Excel")

        doc.Save()
        print("  [OK] Disimpan")
        time.sleep(1)

        # Verifikasi
        found = any(c.Name == "ModJawabanReviu" for c in doc.VBProject.VBComponents)
        print(f"  [{'OK' if found else 'WARN'}] ModJawabanReviu {'terverifikasi' if found else 'TIDAK DITEMUKAN'}")

        doc.Close(SaveChanges=False)
        doc = None
        print("\n[OK] Inject selesai!")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        os.unlink(tmp.name)
        if doc:
            try: doc.Close(SaveChanges=False)
            except: pass
        if word:
            try:
                word.DisplayAlerts = False
                word.Quit()
            except: pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else None
    inject_reviu(path)
