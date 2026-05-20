"""
Inject VBA ModDraftPaketPL ke 0. BAPLJKK - Template.xlsm (dan semua copy paket PL).
Terpisah dari inject_buttons.py (PK/tender) — tidak ada dependency silang.

Usage:
    python inject_pl.py                           # inject ke semua .xlsm BAPLJKK di POKJA root
    python inject_pl.py "path/to/file.xlsm"       # inject ke file spesifik
"""
import win32com.client
import pythoncom
import os
import sys
import tempfile
import glob
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
BAS_FILE = SCRIPT_DIR / "ModDraftPaketPL.bas"
MOD_NAME = "ModDraftPaketPL"
BAS_FILE_KODE_UNIK = SCRIPT_DIR / "ModKodeUnikPL.bas"
MOD_NAME_KODE_UNIK = "ModKodeUnikPL"

# Workbook_Open untuk BAPLJKK — hanya auto-relink Word template
WORKBOOK_OPEN_CODE = (
    "Private Sub Workbook_Open()\n"
    "    ' Workbook_Open dimatikan — relink manual lewat tombol Relink Word\n"
    "End Sub\n"
)


def inject_pl(filepath: str):
    filepath = os.path.abspath(filepath)
    print(f"\nInjecting PL module to: {filepath}")

    if not os.path.exists(filepath):
        print(f"  [ERROR] File tidak ditemukan: {filepath}")
        return False

    if not BAS_FILE.exists():
        print(f"  [ERROR] ModDraftPaketPL.bas tidak ditemukan: {BAS_FILE}")
        return False

    # Baca + substitusi secret
    content = BAS_FILE.read_text(encoding="utf-8")
    if "%%SUPABASE_URL%%" in content or "%%SUPABASE_KEY%%" in content:
        from dotenv import load_dotenv
        env_path = SCRIPT_DIR / "secret_supabase.env"
        load_dotenv(env_path)
        sb_url = os.environ.get("SUPABASE_URL", "").strip('"')
        sb_key = os.environ.get("SUPABASE_KEY", "").strip('"')
        if not sb_url or not sb_key:
            print("  [ERROR] SUPABASE_URL / SUPABASE_KEY tidak ditemukan di secret_supabase.env")
            return False
        content = content.replace("%%SUPABASE_URL%%", sb_url)
        content = content.replace("%%SUPABASE_KEY%%", sb_key)

    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False, mode="w", encoding="utf-8")
    tmp.write(content)
    tmp.close()
    tmp_path = tmp.name

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(filepath)
        vb = wb.VBProject

        # Hapus module lama jika ada
        for comp in vb.VBComponents:
            if comp.Name == MOD_NAME:
                vb.VBComponents.Remove(comp)
                print(f"  {MOD_NAME} lama dihapus")
                break

        # Import module baru (ModDraftPaketPL)
        imported = vb.VBComponents.Import(tmp_path)
        print(f"  [OK] {imported.Name} imported ({imported.CodeModule.CountOfLines} baris)")

        # Import ModKodeUnikPL
        if BAS_FILE_KODE_UNIK.exists():
            for comp in vb.VBComponents:
                if comp.Name == MOD_NAME_KODE_UNIK:
                    vb.VBComponents.Remove(comp)
                    print(f"  {MOD_NAME_KODE_UNIK} lama dihapus")
                    break
            imported_ku = vb.VBComponents.Import(str(BAS_FILE_KODE_UNIK))
            print(f"  [OK] {imported_ku.Name} imported ({imported_ku.CodeModule.CountOfLines} baris)")
        else:
            print(f"  [WARN] ModKodeUnikPL.bas tidak ditemukan, skip")

        # Inject Workbook_Open ke ThisWorkbook
        this_wb_comp = None
        for comp in vb.VBComponents:
            if comp.Name == "ThisWorkbook":
                this_wb_comp = comp
                break
        if this_wb_comp:
            cm = this_wb_comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(WORKBOOK_OPEN_CODE)
            print(f"  [OK] Workbook_Open injected ({cm.CountOfLines} baris)")

        # Tambah 2 tombol di @ Master Data:
        #   "Muat Paket PL"  — panggil MuatDraftPaketPL
        #   "Isi Data PL"    — panggil IsiDataPL
        try:
            ws = wb.Sheets("@ Master Data")

            # Unprotect sheet sebelum modifikasi shape
            try:
                ws.Unprotect("pokja2026")
            except Exception:
                pass

            # Hapus tombol lama
            names_to_delete = []
            BTN_NAMES = ("btnMuatPL", "btnIsiPL", "btnBukaBA_PL", "btnBukaReviu_PL", "btnBukaDokpil_PL", "btnRelinkPL", "btnKodeUnikPL", "btnMuatHPS_PL", "btnMuatKodeUnik_PL", "btnCetakBAReviu_PL", "btnSyncDraftPL", "btnClearHighlightPL", "btnCetakDokpil_PL", "btnCetakReviu_PL")
            for shp in ws.Shapes:
                if shp.Name in BTN_NAMES:
                    names_to_delete.append(shp.Name)
            for name in names_to_delete:
                try:
                    ws.Shapes(name).Delete()
                    print(f"  Tombol lama {name} dihapus")
                except Exception:
                    pass

            BLUE    = (43, 87, 154)
            GREEN_C = (40, 167, 69)
            ORANGE  = (200, 100, 0)
            PURPLE  = (102, 51, 153)
            TEAL    = (0, 128, 128)

            # Posisi absolut (Left, Top) diukur dari layout manual user
            # Baris 1 (Top≈270): MuatPL | IsiPL | BukaDokpil | RelinkWord
            # Baris 2 (Top≈301): BukaBA  | BukaReviu | SyncDraft | ClearHighlight
            # Baris 3 (Top≈332): KodeUnik | MuatKodeUnik | CetakBAReviu | CetakDokpil
            # Baris 4 (Top≈364): MuatHPS (col 2 saja)
            _BTN_W = 130
            _BTN_H = 28
            _X = [758.5, 893.5, 1028.0, 1163.0]   # 4 kolom X
            _Y = [269.5, 301.0, 332.0, 364.0]      # 4 baris Y

            def add_btn(name, label, macro, yi, xi, rgb):
                """yi=indeks baris (0-3), xi=indeks kolom (0-3)"""
                shp = ws.Shapes.AddShape(5, _X[xi], _Y[yi], _BTN_W, _BTN_H)
                shp.Name = name
                r, g, b = rgb
                shp.Fill.ForeColor.RGB = r + (g * 256) + (b * 65536)
                shp.Line.Visible = False
                tf = shp.TextFrame2
                tf.TextRange.Text = label
                tf.TextRange.Font.Fill.ForeColor.RGB = 16777215
                tf.TextRange.Font.Size = 10
                tf.TextRange.Font.Bold = True
                tf.TextRange.ParagraphFormat.Alignment = 2
                tf.VerticalAnchor = 3
                shp.OnAction = macro
                print(f"  [OK] {name} ({label}) -> {macro}")

            RED_DARK   = (180, 0, 0)
            GREY       = (100, 100, 100)
            GREEN_SYNC = (20, 140, 60)

            # Baris 0: Muat PL | Isi Data PL | Buka Dokpil | Relink Word
            add_btn("btnMuatPL",          "Muat Paket PL",    "MuatDraftPaketPL",        0, 0, BLUE)
            add_btn("btnIsiPL",           "Isi Data PL",       "IsiDataPL",               0, 1, GREEN_C)
            add_btn("btnBukaDokpil_PL",   "Buka Dokpil",       "BukaDokpilPlJkk",         0, 2, TEAL)
            add_btn("btnRelinkPL",        "Relink Word",       "RelinkPL",                 0, 3, (128, 0, 0))
            # Baris 1: Buka BA | Buka Reviu | Sync Data Draft | Clear Highlight
            add_btn("btnBukaBA_PL",       "Buka BA",           "BukaBAPlJkk",             1, 0, ORANGE)
            add_btn("btnBukaReviu_PL",    "Buka Reviu",        "BukaReviuPlJkk",          1, 1, PURPLE)
            add_btn("btnSyncDraftPL",     "Sync Data Draft",   "SyncDataDraftPL",          1, 2, GREEN_SYNC)
            add_btn("btnClearHighlightPL","Clear Highlight",    "ClearHighlightPL",         1, 3, GREY)
            # Baris 2: Kode Unik PL | Muat Kode Unik | Cetak BA Reviu PL | Cetak Dokpil PDF
            add_btn("btnKodeUnikPL",      "Kode Unik PL",      "GenerateKodeUnikPaketPL", 2, 0, (128, 0, 128))
            add_btn("btnMuatKodeUnik_PL", "Muat Kode Unik",    "MuatKodeUnikPL",          2, 1, (100, 0, 150))
            add_btn("btnCetakBAReviu_PL", "Cetak BA Reviu PL", "CetakBAReviuPLPDF",       2, 2, RED_DARK)
            add_btn("btnCetakDokpil_PL",  "Cetak Dokpil PDF",  "CetakDokpilPlJkkPDF",     2, 3, (0, 100, 180))
            # Baris 3: Cetak Isi Reviu (kolom 0) | Muat HPS (kolom 2)
            add_btn("btnCetakReviu_PL",   "Cetak Isi Reviu",   "CetakReviuPlJkkPDF",      3, 0, (0, 120, 80))
            add_btn("btnMuatHPS_PL",      "Muat HPS",          "MuatHPSPL",               3, 2, (200, 100, 0))

            # Sengaja TIDAK re-protect @ Master Data — user butuh edit bebas
            # (Aturan PL: sheet @ Master Data harus selalu unprotected)

        except Exception as e:
            print(f"  [WARN] Tombol @ Master Data: {e}")

        wb.Save()
        print(f"  [SAVED] {os.path.basename(filepath)}")
        return True

    except Exception as e:
        print(f"  [ERROR] {e}")
        return False
    finally:
        os.unlink(tmp_path)
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass
        try:
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


def find_bapljkk_files(root: str) -> list:
    """Cari semua file BAPLJKK*.xlsm di bawah root."""
    pattern = os.path.join(root, "**", "*BAPLJKK*.xlsm")
    return [f for f in glob.glob(pattern, recursive=True)
            if ".bak" not in f.lower() and "~$" not in os.path.basename(f)]


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # File spesifik dari argumen
        target = sys.argv[1]
        inject_pl(target)
    else:
        # Auto-scan POKJA root
        pokja_root = str(SCRIPT_DIR.parent.parent)
        files = find_bapljkk_files(pokja_root)
        if not files:
            print(f"Tidak ada file BAPLJKK*.xlsm ditemukan di: {pokja_root}")
            print("Usage: python inject_pl.py \"path/to/0. BAPLJKK - Template.xlsm\"")
            sys.exit(1)

        print(f"Ditemukan {len(files)} file BAPLJKK:")
        for f in files:
            print(f"  {f}")
        print()

        ok = 0
        for f in files:
            if inject_pl(f):
                ok += 1

        print(f"\n{'='*50}")
        print(f"Selesai: {ok}/{len(files)} file berhasil diinjeksi.")
