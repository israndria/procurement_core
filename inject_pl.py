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

# Workbook_Open untuk BAPLJKK — hanya auto-relink Word template
WORKBOOK_OPEN_CODE = (
    "Private Sub Workbook_Open()\n"
    "    On Error Resume Next\n"
    "    ' Auto-relink Word templates ke Excel ini jika path berubah\n"
    "    Dim testWord As String\n"
    "    testWord = ThisWorkbook.Path & Chr(92) & \"3. Dokpil Full PLJKK v1.docx\"\n"
    "    If Dir(testWord) = \"\" Then Exit Sub\n"
    "    Dim sd As String\n"
    "    On Error Resume Next\n"
    "    sd = ModWordLink.ScriptDir_Public()\n"
    "    On Error GoTo 0\n"
    "    If sd = \"\" Then Exit Sub\n"
    "    Dim wsh As Object\n"
    "    Set wsh = CreateObject(\"WScript.Shell\")\n"
    "    Dim cmd As String\n"
    "    cmd = \"powershell -ExecutionPolicy Bypass -File \"\"\" & sd & Chr(92) & \"relink_dotnet.ps1\"\" -ExcelPath \"\"\" & ThisWorkbook.FullName & \"\"\" -CheckOnly\"\n"
    "    Dim exitCode As Long\n"
    "    exitCode = wsh.Run(cmd, 0, True)\n"
    "    If exitCode = 1 Then ModWordLink.RelinkTemplate\n"
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

        # Import module baru
        imported = vb.VBComponents.Import(tmp_path)
        print(f"  [OK] {imported.Name} imported ({imported.CodeModule.CountOfLines} baris)")

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

            # Hapus tombol lama
            names_to_delete = []
            BTN_NAMES = ("btnMuatPL", "btnIsiPL", "btnBukaBA_PL", "btnBukaReviu_PL", "btnBukaDokpil_PL", "btnRelinkPL")
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

            def add_btn(name, label, macro, row, col, rgb, width=120):
                cell = ws.Cells(row, col)
                shp = ws.Shapes.AddShape(5, cell.Left, cell.Top, width, 28)
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

            # Baris 1: Muat + Isi Data
            add_btn("btnMuatPL",          "Muat Paket PL",  "MuatDraftPaketPL",  1, 7, BLUE)
            add_btn("btnIsiPL",           "Isi Data PL",    "IsiDataPL",          1, 8, GREEN_C)
            # Baris 2: Buka Word dokumen + Relink
            add_btn("btnBukaBA_PL",       "Buka BA",        "BukaBAPlJkk",        2, 7, ORANGE)
            add_btn("btnBukaReviu_PL",    "Buka Reviu",     "BukaReviuPlJkk",     2, 8, PURPLE)
            add_btn("btnBukaDokpil_PL",   "Buka Dokpil",    "BukaDokpilPlJkk",    2, 9, TEAL)
            add_btn("btnRelinkPL",        "Relink Word",    "RelinkPL",            2, 10, (128, 0, 0))

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
