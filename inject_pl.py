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

    # Attribute VB_Name menentukan nama module hasil Import, BUKAN nama file
    # temp. Pakai nama sementara dulu agar tidak bentrok dgn module lama yang
    # masih ada saat proses import (baru direname ke MOD_NAME setelah module
    # lama dihapus) -- tahan interupsi di tengah proses.
    tmp_mod_name = f"{MOD_NAME}_NEW"
    content_tmp = content.replace(f'Attribute VB_Name = "{MOD_NAME}"', f'Attribute VB_Name = "{tmp_mod_name}"')

    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False, mode="w", encoding="utf-8")
    tmp.write(content_tmp)
    tmp.close()
    tmp_path = tmp.name

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(filepath, 0, False)
        except Exception as open_error:
            # Workbook hasil patch manual kadang valid sebagai ZIP tetapi
            # ditolak parser Excel. xlRepairFile=1 meminta Excel memulihkan
            # workbook saat dibuka; tetap dipakai hanya sebagai fallback.
            print(f"  [WARN] Open normal gagal, coba Excel repair: {open_error}")
            wb = excel.Workbooks.Open(filepath, 0, False, None, None, None, None, None, 1)
            print("  [OK] Excel repair open berhasil")
        vb = wb.VBProject

        # PL tidak memakai generator kode unik otomatis. Bersihkan modul/button
        # legacy yang bisa ikut terbawa dari injector tender/umum.
        for legacy_name in ("ModKodeUnik", "ModKodeUnikPL"):
            for comp in list(vb.VBComponents):
                if comp.Name == legacy_name:
                    vb.VBComponents.Remove(comp)
                    print(f"  {legacy_name} legacy dihapus")
                    break

        # Import module baru DULU dengan nama sementara (tahan interupsi:
        # kalau proses mati di sini, module lama MOD_NAME masih utuh)
        imported = vb.VBComponents.Import(tmp_path)
        old_comp = None
        for comp in vb.VBComponents:
            if comp.Name == MOD_NAME:
                old_comp = comp
                break
        if old_comp:
            vb.VBComponents.Remove(old_comp)
            print(f"  {MOD_NAME} lama dihapus")
        imported.Name = MOD_NAME
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

        # Tombol di @ Master Data (Muat Paket PL + Isi Data PL sudah dihapus —
        # pengisian @ Master Data kini otomatis via COM saat buat folder).
        try:
            ws = wb.Sheets("@ Master Data")

            # Unprotect sheet sebelum modifikasi shape
            try:
                ws.Unprotect("pokja2026")
            except Exception:
                pass

            # Hapus tombol lama
            names_to_delete = []
            BTN_NAMES = ("btnMuatPL", "btnIsiPL", "btnKodeUnik", "btnBukaBA_PL", "btnBukaReviu_PL", "btnBukaDokpil_PL", "btnRelinkPL", "btnRefreshDataPL", "btnMuatHPS_PL", "btnCetakBAReviu_PL", "btnSyncDraftPL", "btnClearHighlightPL", "btnCetakDokpil_PL", "btnCetakReviu_PL", "btnGabungReviu_PL", "btnIsiEvaluasiPL", "btnCetakBAPLJKK", "btnGabungBAReviu", "btnGabungBAPLJKK")
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

            # Posisi absolut — diukur dari layout paket 1 (@ Master Data) setelah user rapikan
            # Baris 0 (Top=181.4): BukaDokpil | RelinkWord | RefreshDataPL | (kosong)
            # Baris 1 (Top=212.4): BukaBA | BukaReviu | SyncDraft | ClearHighlight
            # Baris 2 (Top=243.0): CetakBAReviu | CetakDokpil | (kosong) | (kosong)
            # Baris 3 (Top=274.9): CetakIsiReviu | GabungReviu | MuatHPS | IsiEvaluasiPL
            # Baris 4 (Top=305.4): CetakBAPLJKK (col 0)
            _X = [641.8, 776.9, 911.3, 1046.4]           # Left per kolom 0-3
            _W = [130.1, 129.9, 130.1, 130.4]             # Width per kolom 0-3
            _Y = [181.4, 212.4, 243.0, 274.9, 305.4]      # Top per baris 0-4
            _BTN_H = 27.4

            def add_btn(name, label, macro, yi, xi, rgb):
                """yi=indeks baris (0-4), xi=indeks kolom (0-3)"""
                shp = ws.Shapes.AddShape(5, _X[xi], _Y[yi], _W[xi], _BTN_H)
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

            # Baris 0: Buka Dokpil | Relink Word
            # (Muat Paket PL + Isi Data PL dihapus — @ Master Data kini diisi otomatis
            #  via COM saat buat folder di Streamlit, lihat isi_master_data_pl.py)
            add_btn("btnBukaDokpil_PL",   "Buka Dokpil",       "BukaDokpilPlJkk",         0, 0, TEAL)
            add_btn("btnRelinkPL",        "Relink Word",       "RelinkPL",                 0, 1, (128, 0, 0))
            add_btn("btnRefreshDataPL",   "Refresh Data PL",   "RefreshDataPL",            0, 2, (0, 150, 100))
            # Baris 1: Buka BA | Buka Reviu | (kosong) | (kosong)
            add_btn("btnBukaBA_PL",       "Buka BA",           "BukaBAPlJkk",             1, 0, ORANGE)
            add_btn("btnBukaReviu_PL",    "Buka Reviu",        "BukaReviuPlJkk",          1, 1, PURPLE)
            # Baris 2: Cetak BA Reviu PL | Cetak Dokpil PDF | (kosong) | (kosong)
            add_btn("btnCetakBAReviu_PL", "Cetak BA Reviu PL", "CetakBAReviuPLPDF",       2, 0, RED_DARK)
            add_btn("btnCetakDokpil_PL",  "Cetak Dokpil PDF",  "CetakDokpilPlJkkPDF",     2, 1, (0, 100, 180))
            # Baris 3: Cetak Isi Reviu (kolom 0) | Gabung Reviu (kolom 1) | Muat HPS (kolom 2) | Isi Evaluasi PL (kolom 3)
            add_btn("btnCetakReviu_PL",   "Cetak Isi Reviu",   "CetakReviuPlJkkPDF",          3, 0, (0, 120, 80))
            add_btn("btnGabungReviu_PL",  "Gabung BA Reviu",   "GabungBAReviu",               3, 1, (0, 128, 96))
            add_btn("btnMuatHPS_PL",      "Muat HPS",          "MuatHPSPL",                  3, 2, (200, 100, 0))
            add_btn("btnIsiEvaluasiPL",   "Isi Evaluasi PL",   "IsiEvaluasiPLStandalone",     3, 3, (160, 60, 0))
            # Baris 4: Cetak BA PLJKK (col 0) | Gabung BA PLJKK (col 1)
            add_btn("btnCetakBAPLJKK",    "Cetak BA PLJKK",    "CetakBAPLJKKPDF",             4, 0, (140, 20, 20))
            add_btn("btnGabungBAPLJKK",   "Gabung BA PLJKK",   "GabungBAPLJKK",               4, 1, (100, 20, 80))

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
