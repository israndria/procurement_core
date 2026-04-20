"""
Inject VBA module (via Import .bas) + tombol di 3 sheet.
Uses VBComponents.Import instead of Add+AddFromString for reliable persistence.
"""
import win32com.client
import pythoncom
import os
import time

def inject_buttons(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Injecting to: {filepath}")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Semua module yang akan di-inject (urutan penting: dependencies dulu)
    modules_to_inject = [
        "ModWordLink",      # Core: buka/print Word, relink, import LPSE
        "ModTerbilang",     # Fungsi terbilang (angka → kata)
        "ModUtilitas",      # Export PDF, Direct Print, konversi bulan
        "ModNavigator",     # Navigasi sheet, input data tender, daftar isi
        "ModKodeUnik",      # Generate kode unik otomatis
        "ModAutoFit",       # Auto-fit baris
        "ModDraftPaket",    # Load draft paket dari Supabase + autofill
    ]

    # Verify semua .bas file ada
    bas_files = {}
    for mod_name in modules_to_inject:
        bas_path = os.path.join(script_dir, f"{mod_name}.bas")
        if os.path.exists(bas_path):
            bas_files[mod_name] = bas_path
        else:
            print(f"  [SKIP] {mod_name}.bas tidak ditemukan")

    if "ModWordLink" not in bas_files:
        print(f"[ERROR] ModWordLink.bas wajib ada!")
        return

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(filepath)
        vb_project = wb.VBProject

        # 1. Remove existing modules yang akan di-inject (+ cleanup ModSearchDrop)
        remove_targets = set(bas_files.keys()) | {"ModSearchDrop"}
        names_to_remove = []
        for comp in vb_project.VBComponents:
            if comp.Name in remove_targets:
                names_to_remove.append(comp.Name)
        for name in names_to_remove:
            for comp in vb_project.VBComponents:
                if comp.Name == name:
                    vb_project.VBComponents.Remove(comp)
                    print(f"  {name} lama dihapus")
                    break

        # Baca secret Supabase untuk substitusi placeholder di ModDraftPaket.bas
        import re as _re
        from dotenv import load_dotenv
        import pathlib
        env_path = pathlib.Path(script_dir) / "secret_supabase.env"
        load_dotenv(env_path)
        sb_url = os.environ.get("SUPABASE_URL", "").strip('"')
        sb_key = os.environ.get("SUPABASE_KEY", "").strip('"')

        # 2. Import semua .bas files
        import tempfile
        for mod_name, bas_path in bas_files.items():
            with open(bas_path, "r", encoding="utf-8") as f:
                content = f.read()
            # Substitusi placeholder secret (hanya ModDraftPaket)
            if "%%SUPABASE_URL%%" in content or "%%SUPABASE_KEY%%" in content:
                content = content.replace("%%SUPABASE_URL%%", sb_url)
                content = content.replace("%%SUPABASE_KEY%%", sb_key)
                tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False,
                                                   mode="w", encoding="utf-8")
                tmp.write(content)
                tmp.close()
                actual_path = tmp.name
            else:
                actual_path = bas_path
            imported = vb_project.VBComponents.Import(actual_path)
            if actual_path != bas_path:
                os.unlink(actual_path)
            print(f"  [OK] Imported: {imported.Name} ({imported.CodeModule.CountOfLines} baris)")
        
        # 2b. Inject Workbook_Open auto-relink ke ThisWorkbook module
        auto_relink_code = (
            "' Worksheet_Change untuk Sheet '1. Input Data'\n"
            "' Ketika cell E2 (selector paket) berubah → autofill semua field\n"
            "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n"
            "    If Sh.Name = \"1. Input Data\" Then\n"
            "        If Not Intersect(Target, Sh.Range(\"E2\")) Is Nothing Then\n"
            "            Dim val As String\n"
            "            val = Trim(CStr(Sh.Range(\"E2\").Value))\n"
            "            If val <> \"\" Then\n"
            "                ModDraftPaket.PilihDraftPaket val\n"
            "            End If\n"
            "        End If\n"
            "    End If\n"
            "End Sub\n"
            "\n"
            "Private Sub Workbook_Open()\n"
            "    ' Auto-relink jika path Excel berubah (misal pindah PC/drive)\n"
            "    On Error Resume Next\n"
            "    Dim needRelink As Boolean\n"
            "    needRelink = False\n"
            "    \n"
            "    ' Cek apakah ScriptDir bisa ditemukan\n"
            "    Dim sd As String\n"
            "    sd = ModWordLink.ScriptDir_Public()\n"
            "    If sd = \"\" Then Exit Sub\n"
            "    \n"
            "    ' Cek file Word pertama — apakah ada?\n"
            "    Dim testWord As String\n"
            "    testWord = ThisWorkbook.Path & \"\\\" & \"1. Full Dokumen BA PK v1.docx\"\n"
            "    If Dir(testWord) = \"\" Then Exit Sub\n"
            "    \n"
            "    ' Baca byte kecil dari settings.xml untuk cek path\n"
            "    ' Jika path Excel di settings.xml tidak cocok → relink\n"
            "    Dim fso As Object\n"
            "    Set fso = CreateObject(\"Scripting.FileSystemObject\")\n"
            "    Dim zipPath As String\n"
            "    zipPath = testWord\n"
            "    \n"
            "    ' Simple check: jalankan relink_dotnet.ps1 dengan flag --check-only\n"
            "    ' yang return exit code 0 jika sudah cocok, 1 jika perlu relink\n"
            "    Dim cmd As String\n"
            "    Dim wsh As Object\n"
            "    Set wsh = CreateObject(\"WScript.Shell\")\n"
            "    cmd = \"powershell -ExecutionPolicy Bypass -File \"\"\" & sd & \"\\relink_dotnet.ps1\"\"\" & \" -ExcelPath \"\"\" & ThisWorkbook.FullName & \"\"\" -CheckOnly\"\n"
            "    Dim exitCode As Long\n"
            "    exitCode = wsh.Run(cmd, 0, True)\n"
            "    Set wsh = Nothing\n"
            "    \n"
            "    If exitCode = 1 Then\n"
            "        ' Path berubah — auto relink\n"
            "        ModWordLink.RelinkTemplate\n"
            "    End If\n"
            "End Sub\n"
        )

        # Tulis ke ThisWorkbook module
        this_wb = None
        for comp in vb_project.VBComponents:
            if comp.Name == "ThisWorkbook":
                this_wb = comp
                break

        if this_wb:
            cm = this_wb.CodeModule
            # Hapus kode lama (kecuali Option Explicit di baris 1-2)
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(auto_relink_code)
            print(f"  [OK] Workbook_Open auto-relink injected ({cm.CountOfLines} baris)")

        # 3. Clean old buttons - langsung target 3 sheet yang diketahui ada tombolnya
        print("\n  Cleaning old buttons...")
        target_clean = ["1. Input Data", "database_reviu", "database_dokpil"]
        for sheet_name in target_clean:
            try:
                ws = wb.Sheets(sheet_name)
                shapes_to_delete = [shp.Name for shp in ws.Shapes if shp.Name.startswith("btn")]
                for name in shapes_to_delete:
                    ws.Shapes(name).Delete()
                    print(f"    Deleted {name} from {sheet_name}")
            except Exception as e:
                print(f"    [WARN] Clean {sheet_name}: {e}")
        
        # 4. Add buttons
        print("\n  Adding buttons...")
        
        BLUE_WORD = (43, 87, 154)
        BLACK = (40, 40, 40)
        GREEN = (40, 167, 69)
        RED_PDF = (200, 35, 51)
        PURPLE = (102, 51, 153)
        TEAL = (0, 128, 128)

        sheet_buttons = [
            ("1. Input Data", [
                ("btnBukaBA",        "Buka BA PK",              "BukaBA",           3, 6, BLUE_WORD),
                ("btnPrintBAReviu",  "Print BA Reviu",          "PrintBAReviuPDF",  3, 7, BLACK),
                ("btnImportWeb",     "Import Data LPSE",        "ImportHTML",       3, 8, GREEN),
                ("btnUndanganPDF",   "Print Undangan PDF",      "PrintUndanganPDF", 4, 6, RED_PDF),
                ("btnPrintPembuktian", "Print BA Pembuktian",   "PrintPembuktianPDF", 4, 7, BLACK),
                ("btnREvaluasi",       "Print REvaluasi",      "PrintREvaluasiPDF", 4, 8, BLACK),
                ("btnPembuktianTimpang", "Print Timpang",      "PrintPembuktianTimpangPDF", 5, 6, BLACK),
                ("btnMuatDraft",     "Muat Draft Paket",       "MuatDraftPaket",   5, 7, PURPLE),
                ("btnKodeUnik",      "Kode Unik Surat",        "GenerateKodeUnik", 5, 8, TEAL),
                ("btnRelink",        "Relink Template",        "RelinkTemplate",   6, 8, (255, 140, 0)),
            ]),
            ("database_reviu", [
                ("btnBukaReviu",   "Buka Reviu",      "BukaReviu",        2, 7, BLUE_WORD),
                ("btnPrintIsiReviu", "Print Isi Reviu", "PrintIsiReviuPDF", 2, 8, BLACK),
            ]),
            ("database_dokpil", [
                ("btnBukaDokpil",   "Buka Dokpil",   "BukaDokpil",   2, 6, BLUE_WORD),
                ("btnPrintDokpil",  "Print Dokpil",  "PrintDokpilPDF",  2, 7, BLACK),
            ]),
        ]
        
        for sheet_name, btns in sheet_buttons:
            try:
                ws = wb.Sheets(sheet_name)
                
                try:
                    ws.Unprotect(Password="pokja2026")
                except:
                    pass
                
                for btn_name, label, macro, row, col, color in btns:
                    try:
                        cell = ws.Cells(row, col)
                        top = cell.Top
                        left = cell.Left
                        
                        shp = ws.Shapes.AddShape(5, left, top, 120, 28)
                        shp.Name = btn_name
                        
                        r, g, b = color
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
                        
                        print(f"    [OK] {label} -> {macro} @ {sheet_name}!{cell.Address}")
                    except Exception as e:
                        print(f"    [WARN] {label}: {e}")
                # (Sengaja tidak Protect kembali - sheet dibiarkan bebas edit)
                    
            except Exception as e:
                print(f"    [WARN] Sheet '{sheet_name}': {e}")
        
        # 5. Save
        print("\n  Saving...")
        try:
            wb.Save()
        except Exception as e:
            print(f"  [WARN] wb.Save() failed: {e}, trying SaveAs...")
            try:
                wb.SaveAs(filepath)
            except Exception as e2:
                print(f"  [WARN] SaveAs also failed: {e2}")
        time.sleep(2)
        
        # 6. Final verify
        print("  Verifying...")
        verified_count = 0
        for comp in wb.VBProject.VBComponents:
            if comp.Name in bas_files:
                verified_count += 1
                print(f"  [VERIFIED] {comp.Name}: {comp.CodeModule.CountOfLines} baris")
        if verified_count < len(bas_files):
            print(f"  [WARN] Hanya {verified_count}/{len(bas_files)} module terverifikasi!")

        wb.Close(SaveChanges=False)
        wb = None

        # 7. Re-open and re-verify (paranoid check)
        print("\n  Re-open verify...")
        wb2 = excel.Workbooks.Open(filepath, ReadOnly=True)
        verified2 = 0
        for comp in wb2.VBProject.VBComponents:
            if comp.Name in bas_files:
                verified2 += 1
                print(f"  [RE-VERIFY] {comp.Name}: {comp.CodeModule.CountOfLines} baris - CONFIRMED!")
        if verified2 < len(bas_files):
            print(f"  [RE-VERIFY WARN] Hanya {verified2}/{len(bas_files)} module persisted!")
        wb2.Close(SaveChanges=False)
        wb = None

        all_ok = verified2 == len(bas_files)
        print(f"\n{'[OK]' if all_ok else '[PARTIAL]'} Injection complete! ({verified2}/{len(bas_files)} modules)")
        
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try:
                excel.DisplayAlerts = False
                excel.Quit()
            except: pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    inject_buttons(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.4.xlsm")
