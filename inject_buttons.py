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
        "ModKodeUnikPL",    # Generate kode unik otomatis untuk PL
        "ModAutoFit",       # Auto-fit baris
        "ModDraftPaket",    # Load draft paket dari Supabase + autofill
        "ModDraftPaketPL",  # Load draft paket PL dari Supabase + autofill
        "ModKKEvaluasi",    # Muat data KK Evaluasi Kualifikasi dari Supabase
        "ModInputBA",       # Muat sheet "0. Input BA" dari Supabase + GCal
        "ModSyncDraft",     # Sync Data Draft + Diff Highlight ke Supabase
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
            "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n"
            "    ' Placeholder — parsing dipicu manual via tombol Parse Draft\n"
            "End Sub\n"
            "\n"
            "Private Sub Workbook_Open()\n"
            "    On Error Resume Next\n"
            "    Dim sd As String\n"
            "    sd = ModWordLink.ScriptDir_Public()\n"
            "    If sd = \"\" Then Exit Sub\n"
            "    Dim testWord As String\n"
            "    testWord = ThisWorkbook.Path & Chr(92) & \"1. Full Dokumen BA PK v1.docx\"\n"
            "    If Dir(testWord) = \"\" Then Exit Sub\n"
            "    Dim wsh As Object\n"
            "    Set wsh = CreateObject(\"WScript.Shell\")\n"
            "    Dim cmd As String\n"
            "    cmd = \"powershell -ExecutionPolicy Bypass -File \"\"\" & sd & Chr(92) & \"relink_dotnet.ps1\"\" -ExcelPath \"\"\" & ThisWorkbook.FullName & \"\"\" -CheckOnly\"\n"
            "    Dim exitCode As Long\n"
            "    exitCode = wsh.Run(cmd, 0, True)\n"
            "    Set wsh = Nothing\n"
            "    If exitCode = 1 Then ModWordLink.RelinkTemplate\n"
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
        target_clean = ["1. Input Data", "@ Master Data", "3. KK Evaluasi Kualifikasi", "0. Input BA", "6. Harga Penawaran"]
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
        BLACK     = (40, 40, 40)
        GREEN     = (40, 167, 69)
        RED_PDF   = (200, 35, 51)
        PURPLE    = (102, 51, 153)
        TEAL      = (0, 128, 128)

        # --- @ Master Data: tombol di KANAN data, di bawah border "Dokumen Pemilihan" ---
        # "Dokumen Pemilihan" selalu di row 20, border selesai row 23.
        # Tombol mulai row 26 (2 baris gap setelah border) — konsisten semua paket.
        # Kolom J (col 10) ke kanan — area kosong di semua paket.
        # Baris 1 (yi=0): BukaBA PK | Print BA Reviu | — | —
        # Baris 2 (yi=1): PrintPembuktian | REvaluasi | — | —
        # Baris 3 (yi=2): PrintTimpang | MuatDraft | ParseDraft | KodeUnik
        # Baris 4 (yi=3): UpdateHPS | BukaReviu | PrintIsiReviu | PrintBAReviu2
        # Baris 5 (yi=4): BukaDokpil | PrintDokpil | Relink | —
        # Baris 6 (yi=5): SyncDraft | DiffHighlight | — | —
        # Baris 7 (yi=6): GabungBAReviu | RefreshDataTender | — | —
        _ws_md_anchor = wb.Sheets("@ Master Data")
        _AX = _ws_md_anchor.Cells(28, 6).Left + 14  # col F row 28 + 5mm kanan
        _AY = _ws_md_anchor.Rows(28).Top + 7        # row 28 + 2.5mm bawah
        _BW = 120.0    # button width
        _BH = 27.1     # button height
        _GX = 5.0      # gap horizontal antar tombol
        _GY = 31.0     # gap vertikal antar baris
        print(f"  Anchor: F28 Left={_AX:.1f}, row 28 Top={_AY:.1f}")

        master_data_btns = [
            # (name, label, macro, yi, xi, color)
            # Baris 1
            ("btnBukaBA",          "Buka BA PK",          "BukaBA",                   0, 0, BLUE_WORD),
            ("btnPrintBAReviu",    "Print BA Reviu",       "PrintBAReviuPDF",          0, 1, BLACK),
            # Baris 2
            ("btnPrintPembuktian", "Print BA Pembuktian", "PrintPembuktianPDF",        1, 0, BLACK),
            ("btnREvaluasi",       "Print REvaluasi",     "PrintREvaluasiPDF",         1, 1, BLACK),
            # Baris 3
            ("btnPembuktianTimpang","Print Timpang",       "PrintPembuktianTimpangPDF",2, 0, BLACK),
            ("btnParseDraft",      "Parse Ulang Draft Lokal", "ParseDraftTerpilih",     2, 1, GREEN),
            ("btnKodeUnik",        "Kode Unik Surat",     "GenerateKodeUnik",          2, 2, TEAL),
            # Baris 4: UpdateHPS masuk kolom 0
            ("btnUpdateHPS",       "Update HPS Saja",     "UpdateHPSSaja",             3, 0, (220, 53, 69)),
            ("btnBukaReviu",       "Buka Reviu",          "BukaReviu",                 3, 1, BLUE_WORD),
            ("btnPrintIsiReviu",   "Print Isi Reviu",     "PrintIsiReviuPDF",          3, 2, BLACK),
            ("btnPrintBAReviu2",   "Print BA Reviu",      "PrintBAReviuPDF",           3, 3, BLACK),
            # Baris 5
            ("btnBukaDokpil",      "Buka Dokpil",         "BukaDokpil",                4, 0, BLUE_WORD),
            ("btnPrintDokpil",     "Print Dokpil",        "PrintDokpilPDF",            4, 1, BLACK),
            ("btnRelink",          "Relink Template",     "RelinkTemplate",            4, 2, (255, 140, 0)),
            # Baris 6
            ("btnGabungBAReviu",   "Gabung BA Reviu",     "GabungBAReviu",             5, 0, (0, 128, 96)),
            ("btnRefreshPaket",    "Refresh Paket",       "RefreshPaket",              5, 1, (0, 128, 128)),
        ]

        def _add_master_btn(ws_md, name, label, macro, yi, xi, color):
            left = _AX + xi * (_BW + _GX)
            top  = _AY + yi * _GY
            shp  = ws_md.Shapes.AddShape(5, left, top, _BW, _BH)
            shp.Name = name
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
            print(f"    [OK] {label} -> {macro}")

        try:
            ws_md = wb.Sheets("@ Master Data")
            try:
                ws_md.Unprotect(Password="pokja2026")
            except Exception:
                pass
            try:
                ws_md.Range("F4:I8").UnMerge()
                ws_md.Range("F4:I8").Interior.Pattern = 0
                ws_md.Range("F4:I8").Font.Bold = False
                ws_md.Range("F4:I4").Merge()
                ws_md.Range("F4").Value = "INPUT TANGGAL BA REVIU DPP"
                ws_md.Range("F4").Font.Bold = True
                ws_md.Range("F4").Interior.Color = 238 + (215 * 256) + (189 * 65536)
                labels = [
                    ("F5", "Tanggal", "G5", "angka 1-31"),
                    ("F6", "Bulan", "G6", "angka 1-12"),
                    ("F7", "Tahun", "G7", "tahun"),
                    ("F8", "Hari", "G8", "otomatis"),
                ]
                for a, v, b, h in labels:
                    ws_md.Range(a).Value = v
                    ws_md.Range(b).Value = h
                if not str(ws_md.Range("H5").Value or "").strip():
                    ws_md.Range("H5").Value = 1
                if not str(ws_md.Range("H6").Value or "").strip():
                    ws_md.Range("H6").Value = 1
                if not str(ws_md.Range("H7").Value or "").strip():
                    ws_md.Range("H7").Value = 2026
                ws_md.Range("I5").Formula = '=CONCATENATE(H5," ",I6," ",H7)'
                ws_md.Range("I6").Formula = '=IF(H6=1,"Januari",IF(H6=2,"Februari",IF(H6=3,"Maret",IF(H6=4,"April",IF(H6=5,"Mei",IF(H6=6,"Juni",IF(H6=7,"Juli",IF(H6=8,"Agustus",IF(H6=9,"September",IF(H6=10,"Oktober",IF(H6=11,"November",IF(H6=12,"Desember",""))))))))))))'
                ws_md.Range("I7").Formula = "=DATE(H7,H6,H5)"
                ws_md.Range("H8").Formula = '=CHOOSE(WEEKDAY(I7),"Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu")'
                ws_md.Range("H5:H7").Interior.Color = 204 + (242 * 256) + (255 * 65536)
                ws_md.Range("H5:H7").NumberFormat = "0"
                ws_md.Range("I7").NumberFormat = "dd mmmm yyyy"
                ws_md.Range("F4:I8").Borders.LineStyle = 1
                print("    [OK] Panel tanggal BA Reviu DPP -> @ Master Data!F4:I8")
            except Exception as e:
                print(f"    [WARN] Panel tanggal BA Reviu: {e}")
            # Hapus tombol lama
            btn_names_md = {b[0] for b in master_data_btns}
            for shp in list(ws_md.Shapes):
                if shp.Name in btn_names_md:
                    shp.Delete()
                    print(f"    Deleted {shp.Name}")
            for name, label, macro, yi, xi, color in master_data_btns:
                try:
                    _add_master_btn(ws_md, name, label, macro, yi, xi, color)
                except Exception as e:
                    print(f"    [WARN] {label}: {e}")
        except Exception as e:
            print(f"    [WARN] Sheet '@ Master Data': {e}")

        # Sheet lain: pakai cell reference seperti biasa.
        # Tombol "Muat KK Evaluasi" (sheet 3) & "Muat & Sync" (0. Input BA) DIHAPUS —
        # KK Evaluasi + Input BA sekarang ditulis langsung dari Streamlit Asisten Pokja
        # (Tab 6 & Tab 7), tombol VBA jadi redundant. Pola sama dgn mode PL.
        other_sheet_buttons = [
            ("6. Harga Penawaran", [
                # K1 (col 11) — kanan dropdown J1, macro MuatHargaPenawaran
                ("btnMuatHarga", "Muat Harga Penawaran", "MuatHargaPenawaran", 1, 11, PURPLE),
            ]),
        ]

        for sheet_name, btns in other_sheet_buttons:
            try:
                ws = wb.Sheets(sheet_name)
                try:
                    ws.Unprotect(Password="pokja2026")
                except Exception:
                    pass
                for btn_name, label, macro, row, col, color in btns:
                    try:
                        cell = ws.Cells(row, col)
                        shp = ws.Shapes.AddShape(5, cell.Left, cell.Top, 120, 28)
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
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.4.xlsm"

    # Step 1: Setup @ Master Data via openpyxl (HARUS duluan, sebelum COM)
    # openpyxl tidak support shapes — kalau jalan setelah inject, tombol hilang
    print("--- Step 1: Setup @ Master Data (openpyxl) ---")
    try:
        from setup_master_data import setup_with_openpyxl
        setup_with_openpyxl(target)
    except Exception as e:
        print(f"[WARN] Setup Master Data gagal: {e}")

    # Step 2: Inject VBA + tombol via COM (shapes aman karena COM yang save terakhir)
    print("\n--- Step 2: Inject VBA + Buttons (COM) ---")
    inject_buttons(target)
