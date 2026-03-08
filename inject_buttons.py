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
    
    bas_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ModWordLink.bas")
    if not os.path.exists(bas_path):
        print(f"[ERROR] .bas file not found: {bas_path}")
        return
    print(f"  .bas file: {bas_path}")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        vb_project = wb.VBProject
        
        # 1. Remove existing ModWordLink
        for comp in vb_project.VBComponents:
            if comp.Name == "ModWordLink":
                vb_project.VBComponents.Remove(comp)
                print("  ModWordLink lama dihapus")
                break
        
        # 2. Import .bas file (this is more reliable than Add+AddFromString)
        imported = vb_project.VBComponents.Import(bas_path)
        print(f"  [OK] Imported: {imported.Name} ({imported.CodeModule.CountOfLines} baris)")
        
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

        sheet_buttons = [
            ("1. Input Data", [
                ("btnBukaBA",        "Buka BA PK",              "BukaBA",           3, 6, BLUE_WORD),
                ("btnPrintBAReviu",  "Print BA Reviu",          "PrintBAReviuPDF",  3, 7, BLACK),
                ("btnImportWeb",     "Import Data LPSE",        "ImportHTML",       3, 8, GREEN),
                ("btnUndanganPDF",   "Print Undangan PDF",      "PrintUndanganPDF", 4, 6, RED_PDF),
                ("btnPrintPembuktian", "Print BA Pembuktian",   "PrintPembuktianPDF", 4, 7, BLACK),
            ]),
            ("database_reviu", [
                ("btnBukaReviu",   "Buka Reviu",      "BukaReviu",        2, 7, BLUE_WORD),
                ("btnPrintIsiReviu", "Print Isi Reviu", "PrintIsiReviuPDF", 2, 8, BLACK),
            ]),
            ("database_dokpil", [
                ("btnBukaDokpil",   "Buka Dokpil",   "BukaDokpil",   2, 6, BLUE_WORD),
                ("btnPrintDokpil",  "Print Dokpil",  "PrintDokpil",  2, 7, BLACK),
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
        found = False
        for comp in wb.VBProject.VBComponents:
            if comp.Name == "ModWordLink":
                found = True
                print(f"  [VERIFIED] ModWordLink: {comp.CodeModule.CountOfLines} baris")
                break
        if not found:
            print("  [FAIL] ModWordLink NOT found after save!")
        
        wb.Close(SaveChanges=False)
        wb = None
        
        # 7. Re-open and re-verify (paranoid check)
        print("\n  Re-open verify...")
        wb2 = excel.Workbooks.Open(filepath, ReadOnly=True)
        found2 = False
        for comp in wb2.VBProject.VBComponents:
            if comp.Name == "ModWordLink":
                found2 = True
                print(f"  [RE-VERIFY] ModWordLink: {comp.CodeModule.CountOfLines} baris - CONFIRMED!")
                break
        if not found2:
            print("  [RE-VERIFY FAIL] ModWordLink NOT persisted on disk!")
        wb2.Close(SaveChanges=False)
        wb = None
        
        print(f"\n{'[OK]' if found2 else '[FAIL]'} Injection complete!")
        
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
    inject_buttons(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.xlsm")
