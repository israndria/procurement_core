"""
Script untuk menganalisa file Excel secara mendalam:
- VBA modules dan kode
- Formula di setiap sheet
- Struktur data, merged cells, named ranges
- Potensi improvement
"""
import win32com.client
import pythoncom
import os
import json
from collections import defaultdict

def analyze_excel(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Menganalisa: {filepath}\n")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    report = {
        "file": filepath,
        "sheets": [],
        "vba_modules": [],
        "named_ranges": [],
        "summary": {}
    }
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        
        # ===== 1. ANALISA SHEET =====
        print("=" * 70)
        print("1. STRUKTUR SHEET")
        print("=" * 70)
        total_formulas = 0
        total_merged = 0
        formula_types = defaultdict(int)
        
        for idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(idx)
            sheet_info = {
                "index": idx,
                "name": ws.Name,
                "codename": ws.CodeName,
                "visible": ws.Visible,
                "used_range": "",
                "row_count": 0,
                "col_count": 0,
                "formula_count": 0,
                "merged_count": 0,
                "formulas_sample": [],
                "issues": []
            }
            
            try:
                ur = ws.UsedRange
                sheet_info["used_range"] = ur.Address
                sheet_info["row_count"] = ur.Rows.Count
                sheet_info["col_count"] = ur.Columns.Count
                
                # Hitung formula dan merged cells
                formulas_found = []
                for row in range(1, min(ur.Rows.Count + 1, 100)):  # Limit 100 baris
                    for col in range(1, min(ur.Columns.Count + 1, 30)):  # Limit 30 kolom
                        cell = ws.Cells(row, col)
                        try:
                            if cell.HasFormula:
                                formula = cell.Formula
                                sheet_info["formula_count"] += 1
                                total_formulas += 1
                                
                                # Klasifikasi formula
                                formula_upper = formula.upper()
                                for func in ["INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "IF", "IFERROR", 
                                            "SUMIF", "COUNTIF", "CONCATENATE", "LEFT", "RIGHT", "MID",
                                            "INDIRECT", "OFFSET", "SUMPRODUCT", "ARRAYFORMULA",
                                            "IFNA", "ISBLANK", "LEN", "SUBSTITUTE", "TEXT"]:
                                    if func in formula_upper:
                                        formula_types[func] += 1
                                
                                # Simpan sample formula (max 5 per sheet)
                                if len(formulas_found) < 5:
                                    formulas_found.append({
                                        "cell": cell.Address,
                                        "formula": formula
                                    })
                                
                                # Deteksi error
                                try:
                                    val = cell.Value
                                    if val is not None and isinstance(val, int) and val < -2000000000:
                                        sheet_info["issues"].append(f"{cell.Address}: Error value (kemungkinan #REF! atau #N/A)")
                                except:
                                    pass
                            
                            if cell.MergeCells:
                                if cell.MergeArea.Cells(1,1).Address == cell.Address:
                                    sheet_info["merged_count"] += 1
                                    total_merged += 1
                        except:
                            pass
                
                sheet_info["formulas_sample"] = formulas_found
                
            except Exception as e:
                sheet_info["issues"].append(f"Error membaca sheet: {e}")
            
            report["sheets"].append(sheet_info)
            
            vis = "Visible" if ws.Visible == -1 else "Hidden"
            print(f"  [{idx:2d}] {ws.Name:<45} | {vis} | {sheet_info['row_count']:>4}x{sheet_info['col_count']:<4} | F:{sheet_info['formula_count']:>4} | M:{sheet_info['merged_count']:>3}")
        
        # ===== 2. ANALISA VBA =====
        print(f"\n{'=' * 70}")
        print("2. VBA MODULES")
        print("=" * 70)
        
        try:
            vb_project = wb.VBProject
            for comp in vb_project.VBComponents:
                comp_type_map = {1: "Module", 2: "ClassModule", 3: "Form", 100: "Document"}
                comp_type = comp_type_map.get(comp.Type, f"Unknown({comp.Type})")
                
                line_count = comp.CodeModule.CountOfLines
                code = ""
                if line_count > 0:
                    code = comp.CodeModule.Lines(1, line_count)
                
                vba_info = {
                    "name": comp.Name,
                    "type": comp_type,
                    "line_count": line_count,
                    "code": code
                }
                report["vba_modules"].append(vba_info)
                
                if line_count > 0:
                    print(f"\n  [{comp_type}] {comp.Name} ({line_count} baris)")
                    print(f"  {'─' * 60}")
                    # Tampilkan kode (max 30 baris)
                    lines = code.split('\n')
                    for i, line in enumerate(lines[:30]):
                        print(f"    {line.rstrip()}")
                    if len(lines) > 30:
                        print(f"    ... ({len(lines) - 30} baris lagi)")
        except Exception as e:
            print(f"  Error membaca VBA: {e}")
        
        # ===== 3. NAMED RANGES =====
        print(f"\n{'=' * 70}")
        print("3. NAMED RANGES")
        print("=" * 70)
        
        try:
            for i in range(1, wb.Names.Count + 1):
                name = wb.Names(i)
                nr_info = {
                    "name": name.Name,
                    "refers_to": name.RefersTo,
                    "visible": name.Visible
                }
                report["named_ranges"].append(nr_info)
                vis = "" if name.Visible else " [Hidden]"
                print(f"  {name.Name:<40} -> {name.RefersTo}{vis}")
        except Exception as e:
            print(f"  Error: {e}")
        
        if not report["named_ranges"]:
            print("  (Tidak ada named ranges)")
        
        # ===== 4. RINGKASAN & REKOMENDASI =====
        print(f"\n{'=' * 70}")
        print("4. RINGKASAN")
        print("=" * 70)
        print(f"  Total sheet: {wb.Sheets.Count}")
        print(f"  Total formula (sampled): {total_formulas}")
        print(f"  Total merged areas (sampled): {total_merged}")
        print(f"  VBA modules dengan kode: {sum(1 for m in report['vba_modules'] if m['line_count'] > 0)}")
        print(f"  Named ranges: {len(report['named_ranges'])}")
        
        print(f"\n  Formula yang digunakan:")
        for func, count in sorted(formula_types.items(), key=lambda x: -x[1]):
            print(f"    {func:<20} : {count} kali")
        
        # ===== 5. DETAIL FORMULA PER SHEET =====
        print(f"\n{'=' * 70}")
        print("5. SAMPLE FORMULA PER SHEET")
        print("=" * 70)
        
        for sheet in report["sheets"]:
            if sheet["formulas_sample"]:
                print(f"\n  Sheet: {sheet['name']}")
                for f in sheet["formulas_sample"]:
                    print(f"    {f['cell']}: {f['formula']}")
            if sheet["issues"]:
                print(f"  ISSUES:")
                for issue in sheet["issues"][:5]:
                    print(f"    ! {issue}")
        
        # ===== 6. ANALISA CROSS-SHEET REFERENCES =====
        print(f"\n{'=' * 70}")
        print("6. CROSS-SHEET REFERENCES")
        print("=" * 70)
        
        cross_refs = defaultdict(set)
        for sheet in report["sheets"]:
            for f in sheet["formulas_sample"]:
                formula = f["formula"]
                # Cari referensi ke sheet lain (pola 'SheetName'!)
                import re
                refs = re.findall(r"'([^']+)'!", formula)
                for ref in refs:
                    cross_refs[sheet["name"]].add(ref)
        
        for source, targets in cross_refs.items():
            print(f"  {source} -> referensi ke: {', '.join(targets)}")
        
        # Simpan report
        report_path = os.path.join(os.path.dirname(filepath), "excel_analysis_report.json")
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2, default=str)
        print(f"\n[OK] Report detail disimpan ke: {report_path}")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    target = r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026.xlsm"
    analyze_excel(target)
