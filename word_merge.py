"""
word_merge.py - Merge Word template dengan Excel data
=====================================================
Dipanggil dari VBA Excel via Shell (proses terpisah, Excel tidak hang).

Strategy:
  1. Baca data dari Excel yang sedang terbuka (GetObject - cepat)
  2. Copy template ke file (Merged).docx
  3. Buka copy, ganti semua MERGEFIELD dengan data
  4. Tampilkan Word / print / export PDF

Usage:
  python word_merge.py buka  <word_path> <excel_path> <sheet_name>
  python word_merge.py print <word_path> <excel_path> <sheet_name>
  python word_merge.py pdf   <word_path> <excel_path> <sheet_name> <pdf_name>
"""
import sys
import os
import time
import re
import datetime
import shutil


def format_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def normalize_field_name(name):
    s = str(name).strip()
    s = s.replace(" ", "_")
    s = re.sub(r"[^A-Za-z0-9_]", "", s)
    return s


def read_excel_data(excel_path, sheet_name):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    data = {}

    try:
        wb = win32com.client.GetObject(excel_path)
        ws = wb.Sheets(sheet_name)

        cols = ws.UsedRange.Columns.Count
        headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, cols)).Value[0]
        values = ws.Range(ws.Cells(2, 1), ws.Cells(2, cols)).Value[0]

        for header, value in zip(headers, values):
            if header:
                header = str(header).strip()
                val = format_value(value)
                data[header] = val
                normalized = normalize_field_name(header)
                if normalized != header:
                    data[normalized] = val

    except Exception:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel_path, read_only=True, data_only=True, keep_links=False)
            if sheet_name not in wb.sheetnames:
                wb.close()
                show_error(f"Sheet '" + sheet_name + "' tidak ditemukan di Excel.")
                return None
            ws = wb[sheet_name]
            headers = [c.value for c in ws[1]]
            values = [c.value for c in ws[2]]
            wb.close()
            for header, value in zip(headers, values):
                if header:
                    header = str(header).strip()
                    val = format_value(value)
                    data[header] = val
                    normalized = normalize_field_name(header)
                    if normalized != header:
                        data[normalized] = val
        except Exception as e:
            show_error(f"Error baca Excel:\n{e}")
            return None
    finally:
        pythoncom.CoUninitialize()

    return data


def cleanup_blank_pages(doc):
    doc.Repaginate()

    for sec_idx in range(1, doc.Sections.Count):
        try:
            sec = doc.Sections(sec_idx)
            sec_range = sec.Range
            last_page = sec_range.Information(3)
            first_page = sec_range.Information(1)

            if last_page <= first_page:
                continue

            para_count = sec_range.Paragraphs.Count
            for p in range(para_count, 0, -1):
                para = sec_range.Paragraphs(p)
                para_page = para.Range.Information(3)
                para_text = para.Range.Text.strip()

                if not para_text and para_page == last_page:
                    para.Range.Font.Size = 1
                    para.Format.SpaceBefore = 0
                    para.Format.SpaceAfter = 0
                    para.Format.LineSpacingRule = 0
                elif para_text:
                    break
        except:
            pass

    doc.Repaginate()


def merge_word(word_path, data, mode="buka", pdf_name=""):
    import pythoncom
    import win32com.client

    folder = os.path.dirname(word_path)
    base = os.path.splitext(os.path.basename(word_path))[0]
    copy_path = os.path.join(folder, f"{base} (Merged).docx")

    # Copy template ke (Merged) - template asli tidak diubah
    shutil.copy2(word_path, copy_path)

    pythoncom.CoInitialize()
    wdApp = None
    new_instance = False

    try:
        wdApp = win32com.client.DispatchEx("Word.Application")
        new_instance = True

        wdApp.DisplayAlerts = 0
        wdApp.Visible = False

        wdDoc = wdApp.Documents.Open(
            FileName=copy_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False
        )

        # Replace MERGEFIELD fields (loop backwards)
        field_count = wdDoc.Fields.Count
        for i in range(field_count, 0, -1):
            try:
                field = wdDoc.Fields(i)
                code_text = field.Code.Text.strip()
                if code_text.upper().startswith("MERGEFIELD"):
                    parts = code_text.split()
                    if len(parts) >= 2:
                        fname = parts[1].strip('"').strip()
                        val = None
                        if fname in data:
                            val = data[fname]
                        else:
                            norm = normalize_field_name(fname)
                            if norm in data:
                                val = data[norm]

                        if val is not None:
                            field.Result.Text = str(val)
                        else:
                            field.Result.Text = ""
                        field.Unlink()
            except:
                pass

        # Cleanup blank pages
        # cleanup_blank_pages(wdDoc)

        # Simpan dan Tampilkan hanya jika bukan mode PDF
        if mode in ("buka", "print"):
            wdDoc.Save()
            wdApp.ScreenUpdating = True
            wdApp.Visible = True
            wdDoc.Windows(1).Visible = True
            wdDoc.Activate()
            wdDoc.Repaginate()
            time.sleep(1)

            try:
                import ctypes
                hwnd = ctypes.windll.user32.FindWindowW("OpusApp", None)
                if hwnd:
                    ctypes.windll.user32.ShowWindow(hwnd, 3)
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except:
                wdApp.WindowState = 1

        if mode == "print":
            time.sleep(1)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys("^p", 0)

        elif mode.startswith("pdf"):
            safe_name = pdf_name if pdf_name else "000"
            # wdDoc.Repaginate()  # Mencegah FREEZE saat export PDF

            if mode == "pdf_bareviu":
                pdf_path = os.path.join(folder, f"BA_REVIU_DPP_{safe_name}.pdf")
                from_page = 3
                to_page = 6
            else:
                pdf_path = os.path.join(folder, f"Undangan_{safe_name}.pdf")
                from_page = 1
                to_page = 2

            wdDoc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,
                Range=3,
                From=from_page,
                To=to_page,
            )
            wdDoc.Close(False)
            if new_instance:
                wdApp.Quit()

    except Exception as e:
        if wdApp:
            try:
                wdApp.ScreenUpdating = True
                wdApp.Visible = True
            except: pass
        show_error(f"Error saat merge:\n{e}")
    finally:
        pythoncom.CoUninitialize()


def show_error(msg):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Word Merge Error", 0x10)
    except:
        print(f"ERROR: {msg}")


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python word_merge.py <mode> <word_path> <excel_path> <sheet_name>")
        print("  mode: buka | print | pdf | pdf_bareviu")
        sys.exit(1)

    mode = sys.argv[1]
    word_path = sys.argv[2]
    excel_path = sys.argv[3]
    sheet_name = sys.argv[4]
    pdf_name = sys.argv[5] if len(sys.argv) > 5 else ""

    data = read_excel_data(excel_path, sheet_name)

    if data is None:
        sys.exit(1)

    merge_word(word_path, data, mode, pdf_name)
