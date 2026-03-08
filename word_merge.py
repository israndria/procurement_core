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
    """Baca data Excel via openpyxl (copy ke temp dulu karena file mungkin terkunci oleh Excel)."""
    import tempfile
    from openpyxl import load_workbook

    data = {}
    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, os.path.basename(excel_path))

    try:
        shutil.copy2(excel_path, temp_path)
        wb = load_workbook(temp_path, read_only=True, data_only=True, keep_links=False)

        if sheet_name not in wb.sheetnames:
            wb.close()
            show_error(f"Sheet '{sheet_name}' tidak ditemukan di Excel.")
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
        try:
            os.remove(temp_path)
            os.rmdir(temp_dir)
        except:
            pass

    return data


def cleanup_blank_pages(doc):
    # Metode Ringan & Cepat: Shrink paragraf kosong di seluruh dokumen (tanpa batas halaman).
    try:
        for i, para in enumerate(doc.Paragraphs):
            try:
                txt = para.Range.Text.replace('\r', '').replace('\n', '')
                # Hanya panggil .Information(3) (yang butuh komputasi layout berat) JIKA paragraf kosong.
                # Ini mempercepat script 10x lipat dan mencegah hang/timeout dari RPC Server!
                if not txt.strip():
                    pg_num = para.Range.Information(3) # wdActiveEndPageNumber
                    if pg_num > 1:
                        # Shrink blank space (Enter Berlebihan di Word)
                        para.Range.Font.Size = 1
                        para.Format.SpaceBefore = 0
                        para.Format.SpaceAfter = 0
                        para.Format.LineSpacingRule = 0
            except:
                pass
    except Exception as e:
        print(f"Warning saat cleanup: {e}")
        pass
    
    return


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
                            val = str(val)
                            # Terapkan format switch Word (\* Upper, \* Lower, \* FirstCap)
                            format_str = " ".join(parts[2:]).upper()
                            if "UPPER" in format_str:
                                val = val.upper()
                            elif "LOWER" in format_str:
                                val = val.lower()
                            elif "FIRSTCAP" in format_str:
                                val = val.capitalize()
                            field.Result.Text = val
                        else:
                            field.Result.Text = ""
                        field.Unlink()
            except:
                pass

        # Cleanup blank pages HANYA untuk "1. Full Dokumen BA PK"
        # File "2. Isi Reviu" dan "3. Dokpil" dikecualikan karena strukturnya berbeda dan bisa berantakan
        if "1. Full Dokumen BA PK" in os.path.basename(word_path):
            wdApp.ScreenUpdating = True
            wdApp.Visible = True 
            wdApp.WindowState = 2 # 2=wdWindowStateMinimize (Tampil ke layar tetapi diforced Minimize)
            cleanup_blank_pages(wdDoc)

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

            if mode == "pdf_bareviu":
                pdf_path = os.path.join(folder, f"BA_REVIU_DPP_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=3,
                    To=6,
                )
            elif mode == "pdf_revaluasi":
                pdf_path = os.path.join(folder, f"REvaluasi_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=30,
                    To=37,
                )
            elif mode == "pdf_all":
                pdf_path = os.path.join(folder, f"Isi_Reviu_DPP_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
            elif mode == "pdf_dokpil":
                pdf_path = os.path.join(folder, f"DOKPIL_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
            elif mode == "pdf_pembuktian":
                import tempfile
                from pypdf import PdfReader, PdfWriter
                
                final_pdf_path = os.path.join(folder, f"BA_Pembuktian_Nego_{safe_name}.pdf")
                temp_dir = tempfile.mkdtemp()
                temp_word_pdf = os.path.join(temp_dir, "temp_word.pdf")
                temp_excel_pdf = os.path.join(temp_dir, "temp_excel.pdf")
                
                # 1. Export Word Halaman 7-29
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=temp_word_pdf,
                    ExportFormat=17,
                    Range=3,
                    From=7,
                    To=29,
                )
                
                # 2. Export Excel Sheet 7.2
                xlApp = None
                wb_xl = None
                try:
                    xlApp = win32com.client.DispatchEx("Excel.Application")
                    xlApp.Visible = False
                    xlApp.DisplayAlerts = False
                    wb_xl = xlApp.Workbooks.Open(excel_path, ReadOnly=True)
                    ws_nego = wb_xl.Sheets("7.2 Dengan Nego")
                    ws_nego.PageSetup.Orientation = 2 # xlLandscape
                    ws_nego.PageSetup.Zoom = False
                    ws_nego.PageSetup.FitToPagesWide = 1
                    ws_nego.PageSetup.FitToPagesTall = 1
                    ws_nego.ExportAsFixedFormat(0, temp_excel_pdf) # 0 = xlTypePDF
                finally:
                    if wb_xl:
                        try: wb_xl.Close(False)
                        except: pass
                    if xlApp:
                        try: xlApp.Quit()
                        except: pass
                
                # 3. Gabungkan dengan PyPDF2
                writer = PdfWriter()
                reader_word = PdfReader(temp_word_pdf)
                reader_excel = PdfReader(temp_excel_pdf)
                
                # Word 7-17 itu index range(0, 11) karena From=7 adalah index 0 di pdf word
                for i in range(11):
                    if i < len(reader_word.pages):
                        writer.add_page(reader_word.pages[i])
                
                # Sisipkan Nego Pertama (Setelah Hal 17)
                if len(reader_excel.pages) > 0:
                    writer.add_page(reader_excel.pages[0])
                
                # Word 18-26 itu index range(11, 20)
                for i in range(11, 20):
                    if i < len(reader_word.pages):
                        writer.add_page(reader_word.pages[i])
                    
                # Sisipkan Nego Kedua (Setelah Hal 26)
                if len(reader_excel.pages) > 0:
                    writer.add_page(reader_excel.pages[0])
                    
                # Word 27-29 itu index range(20, end)
                for i in range(20, len(reader_word.pages)):
                    writer.add_page(reader_word.pages[i])
                    
                with open(final_pdf_path, 'wb') as fd_out:
                    writer.write(fd_out)
                    
            elif mode == "pdf_pembuktian_timpang":
                import tempfile
                from pypdf import PdfReader, PdfWriter
                
                final_pdf_path = os.path.join(folder, f"BA_Pembuktian_Timpang_{safe_name}.pdf")
                temp_dir = tempfile.mkdtemp()
                temp_word_pdf = os.path.join(temp_dir, "temp_word.pdf")
                temp_nego_pdf = os.path.join(temp_dir, "temp_nego.pdf")
                temp_timpang_pdf = os.path.join(temp_dir, "temp_timpang.pdf")
                
                # 1. Export Word Halaman 7-43 (karena butuh nyampe 43)
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=temp_word_pdf,
                    ExportFormat=17,
                    Range=3,
                    From=7,
                    To=43,
                )
                
                # 2. Export Excel Sheets (Nego dan Timpang)
                xlApp = None
                wb_xl = None
                try:
                    xlApp = win32com.client.DispatchEx("Excel.Application")
                    xlApp.Visible = False
                    xlApp.DisplayAlerts = False
                    wb_xl = xlApp.Workbooks.Open(excel_path, ReadOnly=True)
                    
                    # Nego
                    ws_nego = wb_xl.Sheets("7.2 Dengan Nego")
                    ws_nego.PageSetup.Orientation = 2 # xlLandscape
                    ws_nego.PageSetup.Zoom = False
                    ws_nego.PageSetup.FitToPagesWide = 1
                    ws_nego.PageSetup.FitToPagesTall = 1
                    ws_nego.ExportAsFixedFormat(0, temp_nego_pdf)
                    
                    # Timpang
                    ws_timpang = wb_xl.Sheets("Klarifikasi Timpang Fix (2)")
                    ws_timpang.PageSetup.Orientation = 2 # xlLandscape
                    ws_timpang.PageSetup.Zoom = False
                    ws_timpang.PageSetup.FitToPagesWide = 1
                    ws_timpang.PageSetup.FitToPagesTall = 1
                    ws_timpang.ExportAsFixedFormat(0, temp_timpang_pdf)
                finally:
                    if wb_xl:
                        try: wb_xl.Close(False)
                        except: pass
                    if xlApp:
                        try: xlApp.Quit()
                        except: pass
                        
                # 3. Merajut semuanya
                writer = PdfWriter()
                reader_word = PdfReader(temp_word_pdf)
                reader_nego = PdfReader(temp_nego_pdf)
                reader_timpang = PdfReader(temp_timpang_pdf)
                
                def add_wp(start_idx, end_idx):
                    for i in range(start_idx, end_idx):
                        if i < len(reader_word.pages): writer.add_page(reader_word.pages[i])
                
                def add_nego():
                    if len(reader_nego.pages) > 0: writer.add_page(reader_nego.pages[0])
                    
                def add_timpang():
                    if len(reader_timpang.pages) > 0: writer.add_page(reader_timpang.pages[0])
                
                # Part 1: Word 7-16 (Index 0-9)
                add_wp(0, 10)
                
                # ------ PENYISIPAN PERTAMA ------
                # Sisipkan Word 38, 39 (Index 31, 32)
                add_wp(31, 33)
                # Sisipkan Timpang
                add_timpang()
                # Sisipkan Word 40 (Index 33)
                add_wp(33, 34)
                
                # Part 4: Word 17 (Index 10)
                add_wp(10, 11)
                # Part 5: Sisipkan Nego
                add_nego()
                
                # Part 6: Word 18-25 (Index 11-18)
                add_wp(11, 19)
                
                # ------ PENYISIPAN KEDUA ------
                # Sisipkan Word 41, 42 (Index 34, 35)
                add_wp(34, 36)
                # Sisipkan Timpang
                add_timpang()
                # Sisipkan Word 43 (Index 36)
                add_wp(36, 37)
                
                # Part 9: Word 26 (Index 19)
                add_wp(19, 20)
                # Part 10: Sisipkan Nego
                add_nego()
                
                # Part 11: Word 27-29 (Index 20-23)
                add_wp(20, 23)
                
                with open(final_pdf_path, 'wb') as fd_out:
                    writer.write(fd_out)

            else:
                pdf_path = os.path.join(folder, f"Undangan_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=1,
                    To=2,
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
