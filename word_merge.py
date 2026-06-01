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
  python word_merge.py buka    <word_path> <excel_path> <sheet_name>
  python word_merge.py print   <word_path> <excel_path> <sheet_name>
  python word_merge.py pdf     <word_path> <excel_path> <sheet_name> <pdf_name>
  python word_merge.py printer <word_path> <excel_path> <sheet_name> <printer_name> [from_page] [to_page]
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
    s = str(value)
    # Bersihkan line break Excel (\r \n \r\n + literal _x000D_) → spasi tunggal.
    # Excel simpan multiline sebagai CR → muncul "_x000D_" mentah di Word saat merge.
    s = s.replace("_x000D_", "\n").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\s*\n\s*", " ", s).strip()
    return s


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

        # Word mail-merge menomori kolom dengan nama duplikat: occurrence ke-2 dst
        # dapat suffix "1", "2", ... (mis. Hari, Hari1, Hari2). Replikasi agar
        # MERGEFIELD seperti "Harga_Penawaran1"/"Hari1" ketemu saat re-merge.
        seen = {}
        for header, value in zip(headers, values):
            if header:
                header = str(header).strip()
                val = format_value(value)
                normalized = normalize_field_name(header)
                # nomori per nama-ternormalisasi (sesuai perilaku Word data source):
                # occurrence pertama = nama polos, ke-2 dst = suffix 1,2,...
                # pakai setdefault agar occurrence PERTAMA menang (match Word base).
                n = seen.get(normalized, 0)
                if n == 0:
                    data.setdefault(header, val)
                    if normalized != header:
                        data.setdefault(normalized, val)
                else:
                    data.setdefault(f"{normalized}{n}", val)
                    if header != normalized:
                        data.setdefault(f"{header}{n}", val)
                seen[normalized] = n + 1
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


def _replace_merge_fields(wdDoc, data):
    """Replace semua MERGEFIELD di wdDoc dgn nilai dari data + apply format switch.
    Field di-Unlink jadi teks statis. Loop backwards supaya index aman."""
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


def merge_word(word_path, data, mode="buka", pdf_name=""):
    import pythoncom
    import win32com.client

    # Mode bapljkk: copy template -> (Merged), replace MERGEFIELD dari Excel, baru export.
    # (sebelumnya buka ReadOnly tanpa merge -> PDF tampil cached hasil merge lama/template)
    if mode in ("pdf_bapljkk", "printer_bapljkk"):
        _word_path_win = os.path.abspath(os.path.normpath(word_path))
        _folder = os.path.dirname(_word_path_win)
        _base_b, _ext_b = os.path.splitext(os.path.basename(_word_path_win))
        if _ext_b.lower() not in (".docx", ".docm"):
            _ext_b = ".docx"
        _merged_b = os.path.join(_folder, f"{_base_b} (Merged){_ext_b}")
        shutil.copy2(_word_path_win, _merged_b)
        pythoncom.CoInitialize()
        wdApp = win32com.client.DispatchEx("Word.Application")
        wdApp.DisplayAlerts = 0
        wdApp.Visible = False
        try:
            wdDoc = wdApp.Documents.Open(
                FileName=_merged_b,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Visible=False,
            )
            # re-merge field dari data Excel (satu_data) -> PDF selalu fresh
            if data:
                _replace_merge_fields(wdDoc, data)
                wdDoc.Save()
            if mode == "pdf_bapljkk":
                _kode_pljkk = pdf_name if pdf_name else "PL"
                _xlsm_path = None
                try:
                    import glob as _glob_pl
                    _xlsm_pl = _glob_pl.glob(os.path.join(_folder, "*.xlsm"))
                    if _xlsm_pl:
                        _xlsm_path = os.path.normpath(_xlsm_pl[0])
                        _xl_pl = win32com.client.DispatchEx("Excel.Application")
                        _xl_pl.Visible = False
                        _wb_pl = _xl_pl.Workbooks.Open(_xlsm_path, ReadOnly=True)
                        _ku_pl = str(_wb_pl.Sheets("@ Master Data").Range("G2").Value).strip()
                        _wb_pl.Close(False)
                        _xl_pl.Quit()
                        if _ku_pl and _ku_pl not in ("", "None", "null"):
                            _kode_pljkk = _ku_pl
                except Exception:
                    pass
                _pdf_path = os.path.join(_folder, f"BA_PLJKK_{_kode_pljkk}.pdf")
                _tmp_word = _pdf_path + "_tmpword.pdf"
                _tmp_72   = _pdf_path + "_tmp72.pdf"
                # Export Word -> tmp
                if wdDoc.Sections.Count >= 3:
                    _start = wdDoc.Sections(3).Range.Start
                else:
                    _start = wdDoc.Content.Start
                _rng = wdDoc.Range(_start, wdDoc.Content.End)
                _rng.ExportAsFixedFormat(OutputFileName=_tmp_word, ExportFormat=17)
                # Export sheet 7.2 Dengan Nego dari Excel -> tmp (jika ada)
                _has_72 = False
                if _xlsm_path:
                    try:
                        _xl72 = win32com.client.DispatchEx("Excel.Application")
                        _xl72.Visible = False
                        _wb72 = _xl72.Workbooks.Open(_xlsm_path, ReadOnly=True)
                        _ws72 = None
                        try:
                            _ws72 = _wb72.Sheets("7.2 Dengan Nego")
                        except Exception:
                            pass
                        if _ws72 is not None:
                            _ws72.ExportAsFixedFormat(
                                Type=0,  # xlTypePDF
                                Filename=_tmp_72,
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                            _has_72 = True
                        _wb72.Close(False)
                        _xl72.Quit()
                    except Exception:
                        try:
                            _xl72.Quit()
                        except Exception:
                            pass
                # Gabung PDF: word_part1 + sheet72(2x) + word_part2
                # Cari halaman "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI TEKNIS DAN BIAYA" di word PDF
                if _has_72:
                    try:
                        import pdfplumber
                        from pypdf import PdfWriter, PdfReader
                        _rdr_word = PdfReader(_tmp_word)
                        _n_word = len(_rdr_word.pages)
                        # Cari halaman pertama daftar hadir klarifikasi nego
                        _split_page = _n_word  # default: tidak ketemu = append di akhir
                        with pdfplumber.open(_tmp_word) as _plb:
                            for _pi, _pp in enumerate(_plb.pages):
                                _txt = (_pp.extract_text() or "").upper()
                                if "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" in _txt:
                                    _split_page = _pi  # 0-based
                                    break
                        _rdr_72 = PdfReader(_tmp_72)
                        _writer = PdfWriter()
                        # Bagian 1: word halaman 0.._split_page-1
                        for _pi in range(_split_page):
                            _writer.add_page(_rdr_word.pages[_pi])
                        # Sheet 7.2: 2x
                        for _ in range(2):
                            for _pi in range(len(_rdr_72.pages)):
                                _writer.add_page(_rdr_72.pages[_pi])
                        # Bagian 2: word halaman _split_page..akhir
                        for _pi in range(_split_page, _n_word):
                            _writer.add_page(_rdr_word.pages[_pi])
                        with open(_pdf_path, "wb") as _fout:
                            _writer.write(_fout)
                    except Exception as _merge_err:
                        # Fallback: pakai word PDF saja
                        import shutil as _sh
                        _sh.copy2(_tmp_word, _pdf_path)
                    finally:
                        for _tf in (_tmp_word, _tmp_72):
                            try:
                                os.remove(_tf)
                            except Exception:
                                pass
                else:
                    # Tidak ada sheet 7.2, rename tmp -> final
                    import shutil as _sh
                    _sh.move(_tmp_word, _pdf_path)
                show_success(_pdf_path)
            elif mode == "printer_bapljkk":
                _printer_name = pdf_name
                wdApp.ActivePrinter = _printer_name
                if wdDoc.Sections.Count >= 3:
                    _start = wdDoc.Sections(3).Range.Start
                else:
                    _start = wdDoc.Content.Start
                _rng = wdDoc.Range(_start, wdDoc.Content.End)
                _rng.Select()
                wdDoc.PrintOut(Background=False, Range=1)  # wdPrintSelection=1
                time.sleep(3)
                show_print_success(_printer_name)
            wdDoc.Close(False)
        except Exception as e:
            show_error(f"Error cetak BA PLJKK:\n{e}")
        finally:
            wdApp.Quit()
            pythoncom.CoUninitialize()
            try:
                os.remove(_merged_b)
            except Exception:
                pass
        return

    folder = os.path.dirname(word_path)
    base, ext = os.path.splitext(os.path.basename(word_path))
    if ext.lower() not in (".docx", ".docm"):
        ext = ".docx"
    copy_path = os.path.join(folder, f"{base} (Merged){ext}")

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
        # Untuk mode PDF: Word tetap hidden (ScreenUpdating cukup untuk layout calculation)
        if "1. Full Dokumen BA PK" in os.path.basename(word_path):
            wdApp.ScreenUpdating = True
            if mode in ("buka", "print"):
                wdApp.Visible = True
                wdApp.WindowState = 2
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

        elif mode == "printer":
            # Direct print ke printer fisik (tanpa dialog)
            printer_name = pdf_name  # arg ke-5 dipakai sebagai nama printer
            from_page = int(sys.argv[6]) if len(sys.argv) > 6 else 0
            to_page = int(sys.argv[7]) if len(sys.argv) > 7 else 0

            wdDoc.Save()
            wdApp.ScreenUpdating = True
            wdApp.Visible = False

            try:
                # Set printer tujuan
                wdApp.ActivePrinter = printer_name

                # PrintOut dengan atau tanpa page range
                if from_page > 0 and to_page > 0:
                    wdDoc.PrintOut(
                        Background=False,
                        Range=3,   # wdPrintFromTo
                        From=str(from_page),
                        To=str(to_page),
                    )
                else:
                    wdDoc.PrintOut(Background=False)

                # Tunggu spooler selesai
                time.sleep(3)
                show_print_success(printer_name)
            except Exception as print_err:
                show_error(f"Gagal print ke {printer_name}:\n{print_err}")

            wdDoc.Close(False)
            if new_instance:
                wdApp.Quit()
            return  # skip cleanup di bawah

        elif mode == "printer_bapljkk":
            # Print Section 3 s/d akhir ke printer fisik (skip Reviu)
            printer_name = pdf_name  # arg ke-5 = nama printer

            wdDoc.Save()
            wdApp.ScreenUpdating = True
            wdApp.Visible = False

            try:
                wdApp.ActivePrinter = printer_name
                if wdDoc.Sections.Count >= 3:
                    _start = wdDoc.Sections(3).Range.Start
                else:
                    _start = wdDoc.Content.Start
                _rng = wdDoc.Range(_start, wdDoc.Content.End)
                _rng.Select()
                wdDoc.PrintOut(Background=False, Range=1)  # wdPrintSelection=1
                time.sleep(3)
                show_print_success(printer_name)
            except Exception as print_err:
                show_error(f"Gagal print ke {printer_name}:\n{print_err}")

            wdDoc.Close(False)
            if new_instance:
                wdApp.Quit()
            return

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
                show_success(pdf_path)
            elif mode == "pdf_bareviu_pl":
                pdf_path = os.path.join(folder, f"BA_REVIU_PL_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=1,
                    To=3,
                )
                show_success(pdf_path)
            elif mode == "pdf_bapljkk":
                # Export Section 3 s/d akhir (skip Section 1+2 = Reviu DPP)
                # Pakai Range agar tidak perlu tahu nomor halaman (robust terhadap perubahan isi Reviu)
                _kode_pljkk = safe_name
                try:
                    import glob as _glob_pl
                    _xlsm_pl = _glob_pl.glob(os.path.join(folder, "*.xlsm"))
                    if _xlsm_pl:
                        _xl_pl = win32com.client.DispatchEx("Excel.Application")
                        _xl_pl.Visible = False
                        _wb_pl = _xl_pl.Workbooks.Open(_xlsm_pl[0], ReadOnly=True)
                        _ku_pl = str(_wb_pl.Sheets("@ Master Data").Range("G2").Value).strip()
                        _wb_pl.Close(False)
                        _xl_pl.Quit()
                        if _ku_pl and _ku_pl not in ("", "None", "null"):
                            _kode_pljkk = _ku_pl
                except Exception:
                    pass
                pdf_path = os.path.join(folder, f"BA_PLJKK_{_kode_pljkk}.pdf")
                # File baru (BA-only) hanya 1 section; file lama (gabung Reviu) BA mulai Section 3
                if wdDoc.Sections.Count >= 3:
                    _start = wdDoc.Sections(3).Range.Start
                else:
                    _start = wdDoc.Content.Start
                _rng = wdDoc.Range(_start, wdDoc.Content.End)
                _rng.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)
                show_success(pdf_path)
            elif mode == "pdf_revaluasi":
                pdf_path = os.path.join(folder, f"REvaluasi_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=30,
                    To=37,
                )
                show_success(pdf_path)
            elif mode == "pdf_all":
                # Ambil kode_unik dari @ Master Data!G2
                _ku_all = safe_name
                try:
                    import glob as _glob_all
                    _xlsm_all = _glob_all.glob(os.path.join(folder, "*.xlsm"))
                    if _xlsm_all:
                        _xl_all = win32com.client.DispatchEx("Excel.Application")
                        _xl_all.Visible = False
                        _wb_all = _xl_all.Workbooks.Open(_xlsm_all[0], ReadOnly=True)
                        _ku_val = str(_wb_all.Sheets("@ Master Data").Range("G2").Value).strip()
                        _wb_all.Close(False)
                        _xl_all.Quit()
                        if _ku_val and _ku_val not in ("", "None", "null"):
                            _ku_all = _ku_val
                except Exception:
                    pass
                pdf_path = os.path.join(folder, f"Isi_Reviu_DPP_{_ku_all}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
                show_success(pdf_path)
            elif mode == "pdf_dokpil":
                # Ambil kode_unik dari Excel @ Master Data!G2 (cari xlsm di folder yang sama)
                kode_unik_pdf = ""
                try:
                    import glob as _glob2
                    _xlsm_list = _glob2.glob(os.path.join(folder, "*.xlsm"))
                    if _xlsm_list:
                        _xl2 = win32com.client.DispatchEx("Excel.Application")
                        _xl2.Visible = False
                        _wb2 = _xl2.Workbooks.Open(_xlsm_list[0], ReadOnly=True)
                        kode_unik_pdf = str(_wb2.Sheets("@ Master Data").Range("G2").Value).strip()
                        _wb2.Close(False)
                        _xl2.Quit()
                except Exception:
                    pass
                _pdf_suffix = kode_unik_pdf if kode_unik_pdf and kode_unik_pdf not in ("", "None", "null") else safe_name
                pdf_path = os.path.join(folder, f"dokpil_{_pdf_suffix}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
                show_success(pdf_path)
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
                    ws_nego.PageSetup.Orientation = 2  # xlLandscape
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
                
                # Sisipkan Nego Pertama (Setelah Hal 17) — semua halaman Excel
                for ep in reader_excel.pages:
                    writer.add_page(ep)

                # Word 18-26 itu index range(11, 20)
                for i in range(11, 20):
                    if i < len(reader_word.pages):
                        writer.add_page(reader_word.pages[i])

                # Sisipkan Nego Kedua (Setelah Hal 26) — semua halaman Excel
                for ep in reader_excel.pages:
                    writer.add_page(ep)
                    
                # Word 27-29 itu index range(20, end)
                for i in range(20, len(reader_word.pages)):
                    writer.add_page(reader_word.pages[i])
                    
                final_pdf_path = _safe_write_pdf(writer, final_pdf_path)
                show_success(final_pdf_path)

            elif mode == "pdf_pembuktian_timpang":
                import tempfile
                from pypdf import PdfReader, PdfWriter
                
                final_pdf_path = os.path.join(folder, f"BA_Pembuktian_Timpang_{safe_name}.pdf")
                temp_dir = tempfile.mkdtemp()
                temp_word_pdf = os.path.join(temp_dir, "temp_word.pdf")
                temp_nego_pdf = os.path.join(temp_dir, "temp_nego.pdf")
                temp_timpang_pdf = os.path.join(temp_dir, "temp_timpang.pdf")
                
                # 1. Export Word Halaman 7-44 (44 = tanda terima timpang)
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=temp_word_pdf,
                    ExportFormat=17,
                    Range=3,
                    From=7,
                    To=44,
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
                    ws_nego.PageSetup.Orientation = 2  # xlLandscape
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
                    for p in reader_nego.pages:
                        writer.add_page(p)

                def add_timpang():
                    for p in reader_timpang.pages:
                        writer.add_page(p)
                
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
                
                # Part 11: Word 27 saja (Index 20)
                add_wp(20, 21)

                # Part 12: Tanda Terima Timpang (Word 44 = Index 37), 2x
                add_wp(37, 38)
                add_wp(37, 38)
                
                final_pdf_path = _safe_write_pdf(writer, final_pdf_path)
                show_success(final_pdf_path)

            else:
                pdf_path = os.path.join(folder, f"Undangan_{safe_name}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=1,
                    To=2,
                )
                show_success(pdf_path)

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


def _safe_write_pdf(writer, target_path):
    """Write PDF, fallback ke suffix _v2/_v3/... jika file di-lock proses lain."""
    path = target_path
    for attempt in range(5):
        try:
            with open(path, 'wb') as f:
                writer.write(f)
            return path
        except PermissionError:
            base, ext = os.path.splitext(target_path)
            path = f"{base}_v{attempt + 2}{ext}"
    raise PermissionError(f"Gagal tulis PDF setelah 5 percobaan: {target_path}")


def show_success(pdf_path):
    """Notifikasi popup setelah PDF selesai dibuat."""
    try:
        import ctypes
        filename = os.path.basename(pdf_path)
        ctypes.windll.user32.MessageBoxW(
            0, f"PDF berhasil dibuat:\n{filename}", "Export PDF Selesai", 0x40
        )
    except:
        pass


def show_print_success(printer_name):
    """Notifikasi popup setelah print dikirim ke spooler."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, f"Dokumen dikirim ke antrian print:\n{printer_name}\n\n"
               f"Pastikan printer menyala untuk mencetak.",
            "Print Dikirim", 0x40
        )
    except:
        pass


def run_merge_mode_pl(folder_path: str, excel_path: str) -> list:
    """
    Merge semua Word template PL di folder_path menggunakan data dari excel_path.
    Loop over WORD_SHEET_MAP_PL: (word_filename, sheet_name).
    Return: list hasil per file — {"file": str, "sukses": bool, "pesan": str}
    """
    import glob as _glob
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    try:
        from config import WORD_SHEET_MAP_PL
    except ImportError:
        WORD_SHEET_MAP_PL = [
            ("5. BA PLJKK - Template.docx",               "satu_data"),
            ("1. Reviu DPP PLJKK - Template.docx",        "list_reviu"),
            ("3. Dokpil Full PLJKK - Template.docx",      "list_dokpil"),
        ]

    results = []
    for word_filename, sheet_name in WORD_SHEET_MAP_PL:
        # Cari file — bisa Template atau sudah direname
        base = word_filename.replace(" - Template", "").replace(" - Template", "")
        candidates = _glob.glob(os.path.join(folder_path, word_filename))
        if not candidates:
            # Coba nama tanpa "- Template"
            stem = os.path.splitext(word_filename)[0].replace(" - Template", "").strip()
            ext  = os.path.splitext(word_filename)[1]
            candidates = _glob.glob(os.path.join(folder_path, f"{stem}*{ext}"))
        if not candidates:
            results.append({"file": word_filename, "sukses": False, "pesan": "File tidak ditemukan di folder"})
            continue

        word_path = candidates[0]
        data = read_excel_data(excel_path, sheet_name)
        if data is None:
            results.append({"file": os.path.basename(word_path), "sukses": False, "pesan": f"Gagal baca sheet {sheet_name}"})
            continue

        try:
            merge_word(word_path, data, mode="buka", pdf_name="")
            results.append({"file": os.path.basename(word_path), "sukses": True, "pesan": "OK"})
        except Exception as e:
            results.append({"file": os.path.basename(word_path), "sukses": False, "pesan": str(e)})

    return results


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python word_merge.py <mode> <word_path> <excel_path> <sheet_name>")
        print("  mode: buka | print | pdf | pdf_bareviu | pdf_bareviu_pl")
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
