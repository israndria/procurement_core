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
import glob


def _safe_filename(s: str, max_len: int = 80) -> str:
    s = re.sub(r'[<>:"/\\|?*]', '', str(s)).strip().replace('\n',' ').replace('\r','')
    return s[:max_len] if s else 'Dokumen'


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
                normalized = normalize_field_name(header)
                val = format_value(value)
                # Formula Excel untuk slot peserta yang tidak terisi sering
                # menghasilkan angka 0. Untuk Word, slot kosong harus benar-
                # benar kosong agar baris peserta tidak tampil sebagai "0".
                if (normalized.startswith("Peserta_") or normalized.startswith("Alamat_")) and val == "0":
                    val = ""
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
        if sheet_name == "satu_data":
            _augment_ba_counts(data, temp_path)
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


def _set_data_aliases(data, field_name, value):
    """Set nama field asli + nama hasil normalisasi Word mail merge."""
    data[field_name] = value
    data[normalize_field_name(field_name)] = value


def _angka_kata(n):
    angka = {
        0: "Nol", 1: "Satu", 2: "Dua", 3: "Tiga", 4: "Empat",
        5: "Lima", 6: "Enam", 7: "Tujuh", 8: "Delapan", 9: "Sembilan",
    }
    return angka.get(int(n), str(int(n)))


def _augment_ba_counts(data, excel_copy_path):
    """Ambil hitungan BA dari Sheet 0 Input BA, bukan cache satu_data."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(excel_copy_path, read_only=True, data_only=True, keep_links=False)
        if "0. Input BA" not in wb.sheetnames:
            wb.close()
            return
        ws = wb["0. Input BA"]
        total = int(ws["C25"].value or 0)
        openable = int(ws["C26"].value or 0)
        incomplete = int(ws["C28"].value or 0)
        unreadable = int(ws["C29"].value or 0)
        complete = max(total - incomplete, 0)
        values = {
            "Ket 1": f"Terdapat jumlah dokumen penawaran keseluruhan sebanyak {total} ({_angka_kata(total)}) buah;",
            "Ket 2 a": f"Jumlah dokumen penawaran yang lengkap sebanyak {complete} ({_angka_kata(complete)}) buah;",
            "Ket 2 b": f"Jumlah dokumen penawaran yang tidak lengkap sebanyak {incomplete} ({_angka_kata(incomplete)}) buah;",
            "Ket 3 a": f"Jumlah dokumen penawaran yang dapat dibuka sebanyak {openable} ({_angka_kata(openable)}) buah;",
            "Ket 3 b": f"Jumlah dokumen penawaran yang tidak dapat dibuka sebanyak {unreadable} ({_angka_kata(unreadable)}) buah;",
        }
        for key, value in values.items():
            _set_data_aliases(data, key, value)
        wb.close()
    except Exception:
        pass


def _trim_blank_participant_rows(wdDoc):
    """Hapus baris peserta kosong dari tabel ringkasan BA Pembuktian."""
    try:
        for i in range(1, wdDoc.Tables.Count + 1):
            table = wdDoc.Tables(i)
            table_text = table.Range.Text.upper()
            if "NAMA PESERTA" not in table_text or "KETERANGAN" not in table_text:
                continue
            for r in range(table.Rows.Count, 1, -1):
                row = table.Rows(r)
                first = row.Cells(1).Range.Text.replace("\r", "").replace("\a", "").strip()
                second = row.Cells(2).Range.Text.replace("\r", "").replace("\a", "").strip()
                if re.fullmatch(r"[23]\.?", first) and (not second or second == "0"):
                    row.Delete()
    except Exception:
        pass


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


def _fit_path(folder, filename, max_total=240):
    """Word COM (ExportAsFixedFormat/SaveAs2) menolak path >255 char
    ('String is longer than 255 characters', wdmain11.chm 41873).
    Potong stem filename agar total path aman."""
    path = os.path.join(folder, filename)
    if len(path) <= max_total:
        return path
    stem, ext = os.path.splitext(filename)
    avail = max_total - len(folder) - len(ext) - 1  # -1 utk separator
    if avail < 10:
        avail = 10
    return os.path.join(folder, stem[:avail].rstrip() + ext)


def _strip_mailmerge_datasource(docx_path):
    """Hapus attachment mail merge (w:mailMerge di word/settings.xml) dari copy
    (Merged). Path Excel panjang bikin connection string/SQL >255 char sehingga
    Word error 41873 'String is longer than 255 characters' saat auto-connect
    data source di Documents.Open. Script ini merge field sendiri via COM,
    jadi attachment tidak diperlukan di file copy."""
    import re
    import zipfile
    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            names = zin.namelist()
            if "word/settings.xml" not in names:
                return
            settings = zin.read("word/settings.xml").decode("utf-8")
            new_settings = re.sub(
                r"<w:mailMerge>.*?</w:mailMerge>|<w:mailMerge\s*/>",
                "", settings, flags=re.DOTALL)
            if new_settings == settings:
                return
            items = [(n, zin.read(n)) for n in names]
        tmp = docx_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n, blob in items:
                zout.writestr(n, new_settings.encode("utf-8") if n == "word/settings.xml" else blob)
        os.replace(tmp, docx_path)
    except Exception:
        pass  # gagal strip -> lanjut; worst case error lama muncul lagi


def _set_field_result(field, val):
    if len(val) <= 255:
        field.Result.Text = val
        field.Unlink()
        return
    rng = field.Result
    field.Unlink()
    rng.Text = val


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
                    else:
                        val = ""
                    _set_field_result(field, val)
        except:
            pass


def _protect_signature_layout(wdDoc):
    """Jaga blok tanda tangan tetap utuh saat Word melakukan pagination."""
    for i in range(1, wdDoc.Tables.Count + 1):
        try:
            table = wdDoc.Tables(i)
            text = table.Range.Text.upper()
            if not any(marker in text for marker in (
                "DIREKTUR/PIMPINAN", "KELOMPOK KERJA PEMILIHAN"
            )):
                continue

            # Nama tenaga ahli/pimpinan bisa membuat blok ini tinggi. Jangan
            # izinkan Word memecah baris atau mendorong label dan nama ke page
            # berbeda; Word akan memindahkan blok utuh ke halaman berikutnya.
            for j in range(1, table.Rows.Count + 1):
                row = table.Rows(j)
                try:
                    row.AllowBreakAcrossPages = False
                    row.HeightRule = 0  # wdRowHeightAuto
                except Exception:
                    pass

            paragraphs = table.Range.Paragraphs
            for j in range(1, paragraphs.Count + 1):
                paragraph = paragraphs(j)
                paragraph.Range.ParagraphFormat.KeepTogether = True
                paragraph.Range.ParagraphFormat.KeepWithNext = j < paragraphs.Count
        except Exception:
            pass


def _sisip_2ba_pljkk(pdf_path, folder):
    """
    Sisip 2 file BA (Evaluasi /05/ + Hasil /07/) ke dalam BA_PLJKK final.

    Posisi (anchor by judul section, occurrence-aware):
      - BA Evaluasi  → setelah halaman judul 'DAFTAR HADIR PEMBUKTIAN KUALIFIKASI'
                       occurrence #1 (di antara 2 daftar hadir pembuktian).
      - BA Hasil     → setelah halaman judul 'DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI'
                       occurrence #1 (di antara 2 daftar hadir klarifikasi).

    File BA dicari via prefix '1. BA Evaluasi*.pdf' & '2. BA Hasil*.pdf' di folder.
    Best-effort: jika file BA tidak ada / anchor tidak ketemu, lewati tanpa error.
    Idempotent: jika anchor occurrence tidak cukup, BA terkait dilewati.
    """
    try:
        import glob as _glob
        import pdfplumber
        from pypdf import PdfReader, PdfWriter

        _ev = sorted(_glob.glob(os.path.join(folder, "1. BA Evaluasi*.pdf")))
        _hs = sorted(_glob.glob(os.path.join(folder, "2. BA Hasil*.pdf")))
        _ev_path = _ev[0] if _ev else None
        _hs_path = _hs[0] if _hs else None
        if not _ev_path and not _hs_path:
            return  # tidak ada file BA, lewati

        _rdr = PdfReader(pdf_path)
        _n = len(_rdr.pages)

        # Identifikasi halaman ANCHOR (judul section, bukan kop/teks berulang).
        # Valid jika: judul muncul di AWAL halaman (idx < 200, setelah kop dinas) DAN
        # halaman bukan "BERITA ACARA ..." (halaman BA punya judul section di bawah).
        _PEMB = "DAFTAR HADIR PEMBUKTIAN KUALIFIKASI"
        _KLAR = "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI"
        _pemb_pages = []    # index halaman daftar hadir pembuktian
        _klarif_pages = []  # index halaman daftar hadir klarifikasi & negosiasi
        with pdfplumber.open(pdf_path) as _plb:
            for _i, _pp in enumerate(_plb.pages):
                _u = (_pp.extract_text() or "").upper()
                _is_ba = "BERITA ACARA" in _u
                _ip = _u.find(_PEMB)
                _ik = _u.find(_KLAR)
                if _ip != -1 and _ip < 200 and not _is_ba:
                    _pemb_pages.append(_i)
                if _ik != -1 and _ik < 200 and not _is_ba:
                    _klarif_pages.append(_i)

        # Titik sisip (0-based index halaman SETELAH mana BA disisipkan).
        # Evaluasi: setelah occurrence #1 pembuktian → butuh >=2 occurrence.
        _insert_after = {}  # {page_index: [list_pdf_path]}
        if _ev_path and len(_pemb_pages) >= 2:
            _insert_after.setdefault(_pemb_pages[0], []).append(_ev_path)
        if _hs_path and len(_klarif_pages) >= 2:
            _insert_after.setdefault(_klarif_pages[0], []).append(_hs_path)

        if not _insert_after:
            return  # tidak ada anchor valid, biarkan PDF apa adanya

        _writer = PdfWriter()
        for _i in range(_n):
            _writer.add_page(_rdr.pages[_i])
            for _bp in _insert_after.get(_i, []):
                try:
                    _brdr = PdfReader(_bp)
                    for _bpg in _brdr.pages:
                        _writer.add_page(_bpg)
                except Exception:
                    pass

        # Tulis ke file sementara lalu ganti (hindari korup jika gagal di tengah)
        _tmp = pdf_path + "_withba.pdf"
        with open(_tmp, "wb") as _f:
            _writer.write(_f)
        os.replace(_tmp, pdf_path)
    except Exception:
        pass  # best-effort, jangan gagalkan cetak utama


def _export_sheet_pdf(excel_path, sheet_name, out_pdf, landscape=True, fit_wide=None, fit_tall=None):
    """Export 1 sheet Excel -> PDF. Return True jika sukses, False jika sheet tak ada/gagal."""
    import win32com.client
    xlApp = None
    wb_xl = None
    try:
        xlApp = win32com.client.DispatchEx("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        wb_xl = xlApp.Workbooks.Open(excel_path, ReadOnly=True)
        try:
            ws = wb_xl.Sheets(sheet_name)
        except Exception:
            return False
        if landscape:
            ws.PageSetup.Orientation = 2  # xlLandscape
        if fit_wide is not None or fit_tall is not None:
            ws.PageSetup.Zoom = False
            if fit_wide is not None:
                ws.PageSetup.FitToPagesWide = fit_wide
            if fit_tall is not None:
                ws.PageSetup.FitToPagesTall = fit_tall
        ws.ExportAsFixedFormat(0, out_pdf)  # 0 = xlTypePDF
        return True
    except Exception:
        return False
    finally:
        if wb_xl:
            try: wb_xl.Close(False)
            except Exception: pass
        if xlApp:
            try: xlApp.Quit()
            except Exception: pass


def _stitch_excel_at_anchor(word_pdf, anchor_excel_pairs, out_pdf):
    """
    Sisip PDF Excel ke word_pdf SETELAH tiap halaman yang mengandung anchor teks.

    anchor_excel_pairs: list of (anchor_text_upper, excel_pdf_path).
      Tiap occurrence anchor di word_pdf -> sisip excel_pdf_path setelah halaman itu.
      Anchor dicocokkan berurutan: occurrence ke-N halaman anchor -> excel ke-N (jika anchor
      sama, daftarkan pasangan itu sekali per occurrence — lihat caller).

    Implementasi: scan tiap halaman word, untuk tiap anchor yang match di halaman,
    jadwalkan sisip excel-nya setelah halaman tsb. Robust thd geseran halaman.
    """
    import pdfplumber
    from pypdf import PdfReader, PdfWriter

    rdr_word = PdfReader(word_pdf)
    n_word = len(rdr_word.pages)

    # cache reader excel per path
    _excel_readers = {}
    def _rdr_excel(p):
        if p not in _excel_readers:
            _excel_readers[p] = PdfReader(p)
        return _excel_readers[p]

    # Hitung occurrence per anchor; tiap (anchor,excel) dipakai sekali berurutan.
    # Bangun antrian per anchor_text -> list excel_pdf (FIFO).
    from collections import defaultdict, deque
    queues = defaultdict(deque)
    for atext, epath in anchor_excel_pairs:
        queues[atext.upper()].append(epath)

    # Tentukan halaman -> list excel yang disisip setelahnya.
    insert_after = defaultdict(list)  # page_idx -> [excel_path,...]
    with pdfplumber.open(word_pdf) as plb:
        for pi, pp in enumerate(plb.pages):
            up = (pp.extract_text() or "").upper()
            for atext, q in queues.items():
                if q and atext in up:
                    insert_after[pi].append(q.popleft())

    writer = PdfWriter()
    for pi in range(n_word):
        writer.add_page(rdr_word.pages[pi])
        for epath in insert_after.get(pi, []):
            for epg in _rdr_excel(epath).pages:
                writer.add_page(epg)
    return _safe_write_pdf(writer, out_pdf)


def _build_bapljkk_final_pdf(wd_doc, folder, kode):
    """Export BA Word + 2 copy sheet 7.2 lalu sisipkan Summary SPSE."""
    from pypdf import PdfReader, PdfWriter
    import pdfplumber
    import win32com.client

    pdf_path = _fit_path(folder, f"BA_PLJKK_{kode}.pdf")
    tmp_word = pdf_path + "_tmpword.pdf"
    tmp_72 = pdf_path + "_tmp72.pdf"
    xlsm_paths = glob.glob(os.path.join(folder, "*.xlsm"))
    xlsm_path = os.path.normpath(xlsm_paths[0]) if xlsm_paths else None
    has_72 = False

    try:
        start = wd_doc.Sections(3).Range.Start if wd_doc.Sections.Count >= 3 else wd_doc.Content.Start
        wd_doc.Range(start, wd_doc.Content.End).ExportAsFixedFormat(
            OutputFileName=tmp_word, ExportFormat=17
        )

        if xlsm_path:
            xl_app = None
            wb = None
            try:
                xl_app = win32com.client.DispatchEx("Excel.Application")
                xl_app.Visible = False
                wb = xl_app.Workbooks.Open(xlsm_path, ReadOnly=True)
                wb.Sheets("7.2 Dengan Nego").ExportAsFixedFormat(
                    Type=0, Filename=tmp_72, Quality=0,
                    IncludeDocProperties=True, IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )
                has_72 = True
            except Exception:
                pass
            finally:
                if wb:
                    try: wb.Close(False)
                    except Exception: pass
                if xl_app:
                    try: xl_app.Quit()
                    except Exception: pass

        if has_72:
            rdr_word = PdfReader(tmp_word)
            split_page = len(rdr_word.pages)
            with pdfplumber.open(tmp_word) as pdf:
                for page_index, page in enumerate(pdf.pages):
                    if "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" in (page.extract_text() or "").upper():
                        split_page = page_index
                        break
            rdr_72 = PdfReader(tmp_72)
            writer = PdfWriter()
            for page_index in range(split_page):
                writer.add_page(rdr_word.pages[page_index])
            for _ in range(2):
                for page in rdr_72.pages:
                    writer.add_page(page)
            for page_index in range(split_page, len(rdr_word.pages)):
                writer.add_page(rdr_word.pages[page_index])
            with open(pdf_path, "wb") as output:
                writer.write(output)
        else:
            shutil.move(tmp_word, pdf_path)
    finally:
        for tmp_path in (tmp_word, tmp_72):
            try: os.remove(tmp_path)
            except Exception: pass

    try:
        from gabung_ba_pljkk import gabung as gabung_ba_pljkk
        result = gabung_ba_pljkk(folder)
        if result.get("ok"):
            return result["output"]
    except Exception:
        pass
    return pdf_path


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
        _merged_b = _fit_path(_folder, f"{_base_b[:60].rstrip()} (Merged){_ext_b}")
        shutil.copy2(_word_path_win, _merged_b)
        _strip_mailmerge_datasource(_merged_b)
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
                _trim_blank_participant_rows(wdDoc)
                _protect_signature_layout(wdDoc)
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
                _pdf_path = _fit_path(_folder, f"BA_PLJKK_{_kode_pljkk}.pdf")
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
                # Bentuk BA final sama seperti tombol "Gabung BA PLJKK".
                # File Summary SPSE tersimpan di subfolder 7, bukan di root,
                # sehingga helper lama tidak pernah menemukannya saat Cetak BA.
                try:
                    from gabung_ba_pljkk import gabung as _gabung_ba_pljkk
                    _gabung_result = _gabung_ba_pljkk(_folder)
                    if _gabung_result.get("ok"):
                        _pdf_path = _gabung_result["output"]
                except Exception:
                    pass
                show_success(_pdf_path)
            elif mode == "printer_bapljkk":
                # Printer harus memakai PDF final, bukan Word mentah.
                # Word mentah melewati Summary SPSE dan dua copy sheet 7.2.
                _printer_name = pdf_name
                _final_pdf = _build_bapljkk_final_pdf(wdDoc, _folder, "PL")
                import win32api
                _result = win32api.ShellExecute(
                    0, "printto", _final_pdf, f'"{_printer_name}"', _folder, 0
                )
                if _result <= 32:
                    raise RuntimeError(f"ShellExecute printto gagal ({_result})")
                time.sleep(3)
                show_print_success(_printer_name)
            wdDoc.Close(False)
        except Exception as e:
            show_error(f"Error cetak BA PLJKK:\n{e}")
        finally:
            wdApp.Quit()
            pythoncom.CoUninitialize()
            try:
                if os.path.exists(_merged_b):
                    import send2trash
                    send2trash.send2trash(_merged_b)
            except Exception:
                pass
        return

    folder = os.path.dirname(word_path)
    base, ext = os.path.splitext(os.path.basename(word_path))
    if ext.lower() not in (".docx", ".docm"):
        ext = ".docx"
    copy_path = _fit_path(folder, f"{base[:60].rstrip()} (Merged){ext}")

    # Copy template ke (Merged) - template asli tidak diubah
    shutil.copy2(word_path, copy_path)
    _strip_mailmerge_datasource(copy_path)

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

        _replace_merge_fields(wdDoc, data)
        _trim_blank_participant_rows(wdDoc)

        # Cleanup blank pages untuk file BA utama (satu_data) yang multi-section.
        # File "2. Isi Reviu" & "3. Dokpil" dikecualikan (struktur beda, bisa berantakan).
        # Template BA dipecah per-dokumen: 4=Undangan, 5=BA Utama, 6=Ringkasan, 7=Timpang.
        _bn_cleanup = os.path.basename(word_path)
        if any(_bn_cleanup.startswith(_p) for _p in (
            "1. Full Dokumen", "4. Undangan", "5. Berita Acara", "6. Ringkasan", "7. BA Dengan"
        )):
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

            # Ambil nama paket dari data dict
            _npk = ''
            for _nk in ['Nama_Paket','NamaTender','Nama_Tender','nama_paket','nama_tender']:
                _v = data.get(_nk) or data.get(_nk.lower())
                if _v and str(_v).strip() not in ('','None','null'):
                    _npk = str(_v).strip(); break
            nama_paket_pdf = _safe_filename(_npk) if _npk else safe_name

            if mode == "pdf_full":
                # Export full dokumen (template BA dipecah per-file, no page-range).
                # Label dokumen dari nama file Word (prefix angka)
                _bn_full = os.path.basename(word_path)
                _label = "Dokumen"
                if _bn_full.startswith("1. BA Reviu") or _bn_full.startswith("1. Reviu"):
                    _label = "BA_REVIU_DPP"
                elif _bn_full.startswith("4. Undangan"):   _label = "Undangan"
                elif _bn_full.startswith("6. Ringkasan"):  _label = "REvaluasi"
                pdf_path = _fit_path(folder, f"{_label}_{nama_paket_pdf}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
                show_success(pdf_path)
            elif mode == "pdf_bareviu":
                pdf_path = _fit_path(folder, f"BA_REVIU_DPP_{nama_paket_pdf}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=3,
                    To=6,
                )
                show_success(pdf_path)
            elif mode == "pdf_bareviu_pl":
                pdf_path = _fit_path(folder, f"BA_REVIU_PL_{nama_paket_pdf}.pdf")
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
                pdf_path = _fit_path(folder, f"BA_PLJKK_{nama_paket_pdf}.pdf")
                # File baru (BA-only) hanya 1 section; file lama (gabung Reviu) BA mulai Section 3
                if wdDoc.Sections.Count >= 3:
                    _start = wdDoc.Sections(3).Range.Start
                else:
                    _start = wdDoc.Content.Start
                _rng = wdDoc.Range(_start, wdDoc.Content.End)
                _rng.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)
                show_success(pdf_path)
            elif mode == "pdf_revaluasi":
                pdf_path = _fit_path(folder, f"REvaluasi_{nama_paket_pdf}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=3,  # wdExportFromTo
                    From=30,
                    To=37,
                )
                show_success(pdf_path)
            elif mode == "pdf_all":
                # Output ke subfolder "6. BA Reviu Lengkap" (buat kalau belum ada)
                _ba_reviu_dir = os.path.join(folder, "6. BA Reviu Lengkap")
                os.makedirs(_ba_reviu_dir, exist_ok=True)
                pdf_path = _fit_path(_ba_reviu_dir, f"Isi_Reviu_DPP_{nama_paket_pdf}.pdf")
                try:
                    wdDoc.ExportAsFixedFormat(
                        OutputFileName=pdf_path,
                        ExportFormat=17,
                        Range=0,  # wdExportAllDocument
                    )
                except Exception:
                    wdDoc.SaveAs2(pdf_path, FileFormat=17)
                show_success(pdf_path)
            elif mode == "pdf_dokpil":
                # Ambil nama paket dari sheet satu_data (list_dokpil tidak punya field nama paket)
                _np_dokpil = nama_paket_pdf
                if _np_dokpil == safe_name:  # fallback belum dapat nama
                    try:
                        _data_sd = read_excel_data(excel_path, "satu_data")
                        if _data_sd:
                            for _nk in ['Nama_Paket', 'NamaTender', 'Nama_Tender', 'nama_paket']:
                                _v = _data_sd.get(_nk) or _data_sd.get(_nk.lower())
                                if _v and str(_v).strip() not in ('', 'None', 'null'):
                                    _np_dokpil = _safe_filename(str(_v).strip())
                                    break
                    except Exception:
                        pass
                pdf_path = _fit_path(folder, f"dokpil_{_np_dokpil}.pdf")
                wdDoc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    Range=0,  # wdExportAllDocument
                )
                show_success(pdf_path)
            elif mode == "pdf_pembuktian":
                # File "5. Berita Acara Utama PK": export full Word -> PDF, sisip sheet
                # "7.2 Dengan Nego" SETELAH tiap halaman anchor nego (2 occurrence).
                # Anchor teks robust thd geseran halaman (ganti page-range/index manual).
                import tempfile
                final_pdf_path = _fit_path(folder, f"BA_Pembuktian_Nego_{nama_paket_pdf}.pdf")
                temp_dir = tempfile.mkdtemp()
                temp_word_pdf = os.path.join(temp_dir, "temp_word.pdf")
                temp_nego_pdf = os.path.join(temp_dir, "temp_nego.pdf")

                wdDoc.ExportAsFixedFormat(
                    OutputFileName=temp_word_pdf, ExportFormat=17, Range=0,
                )
                _has_nego = _export_sheet_pdf(excel_path, "7.2 Dengan Nego", temp_nego_pdf, landscape=True)

                _ANCHOR_NEGO = "DAFTAR HADIR NEGOSIASI KUANTITAS DAN HARGA"
                if _has_nego:
                    # 2 occurrence anchor -> sisip nego 2x (FIFO antrian di helper)
                    pairs = [(_ANCHOR_NEGO, temp_nego_pdf), (_ANCHOR_NEGO, temp_nego_pdf)]
                    final_pdf_path = _stitch_excel_at_anchor(temp_word_pdf, pairs, final_pdf_path)
                else:
                    import shutil as _sh
                    _sh.copy2(temp_word_pdf, final_pdf_path)
                show_success(final_pdf_path)

            elif mode == "pdf_pembuktian_timpang":
                # File "7. BA Dengan Timpang PK": export full Word -> PDF, sisip:
                #   - "7.2 Dengan Nego" setelah tiap anchor nego (2 occurrence)
                #   - "Klarifikasi Timpang Fix (2)" setelah tiap anchor timpang (2 occurrence)
                # Urutan sisip per halaman ditentukan posisi anchor di dokumen (robust).
                import tempfile
                final_pdf_path = _fit_path(folder, f"BA_Pembuktian_Timpang_{nama_paket_pdf}.pdf")
                temp_dir = tempfile.mkdtemp()
                temp_word_pdf = os.path.join(temp_dir, "temp_word.pdf")
                temp_nego_pdf = os.path.join(temp_dir, "temp_nego.pdf")
                temp_timpang_pdf = os.path.join(temp_dir, "temp_timpang.pdf")

                wdDoc.ExportAsFixedFormat(
                    OutputFileName=temp_word_pdf, ExportFormat=17, Range=0,
                )
                _has_nego = _export_sheet_pdf(excel_path, "7.2 Dengan Nego", temp_nego_pdf, landscape=True)
                _has_timpang = _export_sheet_pdf(
                    excel_path, "Klarifikasi Timpang Fix (2)", temp_timpang_pdf,
                    landscape=True, fit_wide=1, fit_tall=1,
                )

                _ANCHOR_NEGO = "DAFTAR HADIR NEGOSIASI KUANTITAS DAN HARGA"
                _ANCHOR_TIMPANG = "DAFTAR HADIR KLARIFIKASI HARGA SATUAN TIMPANG"
                pairs = []
                if _has_nego:
                    pairs += [(_ANCHOR_NEGO, temp_nego_pdf), (_ANCHOR_NEGO, temp_nego_pdf)]
                if _has_timpang:
                    pairs += [(_ANCHOR_TIMPANG, temp_timpang_pdf), (_ANCHOR_TIMPANG, temp_timpang_pdf)]

                if pairs:
                    final_pdf_path = _stitch_excel_at_anchor(temp_word_pdf, pairs, final_pdf_path)
                else:
                    import shutil as _sh
                    _sh.copy2(temp_word_pdf, final_pdf_path)
                show_success(final_pdf_path)

            else:
                pdf_path = _fit_path(folder, f"Undangan_{nama_paket_pdf}.pdf")
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
        # Hapus file (Merged) ke Recycle Bin setelah selesai
        try:
            if os.path.exists(copy_path):
                import send2trash
                send2trash.send2trash(copy_path)
        except Exception:
            pass


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
    """Notifikasi popup setelah PDF selesai dibuat, lalu buka file."""
    try:
        import ctypes
        filename = os.path.basename(pdf_path)
        ctypes.windll.user32.MessageBoxW(
            0, f"PDF berhasil dibuat:\n{filename}", "Export PDF Selesai", 0x40
        )
    except:
        pass
    try:
        if os.path.exists(pdf_path):
            os.startfile(pdf_path)
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
            ("1. BA Reviu DPP PLJKK - Template.docx",      "list_reviu"),
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
