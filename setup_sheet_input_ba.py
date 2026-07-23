"""
Buat sheet "0. Input BA" + update referensi hardcode di sheet BA ke sheet baru.
Dijalankan sekali saja per workbook baru (setup_paket_baru atau manual).

Yang dilakukan:
  1. Buat sheet "0. Input BA" jika belum ada, letakkan sebelum "1. Input Data"
  2. Isi header + label layout
  3. Update cell hardcode di sheet 3,5,6,7,9 → formula referensi ke "0. Input BA"
"""

import win32com.client
import pythoncom
import os
import time


def setup_sheet_input_ba(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Setup sheet '0. Input BA': {filepath}")

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(filepath)

        # ── 1. Buat sheet "0. Input BA" jika belum ada ──────────────
        sheet_names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
        if "0. Input BA" not in sheet_names:
            # Sisipkan sebelum "1. Input Data"
            idx_input = next((i for i, n in enumerate(sheet_names, 1) if n == "1. Input Data"), 1)
            wb.Sheets.Add(Before=wb.Sheets(idx_input)).Name = "0. Input BA"
            print("  [OK] Sheet '0. Input BA' dibuat")
        else:
            print("  [SKIP] Sheet '0. Input BA' sudah ada")

        ws = wb.Sheets("0. Input BA")
        ws.Unprotect()

        # Tab warna ungu seperti sheet input
        ws.Tab.Color = 10498160  # ungu

        # ── 2. Isi layout sheet "0. Input BA" ───────────────────────
        # Format: kolom B = label, C = peserta 1, D = peserta 2, E = peserta 3

        def cell(r, c):
            return ws.Cells(r, c)

        def label(r, text, bold=False):
            cell(r, 2).Value = text
            if bold:
                cell(r, 2).Font.Bold = True

        def header(r, text):
            cell(r, 2).Value = text
            cell(r, 2).Font.Bold = True
            cell(r, 2).Interior.Color = 0xD9D9D9  # abu-abu

        # Baris 1: judul
        cell(1, 2).Value = "0. INPUT BA — Sumber Data Berita Acara"
        cell(1, 2).Font.Bold = True
        cell(1, 2).Font.Size = 12

        # Blok Tanggal BA (baris 3-4)
        header(2, "TANGGAL BA")
        label(3, "Tgl Pembukaan Penawaran")
        label(4, "Tgl Pembuktian / Klarif / Penetapan")
        # Kolom D: hari otomatis dari tanggal
        ws.Cells(3, 4).Formula = '=IF(C3="","",TEXT(C3,"dddd"))'
        ws.Cells(4, 4).Formula = '=IF(C4="","",TEXT(C4,"dddd"))'
        # Format kolom C sebagai tanggal
        ws.Cells(3, 3).NumberFormat = "dd mmmm yyyy"
        ws.Cells(4, 3).NumberFormat = "dd mmmm yyyy"

        # Blok Identitas Peserta (baris 7-10)
        header(6, "IDENTITAS PESERTA")
        label(7, "Nama Perusahaan")
        label(8, "NPWP")
        label(9, "Alamat")
        label(10, "Direktur / Pemilik")

        # Header kolom peserta
        for col, txt in [(3, "Peserta 1"), (4, "Peserta 2"), (5, "Peserta 3")]:
            ws.Cells(6, col).Value = txt
            ws.Cells(6, col).Font.Bold = True

        # Blok Personel & Alat (baris 13-19)
        header(12, "PERSONEL & PERALATAN (dari Dok. Teknis PDF)")
        label(13, "Personel Manajerial 1")
        label(14, "Personel Manajerial 2")
        header(16, "Peralatan Utama")
        label(17, "Alat 1")
        label(18, "Alat 2")
        label(19, "Alat 3")

        # Blok Dokumen Penawaran (baris 22-24)
        header(21, "DOKUMEN PENAWARAN")
        label(22, "Jml Peserta Daftar")
        label(23, "Jml Dok Terkirim (dapat dibuka)")
        label(24, "Jml Dok Tidak Terkirim")

        # Blok Hasil dari KK Evaluasi (baris 27-28)
        header(26, "HASIL EVALUASI (dari sheet KK Evaluasi)")
        label(27, "SKP")
        label(28, "Hasil Pembuktian Kualifikasi")
        # SKP dan Hasil Pembuktian diisi oleh VBA MuatInputBA() — bukan formula
        # (formula ke KK Evaluasi akan circular via sheet 6 W29)
        ws.Cells(27, 3).Value = ""  # diisi VBA
        ws.Cells(28, 3).Value = ""  # diisi VBA

        # Lebar kolom
        ws.Columns(2).ColumnWidth = 38
        ws.Columns(3).ColumnWidth = 30
        ws.Columns(4).ColumnWidth = 20
        ws.Columns(5).ColumnWidth = 20

        print("  [OK] Layout sheet '0. Input BA' selesai")

        # Slot alat kosong harus benar-benar blank. Referensi langsung ke
        # database_reviu mengubah sel kosong menjadi 0 sehingga BA menulis
        # "3. 0 0 (0)" dan seterusnya.
        _update_equipment_formulas(ws)

        # ── 3. Update referensi di sheet BA ─────────────────────────
        _update_sheet3(wb)
        _update_sheet5(wb)
        _update_sheet6(wb)
        _update_sheet7(wb)
        _update_sheet9(wb)
        _update_nego_rounding(wb)

        # Hide sheet BA (opsional — biarkan visible dulu untuk verifikasi)
        # for sn in ["3. BA Pembukaan Penawaran","5. BA Pembuktian Kualifikasi",
        #            "6. BA KLARIF SKP ALAT","7. BA Klarifikasi HS","9. BA Penetapan Pemenang"]:
        #     try: wb.Sheets(sn).Visible = False
        #     except: pass

        # ── 4. Save ─────────────────────────────────────────────────
        wb.Save()
        time.sleep(1)
        print("  [OK] Disimpan")

    except Exception as e:
        print(f"[ERROR] {e}")
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


# ── Helper: update tiap sheet BA ────────────────────────────

def _update_sheet3(wb):
    """Sheet '3. BA Pembukaan Penawaran': tanggal S13/S14/S15 + jumlah dok P26/P28/P29/P31/P32"""
    try:
        ws = wb.Sheets("3. BA Pembukaan Penawaran")
        ws.Unprotect()
        # Tanggal: S13=hari, S14=bulan, S15=tahun → ambil dari "0. Input BA" C3
        ws.Range("S13").Formula = "=DAY('0. Input BA'!C3)"
        ws.Range("S14").Formula = "=MONTH('0. Input BA'!C3)"
        ws.Range("S15").Formula = "=YEAR('0. Input BA'!C3)"
        # Jumlah dokumen
        ws.Range("P26").Formula = "='0. Input BA'!C23"   # jml kirim = dapat dibuka
        ws.Range("P28").Formula = "='0. Input BA'!C23"   # lengkap = kirim
        ws.Range("P29").Formula = "='0. Input BA'!C24"   # tidak lengkap
        ws.Range("P31").Formula = "='0. Input BA'!C23"   # dapat dibuka
        ws.Range("P32").Formula = "='0. Input BA'!C24"   # tidak dapat dibuka
        print("  [OK] Sheet 3 updated (tanggal + jumlah dok)")
    except Exception as e:
        print(f"  [WARN] Sheet 3: {e}")


def _update_equipment_formulas(ws):
    """Isi G17:G22 hanya jika nama/kapasitas/jumlah alat benar-benar ada."""
    for i in range(6):
        row = 17 + i
        src = 9 + i
        ws.Cells(row, 7).Formula = (
            f'=IF(AND(OR(database_reviu!E{src}="",database_reviu!E{src}=0),'
            f'OR(database_reviu!E{src + 6}="",database_reviu!E{src + 6}=0),'
            f'OR(database_reviu!E{src + 12}="",database_reviu!E{src + 12}=0)),"",'
            f'database_reviu!E{src}&", "&database_reviu!E{src + 6}&'
            f'" ("&database_reviu!E{src + 12}&")")'
        )
        # Q adalah item bernomor yang dirangkai ke H17. Dasarkan keberadaan
        # item pada G saja; kolom J berisi koma separator template sehingga
        # pengecekan G&J lama menghasilkan "3. " untuk slot kosong.
        ws.Cells(row, 17).Formula = (
            f'=IF(TRIM(G{row})="","",'
            f'SUBSTITUTE(UPPER(P{row}&". "&O{row}&G{row}&J{row}),",",""))'
        )


def _update_nego_rounding(wb):
    """Pertahankan nilai nego tepat ribuan; bulatkan hanya bila perlu."""
    try:
        ws = wb.Sheets("7.2 Dengan Nego")
        ws.Unprotect()
        ws.Range("P23").Formula = (
            "=IF(MOD(ROUND(P22,0),1000)=0,ROUND(P22,0),"
            "ROUNDDOWN(ROUND(P22,0),-3))"
        )
        try:
            ws_hs = wb.Sheets("7. BA Klarifikasi HS")
            ws_hs.Unprotect()
            ws_hs.Range("G59").Formula = (
                "=IF(MOD(ROUND(G58,0),1000)=0,ROUND(G58,0),"
                "ROUNDDOWN(ROUND(G58,0),-3))"
            )
        except Exception:
            pass
        print("  [OK] Rumus pembulatan nego diperbarui")
    except Exception as e:
        print(f"  [WARN] Rumus pembulatan nego: {e}")


def _update_sheet5(wb):
    """Sheet '5. BA Pembuktian Kualifikasi': tanggal S6/S7/S8, hasil H18/K18"""
    try:
        ws = wb.Sheets("5. BA Pembuktian Kualifikasi")
        ws.Unprotect()
        ws.Range("S6").Formula = "=DAY('0. Input BA'!C4)"
        ws.Range("S7").Formula = "=MONTH('0. Input BA'!C4)"
        ws.Range("S8").Formula = "=YEAR('0. Input BA'!C4)"
        # Hasil pembuktian peserta 1 (H18=Memenuhi/Tidak, K18=Lulus/TMS)
        ws.Range("H18").Formula = "='0. Input BA'!C28"
        ws.Range("K18").Formula = '=IF(\'0. Input BA\'!C28="Memenuhi","Lulus","TMS")'
        print("  [OK] Sheet 5 updated (tanggal + hasil pembuktian)")
    except Exception as e:
        print(f"  [WARN] Sheet 5: {e}")


def _update_sheet6(wb):
    """Sheet '6. BA KLARIF SKP ALAT': personel V37/V38, SKP W29, alat V48"""
    try:
        ws = wb.Sheets("6. BA KLARIF SKP ALAT")
        ws.Unprotect()
        # Personel manajerial
        ws.Range("V37").Formula = "='0. Input BA'!C13"
        ws.Range("V38").Formula = "='0. Input BA'!C14"
        # SKP: W29 = angka — JANGAN dirujuk dari "0. Input BA" karena akan circular
        # (KK Eval C33 → BA KLARIF W29 → KK Eval C33). Biarkan W29 diisi VBA MuatInputBA langsung.
        # ws.Range("W29").Formula = ...  ← sengaja tidak diubah
        # Alat: V48 adalah array formula — biarkan, cukup update sumber alat di "0. Input BA"
        # Formula alat utama di sheet 6 F48 pakai V48 — update V48 sebagai teks gabungan
        ws.Range("V48").Formula = "=\"yaitu \"&'0. Input BA'!C17&IF('0. Input BA'!C18<>\"\",\" dan \"&'0. Input BA'!C18,\"\")"
        print("  [OK] Sheet 6 updated (personel + SKP + alat)")
    except Exception as e:
        print(f"  [WARN] Sheet 6: {e}")


def _update_sheet7(wb):
    """Sheet '7. BA Klarifikasi HS': NPWP M55, alamat M56, direktur B71/B72"""
    try:
        ws = wb.Sheets("7. BA Klarifikasi HS")
        ws.Unprotect()
        # NPWP raw (tanpa format) — formula di G55 sudah format dari M55
        ws.Range("M55").Formula = "='0. Input BA'!C8"
        ws.Range("M56").Formula = "='0. Input BA'!C9"
        # Direktur: B71 = nama (UPPER), B72 = gelar
        # Pisah nama+gelar: nama = bagian sebelum koma, gelar = koma+gelar
        ws.Range("B71").Formula = "=IFERROR(LEFT(TRIM('0. Input BA'!C10),FIND(\",\",'0. Input BA'!C10)-1),TRIM('0. Input BA'!C10))"
        ws.Range("B72").Formula = "=IFERROR(MID(TRIM('0. Input BA'!C10),FIND(\",\",'0. Input BA'!C10),100),\"\")"
        print("  [OK] Sheet 7 updated (NPWP + alamat + direktur)")
    except Exception as e:
        print(f"  [WARN] Sheet 7: {e}")


def _update_sheet9(wb):
    """Sheet '9. BA Penetapan Pemenang': tanggal R15/R16/R17"""
    try:
        ws = wb.Sheets("9. BA Penetapan Pemenang")
        ws.Unprotect()
        ws.Range("R15").Formula = "=DAY('0. Input BA'!C4)"
        ws.Range("R16").Formula = "=MONTH('0. Input BA'!C4)"
        ws.Range("R17").Formula = "=YEAR('0. Input BA'!C4)"
        print("  [OK] Sheet 9 updated (tanggal penetapan)")
    except Exception as e:
        print(f"  [WARN] Sheet 9: {e}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        setup_sheet_input_ba(sys.argv[1])
    else:
        setup_sheet_input_ba(r"D:\Dokumen\@ POKJA 2026\Paket Experiment\@ BA PK 2026 (Improved) v1.4.xlsm")
