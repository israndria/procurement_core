"""Scrub paket baru dari data donor tanpa merusak layout/macro."""

from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _cell_text(value) -> str:
    return str(value).strip() if value not in (None, "") else ""


def clean_workbook_donor_state(workbook, *, clear_draft_lists: bool = False) -> list[str]:
    """Putus state donor Excel yang dapat menghidupkan kembali data stale.

    Workbook PL standalone tidak boleh membawa externalLink dari workbook
    tender donor. Defined name scoped yang menunjuk ``[1]``/``#REF!`` juga
    harus dihapus; Excel dapat memuat ulang cache external book saat dropdown
    atau formula disentuh walaupun sheet yang terlihat sudah dikosongkan.
    """
    logs: list[str] = []

    try:
        links = workbook.LinkSources(1)  # xlLinkTypeExcelLinks
        if links:
            if isinstance(links, str):
                links = [links]
            for link in list(links):
                try:
                    workbook.BreakLink(Name=str(link), Type=1)
                    logs.append(f"External link diputus: {link}")
                except Exception as exc:
                    logs.append(f"⚠ External link gagal diputus ({link}): {exc}")
    except Exception:
        # Workbook tanpa link mengembalikan None/COM error pada beberapa Excel.
        pass

    # Kumpulkan object sebelum Delete agar penghapusan tidak menggeser index
    # COM collection saat sedang diiterasi.
    stale_names = []
    try:
        for index in range(1, int(workbook.Names.Count) + 1):
            name = workbook.Names.Item(index)
            refers_to = str(name.RefersTo or "")
            if "[" in refers_to or "#REF!" in refers_to:
                stale_names.append(name)
    except Exception as exc:
        logs.append(f"⚠ Audit defined name gagal: {exc}")

    for name in stale_names:
        try:
            name_text = str(name.Name)
            name.Delete()
            logs.append(f"Defined name donor dihapus: {name_text}")
        except Exception as exc:
            logs.append(f"⚠ Defined name donor gagal dihapus: {exc}")

    if clear_draft_lists:
        for sheet_name in ("_DraftPaketList", "_DraftPaketPLList"):
            try:
                workbook.Worksheets(sheet_name).Range("A1:A500").ClearContents()
                logs.append(f"List donor dikosongkan: {sheet_name}")
            except Exception:
                # Varian workbook lama dapat tidak memiliki salah satu sheet.
                pass

    return logs


def _excel_markers(excel_path: Path) -> list[str]:
    """Ambil marker donor dari workbook tanpa menyimpan ulang workbook."""
    try:
        from openpyxl import load_workbook

        wb = load_workbook(str(excel_path), read_only=True, data_only=True, keep_vba=True)
        markers: list[str] = []
        for sheet_name in ("@ Master Data", "0. Input BA", "3. KK Evaluasi Kualifikasi"):
            if sheet_name not in wb.sheetnames:
                continue
            for row in wb[sheet_name].iter_rows():
                for cell in row:
                    value = _cell_text(cell.value)
                    if value and any(x in value.lower() for x in (
                        "kode unik", "kode tender", "nama tender", "kode pokja",
                        "pembangunan", "normalisasi", "pengerukkan", "cv ",
                        "pt ", "firma ", "rup", "pokja",
                    )):
                        markers.append(value)
        wb.close()
        return sorted(set(markers), key=len, reverse=True)
    except Exception:
        return []


def _clear_docm_content_controls(data: bytes) -> bytes:
    """Kosongkan isi CC, pertahankan tag, format, dan struktur CC."""
    with tempfile.TemporaryDirectory() as td:
        src, out = Path(td) / "in.docm", Path(td) / "out.docm"
        src.write_bytes(data)
        with zipfile.ZipFile(src, "r") as zin:
            if "word/document.xml" not in zin.namelist():
                return data
            root = ET.fromstring(zin.read("word/document.xml"))
            count = 0
            for sdt in root.findall(f".//{W}sdt"):
                content = sdt.find(f"{W}sdtContent")
                if content is None:
                    continue
                for node in content.iter():
                    if node.tag == f"{W}t":
                        node.text = ""
                count += 1
            if not count:
                return data
            new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    payload = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(item, payload)
            return out.read_bytes()


def _replace_markers_in_office(data: bytes, markers: list[str]) -> bytes:
    """Hapus marker donor dari XML Word; format dokumen tetap."""
    if not markers:
        return data
    with tempfile.TemporaryDirectory() as td:
        src, out = Path(td) / "in.docx", Path(td) / "out.docx"
        src.write_bytes(data)
        changed = False
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    text = payload.decode("utf-8", "ignore")
                    new_text = text
                    for marker in markers:
                        new_text = new_text.replace(marker, "")
                    if new_text != text:
                        payload = new_text.encode("utf-8")
                        changed = True
                zout.writestr(item, payload)
        return out.read_bytes() if changed else data


def _clear_constants_com(excel_path: Path, targets: dict[str, list[str]]) -> list[str]:
    """Clear konstanta via satu sesi Excel COM. Formula/macro/shape tetap utuh."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        try:
            xl.AutomationSecurity = 1
        except Exception:
            pass
        # Scrub harus menjadi operasi data-only. Jangan biarkan Workbook_Open/
        # Workbook_BeforeSave donor menulis ulang snapshot atau cache lama.
        xl.EnableEvents = False
        try:
            xl.Calculation = -4135  # xlCalculationManual
            xl.CalculateBeforeSave = False
        except Exception:
            pass
        wb = xl.Workbooks.Open(str(excel_path), UpdateLinks=0, ReadOnly=False)
        logs: list[str] = []
        # xlCellTypeConstants = 2; satu ClearContents jauh lebih cepat daripada
        # iterasi COM per sel, terutama sheet KK yang lebarnya 195 kolom.
        for sheet_name, ranges in targets.items():
            ws = wb.Worksheets(sheet_name)
            cleared = 0
            for address in ranges:
                try:
                    constants = ws.Range(address).SpecialCells(2)
                    constants.ClearContents()
                    cleared += 1
                except Exception:
                    # 1004 = tidak ada konstanta pada area; aman diabaikan.
                    pass
            logs.append(f"Excel scrub {sheet_name}: {cleared} area dibersihkan")
        logs.extend(clean_workbook_donor_state(wb, clear_draft_lists=True))
        wb.Save()
        wb.Close(SaveChanges=False)
        return logs
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()


def scrub_excel_copy(excel_path: str | Path) -> list[str]:
    """Bersihkan area data donor dari copy workbook tender baru."""
    path = Path(excel_path)
    logs: list[str] = []
    targets = {
        "@ Master Data": ["C3:C70", "H2:I23"],
        "0. Input BA": [
            "C3:G5", "C7:E14", "G7:G14", "I7:I14", "C17:E22",
            "G17:N22", "C25:C29", "C32:G33", "I32:I33", "C38:L53",
        ],
        "3. KK Evaluasi Kualifikasi": ["C3:GR92"],
        # 10 blok x 9 kolom, blok terakhir berakhir di CK.
        "6. Harga Penawaran": ["A1:CK171"],
        "6. Harga Penawaran (2)": ["A1:CK171"],
        "6. Harga Penawaran (3)": ["A1:CK171"],
        "5. HPS": ["A1:Z300"],
    }
    try:
        logs.extend(_clear_constants_com(path, targets))
    except Exception as exc:
        logs.append(f"⚠ Excel scrub gagal: {exc}")
    return logs


def scrub_excel_pl_copy(excel_path: str | Path, *, is_pk: bool) -> list[str]:
    """Bersihkan data contoh workbook PL tanpa mengubah layout/macro.

    Donor PL JKK dan PL PK punya layout berbeda. Area input paket dibersihkan
    berdasarkan family agar contoh paket lama, provider, SBU, personil, alat,
    dan nilai evaluasi tidak ikut terbawa ke paket baru. Blok INFO PP yang
    memang hardcode tetap dipertahankan.
    """
    path = Path(excel_path)
    if is_pk:
        # Hanya kolom C yang berisi data paket. Jangan gunakan satu range
        # besar C29:C80: beberapa baris C adalah bagian merged header A:D
        # (A29:D29, A32:D32, A65:D65, A76:D76). SpecialCells pada range
        # tersebut dapat mengembalikan merged area penuh dan menghapus label
        # kolom B. Kolom F adalah panel/keterangan BA Reviu dan wajib utuh.
        md_ranges = [
            "C3:C10", "C13:C28", "C30:C31", "C33:C64",
            "C66:C75", "C77:C80", "C87:C89", "H8:H10",
        ]
    else:
        # JKK memiliki merged header berbeda; tetap batasi scrub ke sel data
        # eksplisit dan pertahankan seluruh kolom B/F sebagai label/panel.
        md_ranges = [
            "C3:C10", "C13:C28", "C30:C54", "C61:C63", "H8:H10",
        ]
    targets = {
        "@ Master Data": md_ranges,
        "5. HPS": ["A2:I300"],
        # Rantai data PLPK lain juga dapat membawa donor walau @ Master Data
        # sudah bersih. Hanya konstanta dibersihkan; formula/layout tetap utuh.
        "6. Penawaran": ["A2:I300"],
        "7.2 Dengan Nego": ["A8:AR26", "U1:U4", "U35", "AH27"],
        "@ Evaluasi": ["C3:D47"],
        "Hasil Evaluasi": ["A2:F300"],
        "Harga Timpang": ["A8:K26"],
        "database_reviu": ["E2:J38"],
        "_DraftPaketList": ["A1:A500"],
        "_DraftPaketPLList": ["A1:A500"],
    }
    logs: list[str] = []
    try:
        logs.extend(_clear_constants_com(path, targets))
    except Exception as exc:
        logs.append(f"⚠ Excel scrub PL gagal: {exc}")
    return logs


def scrub_package_copy(target_dir: str | Path, excel_path: str | Path,
                       word_paths: list[str | Path]) -> list[str]:
    """Scrub copy baru + hasilkan log audit."""
    target = Path(target_dir)
    excel = Path(excel_path)
    markers = _excel_markers(excel)
    logs = scrub_excel_copy(excel)
    for word in word_paths:
        path = Path(word)
        if not path.exists() or path.suffix.lower() not in (".docx", ".docm"):
            continue
        try:
            original = path.read_bytes()
            data = _replace_markers_in_office(original, markers)
            if path.suffix.lower() == ".docm":
                data = _clear_docm_content_controls(data)
            if data != original:
                path.write_bytes(data)
                logs.append(f"Word scrub: {path.name}")
        except Exception as exc:
            logs.append(f"⚠ Word scrub {path.name} gagal: {exc}")

    for name in ("jawaban_reviu.json", "_parse_reviu.json"):
        stale = target / name
        if stale.exists():
            stale.unlink()
            logs.append(f"Hapus state donor: {name}")
    return logs
