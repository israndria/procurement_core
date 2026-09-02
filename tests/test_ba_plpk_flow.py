from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

import pytest
from pypdf import PdfReader, PdfWriter

from gabung_ba_pljkk import (
    deteksi_file,
    ensure_plpk_provider_signature_copy,
    gabung,
)
from document_profiles import inject_header_profile
from word_merge import _find_active_xlsm, _patch_plpk_layout_xml, _resolve_ba_kind


def _blank_pdf(path: Path):
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    with path.open("wb") as output:
        writer.write(output)


def _text_pdf(path: Path, pages: list[str]):
    import fitz

    document = fitz.open()
    for value in pages:
        page = document.new_page(width=500, height=700)
        page.insert_text((40, 60), value, fontsize=12)
    document.save(path)
    document.close()


def _page_texts(path: Path) -> list[str]:
    reader = PdfReader(path)
    return [" ".join((page.extract_text() or "").split()) for page in reader.pages]


def test_deteksi_file_is_scoped_to_ba_type(tmp_path):
    _blank_pdf(tmp_path / "BA_PLPK_GPR_P.Bng.pdf")
    _blank_pdf(tmp_path / "BA_PLJKK_OLD.pdf")

    plpk = deteksi_file(str(tmp_path), "PLPK")
    jkk = deteksi_file(str(tmp_path), "PLJKK")

    assert Path(plpk["ba_utama"]).name == "BA_PLPK_GPR_P.Bng.pdf"
    assert plpk["kode"] == "GPR_P.Bng"
    assert Path(jkk["ba_utama"]).name == "BA_PLJKK_OLD.pdf"
    assert jkk["kode"] == "OLD"


def test_gabung_plpk_writes_full_output_to_package_root(tmp_path):
    _blank_pdf(tmp_path / "BA_PLPK_GPR_P.Bng.pdf")

    result = gabung(str(tmp_path), "PLPK")

    assert result["ok"] is True
    assert Path(result["output"]).name == "BA_PLPK_FULL_Gabungan_GPR_P.Bng.pdf"
    assert Path(result["output"]).parent == tmp_path
    assert Path(result["output"]).read_bytes() == (tmp_path / "BA_PLPK_GPR_P.Bng.pdf").read_bytes()


def test_deteksi_file_ignores_previous_full_output(tmp_path):
    _blank_pdf(tmp_path / "BA_PLJKK_CURRENT.pdf")
    _blank_pdf(tmp_path / "BA_PLJKK_FULL_Gabungan_CURRENT.pdf")

    result = deteksi_file(str(tmp_path), "PLJKK")

    assert Path(result["ba_utama"]).name == "BA_PLJKK_CURRENT.pdf"
    assert result["ba_pembuktian"] is None


def test_plpk_signature_copy_is_dynamic_and_idempotent(tmp_path):
    pdf = tmp_path / "BA_PLPK_DYNAMIC.pdf"
    _text_pdf(
        pdf,
        [
            "BERITA ACARA PEMBUKAAN PENAWARAN",
            "Pejabat Pengadaan pada Dinas Perdagangan\nDIREKTUR/PIMPINAN\nCV CONTOH\n"
            "Demikian Berita Acara Klarifikasi dan Negosiasi ini dibuat",
            "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI",
        ],
    )

    assert ensure_plpk_provider_signature_copy(str(pdf)) is True
    assert _page_texts(pdf) == [
        "BERITA ACARA PEMBUKAAN PENAWARAN",
        "Pejabat Pengadaan pada Dinas Perdagangan DIREKTUR/PIMPINAN CV CONTOH "
        "Demikian Berita Acara Klarifikasi dan Negosiasi ini dibuat",
        "Pejabat Pengadaan pada Dinas Perdagangan DIREKTUR/PIMPINAN CV CONTOH "
        "Demikian Berita Acara Klarifikasi dan Negosiasi ini dibuat",
        "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI",
    ]
    assert ensure_plpk_provider_signature_copy(str(pdf)) is False
    assert len(PdfReader(pdf).pages) == 4


def test_old_pljkk_command_is_upgraded_for_a_plpk_package():
    word_path = r"D:\Paket\29. PLPK - Pemasangan\5. BA PLPK - Pemasangan.docx"
    excel_path = r"D:\Paket\29. PLPK - Pemasangan\0. BAPLPK- Pemasangan.xlsm"

    assert _resolve_ba_kind("pdf_bapljkk", word_path, excel_path) == "PLPK"
    assert _resolve_ba_kind("pdf_bapljkk", r"D:\Paket\BA PLJKK.docx", r"D:\Paket\BAPLJKK.xlsm") == "PLJKK"


def test_headerless_docx_gets_header_part_relationship_and_content_type(tmp_path):
    content_types = (
        b'<?xml version="1.0"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/word/document.xml" ContentType="main"/>'
        b'</Types>'
    )
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b'<w:body><w:p/><w:sectPr><w:pgSz/></w:sectPr></w:body></w:document>'
    )
    rels = (
        b'<?xml version="1.0"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    profile_content_types = (
        b'<?xml version="1.0"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Override PartName="/word/header1.xml" ContentType="header"/>'
        b'</Types>'
    )
    header = b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>'

    template = tmp_path / "template.docx"
    profile = tmp_path / "profile.docx"
    output = tmp_path / "output.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/_rels/document.xml.rels", rels)
    with ZipFile(profile, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", profile_content_types)
        archive.writestr("word/header1.xml", header)

    inject_header_profile(template, profile, output)

    with ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml")
        content_xml = archive.read("[Content_Types].xml")
        assert b"headerReference" in document_xml
        assert b"word/header1.xml" in content_xml
        assert "word/header1.xml" in archive.namelist()


def test_plpk_layout_patch_normalizes_attendance_gap_and_result_heading(tmp_path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body>'
        b'<w:p><w:r><w:t>SEBELUMNYA</w:t></w:r></w:p>'
        b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>'
        b'<w:r><w:t>DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI TEKNIS DAN HARGA</w:t></w:r></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>'
        b'<w:r><w:t>BERITA ACARA HASIL PENGADAAN LANGSUNG</w:t></w:r></w:p>'
        b'<w:p><w:r><w:t>Dinas Pekerjaan Umum dan Penataan Ruang Kabupaten Tapin</w:t></w:r></w:p>'
        b'<w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "plpk-layout.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    _patch_plpk_layout_xml(template, {"SKPDOPD": "Dinas Perdagangan Kabupaten Tapin"})

    with ZipFile(template) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")

    blocks = re.findall(r"<w:p\b[^>]*>.*?</w:p\s*>", xml, re.S)
    attendance_index = next(
        i for i, block in enumerate(blocks) if "DAFTAR HADIR KLARIFIKASI" in block
    )
    assert "w:type=\"page\"" in blocks[attendance_index - 1]
    assert "SEBELUMNYA" in blocks[attendance_index - 2]

    result_pos = xml.index("BERITA ACARA HASIL PENGADAAN LANGSUNG")
    result_block = xml[xml.rfind("<w:p", 0, result_pos) : xml.index("</w:p>", result_pos) + len("</w:p>")]
    assert '<w:pStyle w:val="BodyText"/>' in result_block
    assert '<w:ind w:left="709" w:right="707"/>' in result_block
    assert '<w:jc w:val="center"/>' in result_block
    assert "Dinas Pekerjaan Umum dan Penataan Ruang Kabupaten Tapin" not in xml
    assert "Dinas Perdagangan Kabupaten Tapin" in xml


def test_plpk_layout_patch_adds_break_when_first_attendance_heading_has_none(tmp_path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body>'
        b'<w:p><w:r><w:t>Lanjutan tanda tangan BA</w:t></w:r></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>'
        b'<w:p><w:pPr><w:spacing w:after="200"/></w:pPr>'
        b'<w:r><w:t>DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI TEKNIS DAN HARGA</w:t></w:r></w:p>'
        b'<w:p><w:r><w:t>Salinan daftar hadir</w:t></w:r></w:p>'
        b'<w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "plpk-layout-no-break.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    _patch_plpk_layout_xml(template, {})

    with ZipFile(template) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")

    blocks = re.findall(r"<w:p\b[^>]*>.*?</w:p\s*>", xml, re.S)
    attendance = [
        block for block in blocks if "DAFTAR HADIR KLARIFIKASI" in block
    ][0]
    assert "<w:pageBreakBefore/>" in attendance
    assert xml.count("<w:pageBreakBefore/>") == 1


def test_plpk_vba_word_resolver_skips_generated_merged_copy():
    source = Path(__file__).resolve().parents[1] / "ModDraftPaketPL.bas"
    content = source.read_text(encoding="utf-8")

    assert 'InStr(1, f, "(Merged)", vbTextCompare) = 0' in content
    assert 'InStr(1, f, "(Dengan Header", vbTextCompare) = 0' in content


def test_plpk_vba_passes_current_workbook_to_merge_engine():
    source = Path(__file__).resolve().parents[1] / "ModDraftPaketPL.bas"
    content = source.read_text(encoding="utf-8")
    start = content.index("Private Sub CetakBAByJenis")
    procedure = content[start:content.index("End Sub", start)]

    assert 'Chr(34) & ThisWorkbook.FullName & Chr(34)' in procedure


def test_active_xlsm_resolver_ignores_backup_copies(tmp_path):
    active = tmp_path / "0. BAPLPK - Paket.xlsm"
    active.write_bytes(b"active")
    (tmp_path / "0. BAPLPK - Paket.backup_20260901.xlsm").write_bytes(b"old")
    (tmp_path / "0. BAPLPK - Paket.xmlbatch4-backup.xlsm").write_bytes(b"old")

    assert Path(_find_active_xlsm(tmp_path)) == active


def test_active_xlsm_resolver_does_not_misclassify_package_word_tempat(tmp_path):
    active = tmp_path / "0. BAPLPK - Bangunan Gedung Tempat Kerja.xlsm"
    active.write_bytes(b"active")

    assert Path(_find_active_xlsm(tmp_path)) == active


def test_active_xlsm_resolver_ignores_hashed_backup_copy(tmp_path):
    active = tmp_path / "0. BAPLPK - Paket.xlsm"
    active.write_bytes(b"active")
    (tmp_path / "0. BAPLPK - Paket__f6b01f12.xlsm").write_bytes(b"old")

    assert Path(_find_active_xlsm(tmp_path)) == active


def test_active_xlsm_resolver_uses_explicit_workbook_before_folder_scan(tmp_path):
    active = tmp_path / "0. BAPLPK - Paket.xlsm"
    explicit = tmp_path / "0. BAPLPK - Paket pilihan.xlsm"
    active.write_bytes(b"active")
    explicit.write_bytes(b"explicit")

    assert Path(_find_active_xlsm(tmp_path, preferred=explicit)) == explicit


def test_active_xlsm_resolver_rejects_explicit_backup(tmp_path):
    backup = tmp_path / "0. BAPLPK - Paket.backup_20260901.xlsm"
    backup.write_bytes(b"old")

    with pytest.raises(RuntimeError, match="backup"):
        _find_active_xlsm(tmp_path, preferred=backup)


def test_active_xlsm_resolver_fails_closed_for_multiple_active_workbooks(tmp_path):
    (tmp_path / "0. BAPLPK - Paket A.xlsm").write_bytes(b"a")
    (tmp_path / "0. BAPLPK - Paket B.xlsm").write_bytes(b"b")

    with pytest.raises(RuntimeError, match="Lebih dari satu workbook aktif"):
        _find_active_xlsm(tmp_path)


def test_plpk_layout_patch_removes_only_first_transition_break(tmp_path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body>'
        b'<w:p><w:r><w:t>BA Pembukaan selesai</w:t></w:r></w:p>'
        b'<w:p><w:r><w:t>6.a. BA KLARIFIKASI SKP, ALAT, PERSONEL</w:t></w:r></w:p>'
        b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        b'<w:p><w:r><w:t>BERITA ACARA PEMBUKTIAN KUALIFIKASI</w:t></w:r></w:p>'
        b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        b'<w:p><w:r><w:t>BERITA ACARA PEMBUKTIAN KUALIFIKASI</w:t></w:r></w:p>'
        b'<w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "transition.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    _patch_plpk_layout_xml(template, {})

    with ZipFile(template) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    blocks = re.findall(r"<w:p\b[^>]*>.*?</w:p\s*>", xml, re.S)
    headings = [i for i, block in enumerate(blocks) if "BERITA ACARA PEMBUKTIAN" in block]
    assert len(headings) == 2
    assert "w:type=\"page\"" not in blocks[headings[0] - 1]
    assert "w:type=\"page\"" in blocks[headings[1] - 1]


def test_gabung_ba_vba_requires_confirmation_before_running_python():
    source = Path(__file__).resolve().parents[1] / "ModDraftPaketPL.bas"
    content = source.read_text(encoding="utf-8")

    start = content.index("Private Sub GabungBAByJenis")
    end = content.index("End Sub", start)
    procedure = content[start:end]
    assert "vbYesNo + vbQuestion" in procedure
    assert "<> vbYes Then" in procedure


def test_plpk_layout_patch_only_locks_signature_rows_and_drops_cached_break(tmp_path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body><w:tbl>'
        b'<w:tr><w:tc><w:p><w:r><w:t>Isi BA yang boleh mengalir</w:t></w:r></w:p></w:tc></w:tr>'
        b'<w:tr><w:tc><w:p><w:r><w:t>DIREKTUR/PIMPINAN</w:t><w:lastRenderedPageBreak/></w:r></w:p></w:tc></w:tr>'
        b'<w:tr><w:tc><w:p><w:r><w:t>Nama Penandatangan</w:t></w:r></w:p></w:tc></w:tr>'
        b'<w:tr><w:tc><w:p><w:r><w:t>NIP</w:t></w:r></w:p></w:tc></w:tr>'
        b'</w:tbl><w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "plpk-signature-rows.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    _patch_plpk_layout_xml(template, {})

    with ZipFile(template) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")

    rows = re.findall(r"<w:tr\b[^>]*>.*?</w:tr\s*>", xml, re.S)
    assert "w:cantSplit" not in rows[0]
    assert all("w:cantSplit" in row for row in rows[1:])
    assert "lastRenderedPageBreak" not in xml
    assert all("w:keepNext" in row for row in rows[1:-1])
    assert 'w:trHeight w:val="1200" w:hRule="atLeast"' in rows[2]
    assert '<w:vAlign w:val="bottom"/>' in rows[2]


def test_plpk_layout_patch_preserves_empty_table_cell_paragraphs(tmp_path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body><w:tbl><w:tr>'
        b'<w:tc><w:p><w:pPr><w:spacing w:before="100"/></w:pPr></w:p></w:tc>'
        b'<w:tc><w:p><w:pPr><w:spacing w:before="100"/></w:pPr></w:p></w:tc>'
        b'<w:tc><w:p><w:pPr><w:spacing w:before="100"/></w:pPr></w:p></w:tc>'
        b'</w:tr></w:tbl>'
        b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        b'<w:p><w:r><w:t>DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI TEKNIS DAN HARGA</w:t></w:r></w:p>'
        b'<w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "plpk-table-empty-paragraphs.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    _patch_plpk_layout_xml(template, {})

    with ZipFile(template) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")

    cells = re.findall(r"<w:tc\b[^>]*>.*?</w:tc\s*>", xml, re.S)
    assert len(cells) == 3
    assert all(re.search(r"<w:p\b", cell) for cell in cells)
    assert "w:type=\"page\"" in xml
