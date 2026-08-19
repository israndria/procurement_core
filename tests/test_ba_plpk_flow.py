from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

from pypdf import PdfWriter

from gabung_ba_pljkk import deteksi_file, gabung
from document_profiles import inject_header_profile
from word_merge import _patch_plpk_layout_xml, _resolve_ba_kind


def _blank_pdf(path: Path):
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    with path.open("wb") as output:
        writer.write(output)


def test_deteksi_file_is_scoped_to_ba_type(tmp_path):
    _blank_pdf(tmp_path / "BA_PLPK_GPR_P.Bng.pdf")
    _blank_pdf(tmp_path / "BA_PLJKK_OLD.pdf")

    plpk = deteksi_file(str(tmp_path), "PLPK")
    jkk = deteksi_file(str(tmp_path), "PLJKK")

    assert Path(plpk["ba_utama"]).name == "BA_PLPK_GPR_P.Bng.pdf"
    assert plpk["kode"] == "GPR_P.Bng"
    assert Path(jkk["ba_utama"]).name == "BA_PLJKK_OLD.pdf"
    assert jkk["kode"] == "OLD"


def test_gabung_plpk_writes_plpk_output_name(tmp_path):
    _blank_pdf(tmp_path / "BA_PLPK_GPR_P.Bng.pdf")

    result = gabung(str(tmp_path), "PLPK")

    assert result["ok"] is True
    assert Path(result["output"]).name == "BA_PLPK_GPR_P.Bng.pdf"
    assert Path(result["output"]).parent.name == "7. Berita Acara + Summary Non Tender"
    assert Path(result["output"]).read_bytes() == (tmp_path / "BA_PLPK_GPR_P.Bng.pdf").read_bytes()


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
