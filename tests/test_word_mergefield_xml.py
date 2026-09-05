from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from word_merge import _extract_mergefield_names_docx, validate_merge_source_fields


def test_mergefield_preflight_reads_split_fields_and_header(tmp_path: Path):
    document = (
        b'<?xml version="1.0"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body><w:p>'
        b'<w:r><w:instrText>MER</w:instrText></w:r>'
        b'<w:r><w:instrText>GEFIELD Nama_Paket</w:instrText></w:r>'
        b'</w:p><w:sectPr/></w:body></w:document>'
    )
    header = (
        b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:p><w:fldSimple w:instr=" MERGEFIELD Tanggal_Dokumen "/></w:p></w:hdr>'
    )
    template = tmp_path / "template.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("word/header1.xml", header)

    names = _extract_mergefield_names_docx(template)

    assert set(names) == {"Nama_Paket", "Tanggal_Dokumen"}
    assert validate_merge_source_fields(
        template,
        {"Nama_Paket": "Paket uji", "Tanggal_Dokumen": "1 September 2026"},
    )


def test_mergefield_preflight_rejects_missing_split_field(tmp_path: Path):
    document = (
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:body><w:p><w:r><w:instrText>MERGEFIELD NPWP</w:instrText></w:r></w:p>'
        b'<w:sectPr/></w:body></w:document>'
    )
    template = tmp_path / "missing.docx"
    with ZipFile(template, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)

    try:
        validate_merge_source_fields(template, {"Nama_Paket": "Paket uji"})
    except ValueError as exc:
        assert "NPWP" in str(exc)
    else:
        raise AssertionError("field tanpa sumber harus ditolak")
