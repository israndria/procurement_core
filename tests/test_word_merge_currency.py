from decimal import Decimal
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

from document_profiles import _ensure_header_reference
from word_merge import (
    _parse_currency_text,
    _strip_pl_ba_signature_header_xml,
    format_value,
)


def test_budget_fields_are_formatted_as_rupiah_for_word_only():
    assert format_value(200000000, "Pagu") == "Rp. 200.000.000,00"
    assert format_value(199984583.13, "HPS") == "Rp. 199.984.583,13"
    assert format_value(Decimal("199984583.13"), "Nilai_HPS") == "Rp. 199.984.583,13"
    assert _parse_currency_text(format_value(199984583.13, "HPS")) == 199984583.13


def test_non_budget_numeric_fields_keep_existing_merge_format():
    assert format_value(199984583.13, "Harga Penawaran") == "199984583.13"


def test_header_injector_preserves_existing_section_mapping():
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    document = f'''<w:document xmlns:w="{ns_w}" xmlns:r="{ns_r}"><w:body>
      <w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId8"/></w:sectPr></w:pPr></w:p>
      <w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr>
    </w:body></w:document>'''.encode()

    result = ET.fromstring(_ensure_header_reference(document, "rId8"))
    refs = result.findall(".//{%s}headerReference" % ns_w)
    assert [ref.get("{%s}id" % ns_r) for ref in refs] == ["rId8", "rId9"]


def test_signature_section_keeps_empty_header_without_com_navigation(tmp_path):
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns_header_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    document = f'''<?xml version="1.0"?>
<w:document xmlns:w="{ns_w}" xmlns:r="{ns_r}"><w:body>
  <w:p/><w:p><w:pPr><w:sectPr>
    <w:headerReference w:type="default" r:id="rId8"/>
  </w:sectPr></w:pPr></w:p>
  <w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr>
</w:body></w:document>'''.encode()
    rels = f'''<?xml version="1.0"?>
<Relationships xmlns="{ns_rel}">
  <Relationship Id="rId8" Type="{ns_header_rel}" Target="header1.xml"/>
  <Relationship Id="rId9" Type="{ns_header_rel}" Target="header2.xml"/>
</Relationships>'''.encode()
    header1 = f'<w:hdr xmlns:w="{ns_w}"><w:p><w:r><w:t>HEADER</w:t></w:r></w:p></w:hdr>'.encode()
    header2 = f'<w:hdr xmlns:w="{ns_w}"><w:p><w:r><w:t>HEADER</w:t></w:r></w:p></w:hdr>'.encode()
    target = tmp_path / "1. BA Reviu PLPK - target.docx"
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("word/_rels/document.xml.rels", rels)
        archive.writestr("word/header1.xml", header1)
        archive.writestr("word/header2.xml", header2)

    assert _strip_pl_ba_signature_header_xml(target) is True

    with ZipFile(target) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        refs = root.findall(
            ".//{%s}headerReference" % ns_w
        )
        assert [ref.get("{%s}id" % ns_r) for ref in refs] == ["rId8", "rId9"]
        assert b"HEADER" in archive.read("word/header1.xml")
        assert b"HEADER" not in archive.read("word/header2.xml")
        assert archive.testzip() is None
