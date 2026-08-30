import zipfile

from lxml import etree

import patch_tender_mailmerge
import word_merge


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _field(name, value="0"):
    return (
        f'<w:fldSimple w:instr=" MERGEFIELD {name} \\* MERGEFORMAT ">'
        f'<w:r><w:t>{value}</w:t></w:r></w:fldSimple>'
    )


def _row(slot, family="Pembukaan", value="0"):
    return (
        "<w:tr>"
        f"<w:tc><w:p>{_field(f'{family}_No_{slot}', str(slot))}</w:p></w:tc>"
        f"<w:tc><w:p>{_field(f'{family}_Nama_{slot}', value)}</w:p></w:tc>"
        "</w:tr>"
    )


def _write_docx(path):
    xml = (
        f'<w:document xmlns:w="{W}"><w:body><w:tbl>'
        '<w:tr><w:tc><w:p><w:r><w:t>PESERTA</w:t></w:r></w:p></w:tc></w:tr>'
        + _row(1, value="CV AKTIF")
        + _row(2)
        + _row(10)
        + _row(3, family="Administrasi")
        + "</w:tbl></w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", xml.encode("utf-8"))


def test_xml_blank_rows_supports_dynamic_families_and_slot_10(tmp_path):
    path = tmp_path / "merge.docx"
    _write_docx(path)

    data = {
        "Pembukaan_Nama_1": "CV AKTIF",
        "Pembukaan_Nama_2": "",
        "Pembukaan_Nama_10": "0",
        "Administrasi_Nama_3": "",
    }
    word_merge._blank_empty_participant_rows_xml(str(path), data)

    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")

    assert b"MERGEFIELD Pembukaan_No_1" in xml
    assert b"MERGEFIELD Pembukaan_Nama_1" in xml
    assert b"MERGEFIELD Pembukaan_No_2" not in xml
    assert b"MERGEFIELD Pembukaan_No_10" not in xml
    assert b"MERGEFIELD Administrasi_No_3" not in xml
    assert b">0<" not in xml


def test_mail_merge_field_names_are_normalized_for_word_engine():
    node = etree.Element(f"{{{W}}}instrText")

    patch_tender_mailmerge._set_instruction(("instrText", node), "Pembukaan Nama 10")

    assert "MERGEFIELD Pembukaan_Nama_10" in node.text
    assert "MERGEFIELD Pembukaan Nama 10" not in node.text


def test_normalize_existing_word_field_keeps_merge_switch():
    root = etree.fromstring(
        f'<w:document xmlns:w="{W}">'
        '<w:fldSimple w:instr=" MERGEFIELD Pembukaan Nama 10 \\* MERGEFORMAT "/>'
        "</w:document>"
    )

    assert patch_tender_mailmerge._normalize_document_field_names(root) == 1
    assert root[0].get(f"{{{W}}}instr") == (
        " MERGEFIELD Pembukaan_Nama_10 \\* MERGEFORMAT "
    )


def test_populate_row_replaces_legacy_content_in_every_cell():
    row = etree.fromstring(
        f"""
        <w:tr xmlns:w="{W}">
          <w:tc><w:p><w:r><w:t>1.</w:t></w:r></w:p></w:tc>
          <w:tc><w:p>{_field("Peserta_1", "OLD")}</w:p></w:tc>
        </w:tr>
        """
    )

    patch_tender_mailmerge._populate_row(
        row,
        ["Pembukaan No 1", "Pembukaan Nama 1"],
        static_slot=1,
    )
    xml = etree.tostring(row, encoding="unicode")

    assert "Peserta_1" not in xml
    assert "OLD" not in xml
    assert "MERGEFIELD Pembukaan_No_1" in xml
    assert "MERGEFIELD Pembukaan_Nama_1" in xml
