import zipfile

from setup_paket_baru import _word_link_is_valid, link_word_to_excel


def test_link_word_to_excel_repairs_missing_settings_relationships(tmp_path):
    word_path = tmp_path / "BA.docx"
    excel_path = tmp_path / "Paket.xlsm"
    excel_path.write_bytes(b"placeholder")

    settings = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b"</w:settings>"
    )
    with zipfile.ZipFile(word_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/settings.xml", settings)

    assert not _word_link_is_valid(word_path)
    assert link_word_to_excel(word_path, str(excel_path)) is True
    assert _word_link_is_valid(word_path)

    with zipfile.ZipFile(word_path) as archive:
        rels = archive.read("word/_rels/settings.xml.rels").decode("utf-8")
    assert "mailMergeSource" in rels
    assert "Paket.xlsm" in rels
