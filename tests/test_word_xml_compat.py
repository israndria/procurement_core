from lxml import etree

from word_xml_compat import MAIN_NS, normalize_word_document_xml


def test_restore_word_namespace_map_from_lxml_ns0_document():
    broken = (
        b'<?xml version="1.0"?>'
        b'<ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        b'xmlns:ns1="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'ns1:Ignorable="w14 w15"><ns0:body/></ns0:document>'
    )

    repaired = normalize_word_document_xml(broken)
    root = etree.fromstring(repaired)

    assert root.nsmap["w"] == MAIN_NS
    assert root.nsmap["mc"] == "http://schemas.openxmlformats.org/markup-compatibility/2006"
    assert root.nsmap["w14"]
    assert root.nsmap["w15"]
    assert b"xmlns:w=" in repaired
    assert b"xmlns:ns0=" not in repaired


def test_canonical_document_is_left_unchanged():
    xml = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>'

    assert normalize_word_document_xml(xml) == xml
