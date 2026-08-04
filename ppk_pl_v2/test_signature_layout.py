from __future__ import annotations

import hashlib
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))
from signature_layout import protect_docx  # noqa: E402


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _make_fixture(path: Path) -> None:
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}"><w:body>
  <w:p><w:pPr><w:rPr/><w:keepNext/></w:pPr><w:r><w:t>PEJABAT PEMBUAT KOMITMEN,</w:t></w:r></w:p>
  <w:p><w:r><w:t>«NAMA_PPK»</w:t></w:r></w:p>
  <w:p><w:r><w:t>NIP. «NIP_PPK»</w:t></w:r></w:p>
  <w:p><w:r><w:t>Pengguna Anggaran,</w:t></w:r></w:p>
  <w:p><w:r><w:t>Nama PA Aktual</w:t></w:r></w:p>
  <w:p><w:r><w:t>NIP. 198001012006041001</w:t></w:r></w:p>
  <w:p><w:r><w:t>Tembusan Yth.</w:t></w:r></w:p>
  <w:tbl><w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr><w:tc><w:p><w:r><w:t>Untuk dan atas nama Penyedia</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>[Nama Penyedia] [Direktur]</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:sectPr/></w:body></w:document>'''.encode("utf-8")
    content_types = b'''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", content_types)
        package.writestr("word/document.xml", document)


class SignatureLayoutTests(unittest.TestCase):
    def test_protects_paragraph_chain_and_signature_row(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "fixture.docx"
            _make_fixture(path)
            stats = protect_docx(path)
            self.assertGreaterEqual(stats["table_rows"], 1)
            self.assertGreaterEqual(stats["paragraph_flags"], 1)
            with zipfile.ZipFile(path) as package:
                root = etree.fromstring(package.read("word/document.xml"))
                paragraphs = root.xpath("./w:body/w:p", namespaces=NS)
                self.assertIsNotNone(paragraphs[0].find("./w:pPr/w:keepLines", namespaces=NS))
                self.assertIsNotNone(paragraphs[0].find("./w:pPr/w:keepNext", namespaces=NS))
                ppr_names = [etree.QName(node).localname for node in paragraphs[0].find("./w:pPr", namespaces=NS)]
                self.assertLess(ppr_names.index("keepNext"), ppr_names.index("rPr"))
                pa_paragraph = paragraphs[3]
                self.assertIsNotNone(pa_paragraph.find("./w:pPr/w:keepLines", namespaces=NS))
                row = root.xpath(".//w:tbl/w:tr", namespaces=NS)[0]
                self.assertIsNotNone(row.find("./w:trPr/w:cantSplit", namespaces=NS))
                trpr_names = [etree.QName(node).localname for node in row.find("./w:trPr", namespaces=NS)]
                self.assertLess(trpr_names.index("cantSplit"), trpr_names.index("tblHeader"))
                self.assertIsNone(root.xpath(".//w:trPr/w:cantSplit", namespaces=NS)[0].get("w:val"))
                self.assertIsNone(package.testzip())

    def test_is_idempotent(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "fixture.docx"
            _make_fixture(path)
            protect_docx(path)
            before = hashlib.sha256(path.read_bytes()).digest()
            stats = protect_docx(path)
            after = hashlib.sha256(path.read_bytes()).digest()
            self.assertEqual(stats, {"table_rows": 0, "paragraph_flags": 0})
            self.assertEqual(before, after)

    def test_normalizes_false_existing_flags(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "fixture.docx"
            _make_fixture(path)
            with zipfile.ZipFile(path, "r") as package:
                root = etree.fromstring(package.read("word/document.xml"))
            first_ppr = root.xpath("./w:body/w:p[1]/w:pPr", namespaces=NS)[0]
            first_ppr.find("./w:keepNext", namespaces=NS).set(f"{{{W_NS}}}val", "0")
            row_trpr = root.xpath(".//w:tr/w:trPr", namespaces=NS)[0]
            row_trpr.find("./w:cantSplit", namespaces=NS).set(f"{{{W_NS}}}val", "false")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
                package.writestr("[Content_Types].xml", b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>")
                package.writestr("word/document.xml", etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))

            stats = protect_docx(path)
            self.assertGreaterEqual(stats["table_rows"], 1)
            self.assertGreaterEqual(stats["paragraph_flags"], 1)
            with zipfile.ZipFile(path) as package:
                root = etree.fromstring(package.read("word/document.xml"))
                self.assertIsNone(root.xpath("./w:body/w:p[1]/w:pPr/w:keepNext", namespaces=NS)[0].get(f"{{{W_NS}}}val"))
                self.assertIsNone(root.xpath(".//w:tr/w:trPr/w:cantSplit", namespaces=NS)[0].get(f"{{{W_NS}}}val"))
                self.assertIsNone(package.testzip())


if __name__ == "__main__":
    unittest.main()
