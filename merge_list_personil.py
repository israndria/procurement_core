"""Merge tabel personil JKK dari Excel Master Data PPK V2 ke DOCX marker."""

from __future__ import annotations

import copy
import sys
import zipfile
from pathlib import Path

from lxml import etree

from master_data_v2 import read_domain_data


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"


def _text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces={"w": W_NS}))


def _replace(node: etree._Element, mapping: dict[str, str]) -> None:
    for text_node in node.xpath(".//w:t", namespaces={"w": W_NS}):
        if text_node.text is None:
            continue
        for old, new in mapping.items():
            if old in text_node.text:
                text_node.text = text_node.text.replace(old, new)


def merge_list_personil(template_path: str | Path, excel_path: str | Path, output_path: str | Path) -> Path:
    """Isi tabel marker JKK; template asli tidak ditimpa."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    rows = read_domain_data(excel_path)["jkk_personil"]

    with zipfile.ZipFile(template_path, "r") as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        tables = root.xpath(".//w:tbl", namespaces={"w": W_NS})
        target = next((table for table in tables if "[[JABATAN]]" in _text(table)), None)
        if target is None:
            raise ValueError("Template tidak memiliki row marker [[JABATAN]]")
        tr_list = target.xpath("./w:tr", namespaces={"w": W_NS})
        if len(tr_list) < 2:
            raise ValueError("Template tabel personil harus memiliki header dan template row")
        template_row = tr_list[1]
        for row in tr_list[2:]:
            target.remove(row)

        for index, person in enumerate(rows, 1):
            row = copy.deepcopy(template_row)
            mapping = {
                "[[NO]]": str(person.get("No") or index),
                "[[JABATAN]]": str(person.get("Jabatan", "")),
                "[[SERTIFIKAT]]": str(person.get("Sertifikat", "")),
                "[[PENGALAMAN]]": str(person.get("Pengalaman Kerja", "")),
            }
            _replace(row, mapping)
            target.append(row)
        target.remove(template_row)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    return output_path


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit("Usage: python merge_list_personil.py TEMPLATE.docx DATA.xlsm OUTPUT.docx")
    print(merge_list_personil(sys.argv[1], sys.argv[2], sys.argv[3]))
