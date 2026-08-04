#!/usr/bin/env python3
"""Protect Word signature blocks from page pagination splits.

The helper works at the DOCX XML level so it preserves the donor document's
fonts, colors, tables, relationships, and other package parts. It is safe to
run repeatedly: existing protection flags are kept and no duplicate flags are
added.
"""

from __future__ import annotations

import argparse
import os
import re
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
    "divId", "cnfStyle", "rPr",
]
_PPR_ORDER_INDEX = {name: index for index, name in enumerate(_PPR_ORDER)}

_TRPR_ORDER = [
    "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
    "cantSplit", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins",
    "del", "trPrChange",
]
_TRPR_ORDER_INDEX = {name: index for index, name in enumerate(_TRPR_ORDER)}

_SIGNATURE_START = (
    "pejabatpembuatkomitmen",
    "pejabatpenandatangan",  # handles "P ejabat ..." after squashing
    "pakp",                  # PA/KPA/PPK
    "penggunaanggaran",
    "tandatangan",
    "mengetahuidanmenyetujui",
    "menerimadanmenyetujui",
)

_SIGNATURE_STRONG_ROW = (
    "nama_penyedia_ttd",
    "direktur_ttd",
    "menerima dan menyetujui",
    "untuk dan atas nama penyedia",
    "tanda tangan",
)

_SIGNATURE_STOP = (
    "tembusan",
    "lampiran",
    "distribusi",
)


def _node_text(node) -> str:
    values = node.xpath(
        ".//w:t/text() | .//w:delText/text() | .//w:instrText/text()",
        namespaces=NS,
    )
    return re.sub(r"\s+", " ", " ".join(value.strip() for value in values if value.strip())).strip()


def _squash(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _ensure_ordered_child(parent, local: str, order_index: dict[str, int]):
    existing = parent.find(qn(local))
    if existing is not None:
        # Word stores on/off properties as either a bare element (true) or
        # w:val="0"/"false" (false). Normalize an existing flag to true;
        # otherwise an old false value would make protection ineffective while
        # the idempotency check incorrectly reports it as already present.
        normalized = False
        value_name = qn("val")
        if value_name in existing.attrib:
            existing.attrib.pop(value_name)
            normalized = True
        current_index = parent.index(existing)
        parent.remove(existing)
        wanted = order_index.get(local, 10_000)
        for index, child in enumerate(parent):
            child_order = order_index.get(etree.QName(child).localname, 10_000)
            if child_order > wanted:
                parent.insert(index, existing)
                return existing, normalized or index != current_index
        parent.append(existing)
        return existing, normalized or parent.index(existing) != current_index

    element = etree.Element(qn(local))
    wanted = order_index.get(local, 10_000)
    for index, child in enumerate(parent):
        child_order = order_index.get(etree.QName(child).localname, 10_000)
        if child_order > wanted:
            parent.insert(index, element)
            return element, True
    parent.append(element)
    return element, True


def _protect_paragraph(paragraph, with_next: bool) -> int:
    ppr = paragraph.find(qn("pPr"))
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        paragraph.insert(0, ppr)

    changed = 0
    _, did_add = _ensure_ordered_child(ppr, "keepLines", _PPR_ORDER_INDEX)
    changed += int(did_add)
    if with_next:
        _, did_add = _ensure_ordered_child(ppr, "keepNext", _PPR_ORDER_INDEX)
        changed += int(did_add)
    return changed


def _protect_row(row) -> int:
    trpr = row.find(qn("trPr"))
    if trpr is None:
        trpr = etree.Element(qn("trPr"))
        row.insert(0, trpr)
    _, did_add = _ensure_ordered_child(trpr, "cantSplit", _TRPR_ORDER_INDEX)
    return int(did_add)


def _is_signature_row(text: str, row_index: int, row_count: int) -> bool:
    lowered = text.lower()
    if any(marker in lowered for marker in _SIGNATURE_STRONG_ROW):
        return True

    # PPK signature tables in HPS/Uraian Singkat have a short trailing row
    # containing NAMA_PPK/NIP_PPK. Avoid locking long narrative rows.
    is_trailing = row_index >= max(0, row_count - 2)
    is_short = len(text) <= 900
    if is_trailing and is_short and any(
        marker in lowered
        for marker in (
            "nama_ppk",
            "nip_ppk",
            "nama_pa",
            "nip_pa",
            "pejabat pembuat komitmen",
            "pejabat penandatangan kontrak",
        )
    ):
        return True
    return False


def _protect_signature_rows(root) -> int:
    changed = 0
    for table in root.xpath(".//w:tbl", namespaces=NS):
        rows = table.xpath("./w:tr", namespaces=NS)
        for row_index, row in enumerate(rows):
            row_text = _node_text(row)
            if row_text and _is_signature_row(row_text, row_index, len(rows)):
                changed += _protect_row(row)
    return changed


def _is_signature_start(text: str) -> bool:
    squashed = _squash(text)
    return any(squashed.startswith(marker) for marker in _SIGNATURE_START)


def _is_signature_continuation(text: str) -> bool:
    if not text:
        return True
    lowered = text.lower()
    if any(lowered.startswith(marker) for marker in _SIGNATURE_STOP):
        return False
    if len(text) > 220:
        return False
    # Setelah merge, nama/jabatan/NIP bisa menjadi teks aktual tanpa marker.
    # Blok sudah diawali label signature dan dibatasi stop marker + panjang.
    return True


def _protect_body_signature_chains(root) -> int:
    body = root.find(qn("body"))
    if body is None:
        return 0

    children = list(body)
    changed = 0
    index = 0
    while index < len(children):
        paragraph = children[index]
        if paragraph.tag != qn("p") or not _is_signature_start(_node_text(paragraph)):
            index += 1
            continue

        block = [paragraph]
        cursor = index + 1
        while cursor < len(children) and len(block) < 7:
            candidate = children[cursor]
            if candidate.tag != qn("p"):
                break
            if not _is_signature_continuation(_node_text(candidate)):
                break
            block.append(candidate)
            cursor += 1

        # A start marker alone is not enough. Requiring a continuation avoids
        # changing ordinary narrative paragraphs that happen to mention PPK.
        if len(block) >= 2:
            for block_index, item in enumerate(block):
                changed += _protect_paragraph(item, block_index < len(block) - 1)
            index = cursor
        else:
            index += 1
    return changed


def _protect_xml(document_xml: bytes) -> tuple[bytes, dict[str, int]]:
    root = etree.fromstring(document_xml)
    row_changes = _protect_signature_rows(root)
    paragraph_changes = _protect_body_signature_chains(root)
    stats = {"table_rows": row_changes, "paragraph_flags": paragraph_changes}
    if not any(stats.values()):
        return document_xml, stats
    return (
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
        stats,
    )


def protect_docx(path: str | os.PathLike[str]) -> dict[str, int]:
    """Apply signature pagination protection to one DOCX, idempotently."""
    docx_path = Path(path)
    if not docx_path.is_file():
        raise FileNotFoundError(docx_path)

    temp_path = docx_path.with_suffix(docx_path.suffix + ".signature-layout.tmp")
    changed_parts: dict[str, bytes] = {}
    with zipfile.ZipFile(docx_path, "r") as source:
        if "word/document.xml" not in source.namelist():
            raise ValueError(f"DOCX tidak memiliki word/document.xml: {docx_path}")
        original_xml = source.read("word/document.xml")
        changed_xml, stats = _protect_xml(original_xml)
        if changed_xml == original_xml:
            return stats
        changed_parts["word/document.xml"] = changed_xml
        try:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    target.writestr(item, changed_parts.get(item.filename, source.read(item.filename)))
        except Exception:
            if temp_path.exists():
                temp_path.unlink()
            raise

    # Windows tidak mengizinkan replace target ketika arsip sumber masih
    # terbuka. Validasi dan replace dilakukan setelah with ZipFile selesai.
    try:
        with zipfile.ZipFile(temp_path, "r") as check:
            if check.testzip() is not None:
                raise ValueError(f"DOCX hasil proteksi rusak: {docx_path}")
        os.replace(temp_path, docx_path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise
    return stats


def _is_backup(path: Path) -> bool:
    name = path.name.lower()
    return (
        name.startswith("~$")
        or ".bak." in name
        or ".pre-" in name
        or ".before-" in name
        or name.endswith(".tmp.docx")
    )


def protect_folder(folder: str | os.PathLike[str]) -> list[tuple[Path, dict[str, int]]]:
    folder_path = Path(folder)
    if not folder_path.is_dir():
        raise NotADirectoryError(folder_path)
    results = []
    for path in sorted(folder_path.glob("*.docx")):
        if _is_backup(path):
            continue
        results.append((path, protect_docx(path)))
    return results


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--folder", action="append", required=True, help="Folder template DOCX")
    args = parser.parse_args()
    for folder in args.folder:
        for path, stats in protect_folder(folder):
            print(f"PROTECTED|{path}|{stats}")


if __name__ == "__main__":
    main()
