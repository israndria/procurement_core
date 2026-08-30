"""Patch tender Word templates to consume the canonical ``satu_data`` blocks.

The tender workbook already exposes one row with ten participant slots.  The
historical Word donors still contain three legacy rows and cached sample
fields.  This utility changes only ``word/document.xml`` in the selected
DOCX/DOCM files, preserves every other ZIP entry, and keeps the original in a
dated backup before the first write.

It is intentionally data-agnostic: participant names, invitation status, and
the active count are supplied by Excel at merge time.
"""

from __future__ import annotations

import argparse
import copy
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
BACKUP_SUFFIX = ".pre-mailmerge-20260830.bak"


def _visible_text(node) -> str:
    return "".join((item.text or "") for item in node.iter(W + "t"))


def _field_instructions(node):
    """Return merge-field instruction nodes/attributes in document order."""
    result = []
    for item in node.iter():
        if item.tag == W + "instrText":
            text = item.text or ""
            if "MERGEFIELD" in text.upper():
                result.append(("instrText", item))
        elif item.tag == W + "fldSimple":
            text = item.get(W + "instr", "")
            if "MERGEFIELD" in text.upper():
                result.append(("fldSimple", item))
    return result


def _normalized_merge_field_name(name: str) -> str:
    """Maak Word field name eenduidig en veilig voor field-code parsing."""
    return re.sub(r"[^A-Za-z0-9_]", "", re.sub(r"\s+", "_", str(name).strip()))


def _set_instruction(item, name: str) -> None:
    name = _normalized_merge_field_name(name)
    instruction = f" MERGEFIELD {name} " + r"\* MERGEFORMAT "
    kind, node = item
    if kind == "instrText":
        node.text = instruction
    else:
        node.set(W + "instr", instruction)


def _normalize_document_field_names(root) -> int:
    """Normalisasi semua MERGEFIELD existing agar engine tidak split spasi."""
    changed = 0
    for kind, node in _field_instructions(root):
        raw = node.text if kind == "instrText" else node.get(W + "instr", "")
        prefix = re.split(r"\s+\\", raw or "", maxsplit=1)[0]
        match = re.search(r"MERGEFIELD\s+(.+)$", prefix, re.I)
        if not match:
            continue
        name = _normalized_merge_field_name(match.group(1))
        if not name:
            continue
        instruction = f" MERGEFIELD {name} " + r"\* MERGEFORMAT "
        if instruction == raw:
            continue
        if kind == "instrText":
            node.text = instruction
        else:
            node.set(W + "instr", instruction)
        changed += 1
    return changed


def _add_simple_field(cell, name: str, clear_visible: bool = False) -> None:
    """Add a Word simple merge field while inheriting the first run style."""
    name = _normalized_merge_field_name(name)
    source_run = cell.find(".//" + W + "r")
    source_rpr = source_run.find(W + "rPr") if source_run is not None else None
    source_paragraph = cell.find(".//" + W + "p")
    source_ppr = source_paragraph.find(W + "pPr") if source_paragraph is not None else None
    tc_pr = cell.find(W + "tcPr")
    for child in list(cell):
        cell.remove(child)
    if tc_pr is not None:
        cell.append(tc_pr)
    paragraph = etree.SubElement(cell, W + "p")
    if source_ppr is not None:
        paragraph.append(copy.deepcopy(source_ppr))
    field = etree.Element(W + "fldSimple")
    field.set(W + "instr", f" MERGEFIELD {name} " + r"\* MERGEFORMAT ")
    run = etree.SubElement(field, W + "r")
    if source_rpr is not None:
        run.append(copy.deepcopy(source_rpr))
    text = etree.SubElement(run, W + "t")
    text.text = f"«{name}»"
    paragraph.append(field)


def _set_static_cell_text(cell, value: str) -> None:
    """Ganti isi cell dengan teks statis, tetap mempertahankan style dasar."""
    source_run = cell.find(".//" + W + "r")
    source_rpr = source_run.find(W + "rPr") if source_run is not None else None
    source_paragraph = cell.find(".//" + W + "p")
    source_ppr = source_paragraph.find(W + "pPr") if source_paragraph is not None else None
    tc_pr = cell.find(W + "tcPr")
    for child in list(cell):
        cell.remove(child)
    if tc_pr is not None:
        cell.append(tc_pr)
    paragraph = etree.SubElement(cell, W + "p")
    if source_ppr is not None:
        paragraph.append(copy.deepcopy(source_ppr))
    run = etree.SubElement(paragraph, W + "r")
    if source_rpr is not None:
        run.append(copy.deepcopy(source_rpr))
    text = etree.SubElement(run, W + "t")
    text.text = value


def _set_cell_field(cell, name: str | None, clear_visible: bool = False) -> None:
    if not name:
        return
    _add_simple_field(cell, name, clear_visible=clear_visible)


def _direct_cells(row):
    return row.findall(W + "tc")


def _populate_row(row, fields: list[str | None], static_slot: int | None = None) -> None:
    cells = _direct_cells(row)
    for index, name in enumerate(fields):
        if index >= len(cells):
            continue
        if name:
            _set_cell_field(cells[index], name, clear_visible=index == 0)
        elif index == 0 and static_slot is not None:
            _set_static_cell_text(cells[index], f"{static_slot}.")
        else:
            _set_static_cell_text(cells[index], "")


def _row_fields(row):
    return " ".join(
        (item.text or item.get(W + "instr", ""))
        for _kind, item in _field_instructions(row)
    ).upper()


def _replace_tail_rows(table, start: int, fields_for_slot) -> int:
    """Replace all rows after a section header with ten styled rows."""
    rows = table.findall(W + "tr")
    if start >= len(rows):
        return 0
    donor = copy.deepcopy(rows[start])
    for row in rows[start:]:
        table.remove(row)
    for slot in range(1, 11):
        row = copy.deepcopy(donor)
        _populate_row(row, fields_for_slot(slot), static_slot=slot)
        table.insert(start + slot - 1, row)
    return 10


def _find_header_row(table, predicate):
    rows = table.findall(W + "tr")
    for index, row in enumerate(rows):
        text = re.sub(r"\s+", " ", _visible_text(row)).strip().upper()
        if predicate(text):
            return rows, index
    return rows, None


def _patch_summary_table(table) -> int:
    labels = {
        "NAMA PESERTA TENDER": "Ringkasan Nama",
        "HARGA PENAWARAN TERKOREKSI": "Ringkasan Harga Terkoreksi",
        "HARGA PENAWARAN": "Ringkasan Harga Penawaran",
        "EVALUASI ADMINISTRASI": "Ringkasan Administrasi",
        "EVALUASI KUALIFIKASI": "Ringkasan Kualifikasi",
        "EVALUASI TEKNIS": "Ringkasan Teknis",
        "EVALUASI HARGA": "Ringkasan Evaluasi Harga",
        "KESIMPULAN": "Ringkasan Kesimpulan",
    }
    rows = table.findall(W + "tr")
    if "RINGKASAN_NAMA_10" in _row_fields(table):
        return 0
    start = None
    base_rows = []
    for index, row in enumerate(rows):
        text = re.sub(r"\s+", " ", _visible_text(row)).strip().upper()
        if start is None and "NAMA PESERTA TENDER" in text:
            start = index
        if start is not None:
            label = next((key for key in labels if key in text), None)
            if label:
                base_rows.append((copy.deepcopy(row), labels[label]))
            if "KESIMPULAN" in text:
                break
    if start is None or len(base_rows) != len(labels):
        return 0

    for row in rows[start : start + len(base_rows)]:
        table.remove(row)
    for slot in range(1, 11):
        for donor, prefix in base_rows:
            row = copy.deepcopy(donor)
            cells = _direct_cells(row)
            if len(cells) >= 3:
                _set_cell_field(cells[2], f"{prefix} {slot}")
            table.insert(start, row)
            start += 1
    return 10


def _patch_evaluation_tables(root) -> dict[str, int]:
    counts = {"summary": 0, "opening": 0, "admin": 0, "technical": 0, "price": 0}
    for table in root.findall(".//" + W + "tbl"):
        text = re.sub(r"\s+", " ", _visible_text(table)).strip().upper()
        rows = table.findall(W + "tr")
        if not rows:
            continue
        if "NAMA PESERTA TENDER" in text and "HARGA PENAWARAN" in text:
            counts["summary"] += _patch_summary_table(table)
            continue

        if "KOREKSI ARITMATIK" in text and "PENAWARAN TERHADAP HPS" in text:
            if "PEMBUKAAN_NAMA_10" in _row_fields(table):
                continue
            _rows, header = _find_header_row(table, lambda x: "KOREKSI ARITMATIK" in x)
            if header is not None:
                counts["opening"] += _replace_tail_rows(
                    table,
                    header + 1,
                    lambda i: [
                        f"Pembukaan No {i}",
                        f"Pembukaan Nama {i}",
                        f"Pembukaan Alamat {i}",
                        f"Pembukaan Harga Penawaran {i}",
                        f"Pembukaan Koreksi Aritmatik {i}",
                        f"Pembukaan Persen HPS {i}",
                        f"Pembukaan Keterangan {i}",
                    ],
                )
            continue

        if "DOKUMEN PENAWARAN TEKNIS" in text and "SUB KONTRAK" in text:
            if "ADMINISTRASI_NAMA_10" in _row_fields(table):
                continue
            _rows, header = _find_header_row(table, lambda x: "PESERTA" in x and "DOKUMEN PENAWARAN" in x)
            if header is not None:
                counts["admin"] += _replace_tail_rows(
                    table,
                    header + 2,
                    lambda i: [
                        f"Administrasi No {i}",
                        f"Administrasi Nama {i}",
                        f"Administrasi Metode {i}",
                        f"Administrasi Peralatan {i}",
                        f"Administrasi Personil {i}",
                        f"Administrasi Subkontrak {i}",
                        f"Administrasi RKK {i}",
                        f"Administrasi Dokumen Harga {i}",
                        f"Administrasi Hasil {i}",
                    ],
                )
            continue

        if "PERSYARATAN TEKNIS" in text and "PERSONEL MANAJERIAL" in text:
            if "TEKNIS_NAMA_10" in _row_fields(table):
                continue
            _rows, header = _find_header_row(table, lambda x: "PERSYARATAN TEKNIS" in x)
            if header is not None:
                counts["technical"] += _replace_tail_rows(
                    table,
                    header + 2,
                    lambda i: [
                        None,
                        f"Teknis Nama {i}",
                        f"Teknis Peralatan {i}",
                        f"Teknis Personel {i}",
                        f"Teknis RKK {i}",
                        f"Teknis Hasil {i}",
                    ],
                )
            continue

        if "KEWAJARAN HARGA" in text and "HARGA PENAWARAN TERKOREKSI" in text:
            if "HARGA_NAMA_10" in _row_fields(table):
                continue
            _rows, header = _find_header_row(table, lambda x: "HARGA PENAWARAN" in x and "KEWAJARAN" in x)
            if header is not None:
                counts["price"] += _replace_tail_rows(
                    table,
                    header + 2,
                    lambda i: [
                        None,
                        f"Harga Nama {i}",
                        f"Harga Penawaran {i}",
                        f"Harga Penawaran Terkoreksi {i}",
                        f"Harga Persen HPS {i}",
                        f"Harga Kewajaran {i}",
                        f"Harga Kesimpulan {i}",
                    ],
                )
    return counts


def _patch_proof_tables(root) -> dict[str, int]:
    counts = {"proof": 0, "attendance": 0}
    for table in root.findall(".//" + W + "tbl"):
        text = re.sub(r"\s+", " ", _visible_text(table)).strip().upper()
        rows = table.findall(W + "tr")
        if not rows:
            continue
        if "PEMBUKTIAN_NAMA_10" in _row_fields(table):
            continue
        if "NAMA PESERTA" in text and "CATATAN" in text:
            _rows, header = _find_header_row(table, lambda x: "NAMA PESERTA" in x and "CATATAN" in x)
            if header is not None:
                counts["proof"] += _replace_tail_rows(
                    table,
                    header + 1,
                    lambda i: [
                        f"Pembuktian No {i}",
                        f"Pembuktian Nama {i}",
                        f"Pembuktian Hasil {i}",
                        f"Pembuktian Keterangan {i}",
                        f"Pembuktian Catatan {i}",
                    ],
                )
            continue

        if "TANDA TANGAN + CAP PERUSAHAAN" in text and "NAMA" in text:
            _rows, header = _find_header_row(table, lambda x: "TANDA TANGAN + CAP PERUSAHAAN" in x)
            if header is not None:
                counts["attendance"] += _replace_tail_rows(
                    table,
                    header + 1,
                    lambda i: [
                        f"Pembuktian No {i}",
                        f"Pembuktian Nama {i}",
                        None,
                        None,
                        None,
                    ],
                )
    return counts


def patch_document(path: Path, mode: str, *, backup: bool = True) -> dict:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)
    backup_path = path.with_name(path.name + BACKUP_SUFFIX)
    if backup and not backup_path.exists():
        shutil.copy2(path, backup_path)

    with zipfile.ZipFile(path, "r") as source:
        original_xml = source.read("word/document.xml")
        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(original_xml, parser)
        counts = {}
        if mode in {"summary", "all"}:
            counts.update(_patch_evaluation_tables(root))
        if mode in {"proof", "all"}:
            counts.update(_patch_proof_tables(root))
        counts["normalized_fields"] = _normalize_document_field_names(root)
        if not any(counts.values()):
            return {
                "file": str(path),
                "changed": False,
                "counts": counts,
                "backup": str(backup_path) if backup_path.exists() else None,
            }
        changed_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        fd, temp_name = tempfile.mkstemp(prefix=".tender-mailmerge-", suffix=".tmp", dir=path.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(temp_name, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    payload = changed_xml if item.filename == "word/document.xml" else source.read(item.filename)
                    target.writestr(item, payload)
        except Exception:
            if os.path.exists(temp_name):
                os.remove(temp_name)
            if backup_path.exists():
                shutil.copy2(backup_path, path)
            raise

    # The source archive must be closed before replacing the file on Windows.
    try:
        os.replace(temp_name, path)
    except Exception:
        if os.path.exists(temp_name):
            os.remove(temp_name)
        if backup_path.exists():
            shutil.copy2(backup_path, path)
        raise
    return {"file": str(path), "changed": True, "counts": counts, "backup": str(backup_path)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("summary", "proof", "all"), required=True)
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    for raw in args.paths:
        result = patch_document(Path(raw), args.mode, backup=not args.no_backup)
        print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
