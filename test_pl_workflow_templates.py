from pathlib import Path
import zipfile

import pytest
from lxml import etree

from config import PL_WORKFLOW_REGISTRY, pl_workflow_template_dir
from setup_paket_baru import _setup_folder
from word_merge import _prepare_dokpil_equipment_docx


def _populate_template_dir(root: Path, workflow: str) -> Path:
    cfg = PL_WORKFLOW_REGISTRY[workflow]
    root.mkdir(parents=True, exist_ok=True)
    (root / cfg["excel_template"]).touch()
    for name, _sheet in cfg["word_map"]:
        (root / name).touch()
    return root


def test_incomplete_v2_falls_back_to_complete_legacy(tmp_path):
    package_root = tmp_path / "Paket Experiment - Pengadaan Langsung"
    v2 = package_root / "V2 - Template PL" / "Konstruksi"
    legacy = package_root / "Development - PL - PK"
    cfg = PL_WORKFLOW_REGISTRY["PL_KONSTRUKSI"]

    (v2 / cfg["excel_template"]).parent.mkdir(parents=True, exist_ok=True)
    (v2 / cfg["excel_template"]).touch()
    (v2 / cfg["word_map"][0][0]).touch()
    _populate_template_dir(legacy, "PL_KONSTRUKSI")

    assert Path(pl_workflow_template_dir("PL_KONSTRUKSI", root=str(package_root))) == legacy


def test_konstruksi_registry_matches_complete_v2_donor():
    cfg = PL_WORKFLOW_REGISTRY["PL_KONSTRUKSI"]
    assert [name for name, _sheet in cfg["word_map"]] == [
        "1. BA Reviu PLPK - Template.docx",
        "2. Isi Reviu PLPK - Template.docm",
        "3. Dokpil Full PK - Template.docx",
        "5. BA PLPK - Template.docx",
        "7. BA Dengan Timpang PLPK - Template.docx",
    ]
    assert [sheet for _name, sheet in cfg["word_map"]] == [
        "satu_data", "list_reviu", "list_dokpil", "satu_data", "satu_data"
    ]


def test_setup_preflight_does_not_create_partial_folder(tmp_path):
    source = tmp_path / "incomplete"
    output = tmp_path / "output"
    source.mkdir()

    with pytest.raises(FileNotFoundError, match="Template workflow tidak lengkap"):
        _setup_folder(
            "1. PLPK - Paket Uji",
            source,
            "0. BAPLPK- Template.xlsm",
            [("1. Dokumen Wajib.docx", "satu_data")],
            output_base=output,
            workflow="PL_KONSTRUKSI",
        )

    assert not (output / "1. PLPK - Paket Uji").exists()


def test_equipment_markers_fill_all_nested_tables(tmp_path):
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = "{" + ns + "}"

    def element(tag, text=None):
        node = etree.Element(w + tag)
        if text is not None:
            p = etree.SubElement(node, w + "p")
            run = etree.SubElement(p, w + "r")
            etree.SubElement(run, w + "t").text = text
        return node

    body = etree.Element(w + "body")
    for _ in range(2):
        outer = etree.SubElement(body, w + "tbl")
        outer_row = etree.SubElement(outer, w + "tr")
        outer_cell = etree.SubElement(outer_row, w + "tc")
        inner = etree.SubElement(outer_cell, w + "tbl")
        header = etree.SubElement(inner, w + "tr")
        for value in ("No", "Jenis", "Kapasitas (Minimal)", "Jumlah"):
            header.append(element("tc", value))
        donor = etree.SubElement(inner, w + "tr")
        for value in ("[[NO_ALAT]]", "[[NAMA_ALAT]]", "[[KAPASITAS_ALAT]]", "[[JUMLAH_ALAT]]"):
            donor.append(element("tc", value))

    document = etree.Element(w + "document", nsmap={"w": ns})
    document.append(body)
    target = tmp_path / "isi_reviu.docm"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8"))

    _prepare_dokpil_equipment_docx(
        str(target),
        {
            "_dokpil_equipment": [
                {"no": "1", "jenis": "Generator Set", "kapasitas": "Minimal 1 Kva", "jumlah": "1 Unit"},
                {"no": "2", "jenis": "Pick Up", "kapasitas": "1 - 1,5 Ton", "jumlah": "1 Unit"},
            ]
        },
    )

    with zipfile.ZipFile(target) as archive:
        result = etree.fromstring(archive.read("word/document.xml"))
    tables = list(result.iter(w + "tbl"))
    assert len(tables) == 4  # two outer + two nested tables
    nested = [
        table
        for table in tables
        if len([child for child in table if child.tag == w + "tr"]) == 3
    ]
    assert len(nested) == 2
    assert all("[[NO_ALAT]]" not in "".join(table.itertext()) for table in nested)
    assert all(len([child for child in table if child.tag == w + "tr"]) == 3 for table in nested)
    merged_text = "".join(result.itertext())
    assert merged_text.count("Generator Set") == 2
    assert merged_text.count("Pick Up") == 2
