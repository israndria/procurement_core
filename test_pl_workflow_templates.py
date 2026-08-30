from pathlib import Path
import zipfile

import pytest
from lxml import etree

from config import PL_WORKFLOW_REGISTRY, pl_workflow_template_dir
from document_profiles import strip_static_headers
from setup_paket_baru import _setup_folder, _win_extended_path, _write_setup_status
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


def test_copy_failure_quarantines_new_partial_folder(tmp_path, monkeypatch):
    source = _populate_template_dir(tmp_path / "template", "PL_KONSTRUKSI")
    output = tmp_path / "output"

    def fail_copy(*_args, **_kwargs):
        raise FileNotFoundError("simulated source race")

    monkeypatch.setattr("setup_paket_baru._copy2_retry", fail_copy)
    with pytest.raises(FileNotFoundError, match="simulated source race"):
        _setup_folder(
            "1. PLPK - Paket Race",
            source,
            PL_WORKFLOW_REGISTRY["PL_KONSTRUKSI"]["excel_template"],
            PL_WORKFLOW_REGISTRY["PL_KONSTRUKSI"]["word_map"],
            output_base=output,
            workflow="PL_KONSTRUKSI",
        )

    assert not (output / "1. PLPK - Paket Race").exists()
    assert len(list((output / "_setup-failed").iterdir())) == 1


def test_complete_setup_clears_stale_failure_timestamp(tmp_path):
    package_dir = tmp_path / "paket"
    package_dir.mkdir()
    meta = package_dir / ".template-meta.json"
    meta.write_text(
        '{"setup_status":"failed","failed_at":"2026-08-30T07:26:08"}',
        encoding="utf-8",
    )

    _write_setup_status(
        str(package_dir), "complete", completed_at="2026-08-30T07:27:44"
    )

    data = __import__("json").loads(meta.read_text(encoding="utf-8"))
    assert data["setup_status"] == "complete"
    assert "failed_at" not in data
    assert data["completed_at"] == "2026-08-30T07:27:44"


def test_long_windows_path_uses_extended_namespace():
    path = "C:\\" + ("x" * 260)
    if __import__("os").name == "nt":
        assert _win_extended_path(path).startswith("\\\\?\\")
    else:
        assert _win_extended_path(path) == path


@pytest.mark.parametrize(
    ("workflow", "folder_name"),
    [
        ("PL_KONSTRUKSI", "1. PLPK - Paket Konstruksi"),
        ("PL_PERENCANAAN", "1. PLJKK - Paket Perencanaan"),
    ],
)
def test_setup_pl_provisions_revision_upload_folder(
    tmp_path, monkeypatch, workflow, folder_name
):
    monkeypatch.setattr("document_profiles.is_official_header_document", lambda _path: False)
    monkeypatch.setattr("template_scrub.scrub_excel_pl_copy", lambda *_args, **_kwargs: [])
    monkeypatch.setattr("setup_paket_baru.link_word_to_excel", lambda *_args, **_kwargs: True)
    source = _populate_template_dir(tmp_path / "template", workflow)
    output = tmp_path / "output"
    cfg = PL_WORKFLOW_REGISTRY[workflow]

    _setup_folder(
        folder_name,
        source,
        cfg["excel_template"],
        cfg["word_map"],
        output_base=output,
        workflow=workflow,
    )

    package_dir = output / folder_name
    assert (package_dir / "10. Revisi Uploadan PPK").is_dir()
    xml_data = package_dir / "11. XML Data"
    assert xml_data.is_dir()
    assert list(xml_data.iterdir()) == []
    assert __import__("json").loads(
        (package_dir / ".template-meta.json").read_text(encoding="utf-8")
    )["setup_status"] == "complete"


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
        for value in ("No", "Nama Peralatan", "Kapasitas", "Jumlah (Unit/Buah)"):
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


def test_strip_static_headers_preserves_docx_structure(tmp_path):
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = "{" + ns + "}"
    target = tmp_path / "1. BA Reviu DPP PLJKK - Paket.docx"
    document_xml = etree.Element(w + "document", nsmap={"w": ns})
    etree.SubElement(document_xml, w + "body")
    header_xml = etree.Element(w + "hdr", nsmap={"w": ns})
    paragraph = etree.SubElement(header_xml, w + "p")
    run = etree.SubElement(paragraph, w + "r")
    etree.SubElement(run, w + "t").text = "DINAS PEKERJAAN UMUM DAN PENATAAN RUANG"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", etree.tostring(document_xml))
        archive.writestr("word/header1.xml", etree.tostring(header_xml))

    strip_static_headers(target)

    with zipfile.ZipFile(target) as archive:
        result = etree.fromstring(archive.read("word/header1.xml"))
        assert "".join(result.itertext()) == ""
        assert archive.testzip() is None
        assert archive.read("word/document.xml") == etree.tostring(document_xml)
