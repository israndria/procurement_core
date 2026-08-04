from pathlib import Path

import pytest

from config import PL_WORKFLOW_REGISTRY, pl_workflow_template_dir
from setup_paket_baru import _setup_folder


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
