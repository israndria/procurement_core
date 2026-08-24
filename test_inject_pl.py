from pathlib import Path

import pytest

from inject_pl import MOD_NAME, _validate_vba_source, find_bapljkk_files


def test_snapshot_buttons_use_requested_labels_and_current_macros():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")

    assert '"Save Data", "SaveDataPL"' in source
    assert '"Load Data", "LoadDataPL"' in source
    assert '"Save Input Data"' not in source
    assert '"Load Input Data"' not in source


def test_default_discovery_includes_jkk_and_pk_and_skips_backups(tmp_path):
    (tmp_path / "jkk").mkdir()
    (tmp_path / "pk").mkdir()
    (tmp_path / "jkk" / "0. BAPLJKK - Paket.xlsm").touch()
    (tmp_path / "pk" / "0. BAPLPK - Paket.xlsm").touch()
    (tmp_path / "pk" / "0. BAPLPK - Paket.bak.xlsm").touch()
    (tmp_path / "pk" / "~$0. BAPLPK - Paket.xlsm").touch()

    found = find_bapljkk_files(str(tmp_path))

    assert found == sorted(
        [
            str(tmp_path / "jkk" / "0. BAPLJKK - Paket.xlsm"),
            str(tmp_path / "pk" / "0. BAPLPK - Paket.xlsm"),
        ]
    )


def test_validate_vba_source_accepts_vba_double_quotes():
    source = (
        f'Attribute VB_Name = "{MOD_NAME}"\n'
        'Public Sub Probe()\n'
        '    Range("A1").Formula = "=IF(B1="""","""",B1)"\n'
        "End Sub\n"
    )

    _validate_vba_source(source)


def test_validate_vba_source_rejects_python_style_formula_escape():
    source = (
        f'Attribute VB_Name = "{MOD_NAME}"\n'
        'Public Sub Probe()\n'
        '    Range("A1").Formula = "=IF(B1=\\"\\",B1)"\n'
        "End Sub\n"
    )

    with pytest.raises(ValueError, match="escape Python/JSON"):
        _validate_vba_source(source)


def test_snapshot_keeps_numeric_text_as_text():
    source = Path(__file__).with_name("ModDraftPaketPL.bas").read_text(encoding="utf-8")

    assert "VarType(cellValue) <> vbString" in source
