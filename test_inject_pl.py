from pathlib import Path

import pytest

from inject_pl import (
    MOD_NAME,
    PLPK_BUTTON_GEOMETRY,
    _is_backup_workbook_path,
    _validate_vba_source,
    find_bapljkk_files,
)


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
    backup_dir = tmp_path / "pk" / ".vba-backup"
    backup_dir.mkdir()
    (backup_dir / "0. BAPLPK - Paket.before-ModDraftPaketPL.xlsm").touch()
    archive_dir = tmp_path / "_backup_archive"
    archive_dir.mkdir()
    (archive_dir / "0. BAPLJKK - Paket Arsip.xlsm").touch()

    found = find_bapljkk_files(str(tmp_path))

    assert found == sorted(
        [
            str(tmp_path / "jkk" / "0. BAPLJKK - Paket.xlsm"),
            str(tmp_path / "pk" / "0. BAPLPK - Paket.xlsm"),
        ]
    )


def test_explicit_backup_workbook_is_detected():
    assert _is_backup_workbook_path(
        r"D:\Paket\.vba-backup\0. BAPLPK - Paket.xlsm"
    )
    assert _is_backup_workbook_path(
        r"D:\Paket\0. BAPLPK - Paket.before-ModDraftPaketPL-20260902.xlsm"
    )
    assert not _is_backup_workbook_path(
        r"D:\Paket\0. BAPLPK - Paket Aktif.xlsm"
    )
    assert not _is_backup_workbook_path(
        r"D:\Paket\0. BAPLPK - Paket Panjang__f6b01f12.xlsm"
    )


def test_plpk_geometry_matches_konstruksi_template_layout():
    assert PLPK_BUTTON_GEOMETRY["btnBukaDokpil_PL"] == (657.9, 175.0, 130.2, 40.0)
    assert PLPK_BUTTON_GEOMETRY["btnBukaBA_PL"] == (657.9, 218.5, 130.2, 27.0)
    assert PLPK_BUTTON_GEOMETRY["btnMuatHPS_PL"] == (929.0, 249.8, 129.9, 28.2)
    assert PLPK_BUTTON_GEOMETRY["btnSaveInputData"] == (926.7, 280.4, 130.2, 40.0)
    assert PLPK_BUTTON_GEOMETRY["btnLoadInputData"] == (925.7, 323.3, 130.4, 39.4)


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


def test_ba_print_aborts_if_recalculation_or_save_fails():
    source = Path(__file__).with_name("ModDraftPaketPL.bas").read_text(encoding="utf-8")

    assert "Private Function PrepareWorkbookForMailMerge() As Boolean" in source
    assert "Application.CalculateFullRebuild" in source
    assert "If Not PrepareWorkbookForMailMerge() Then Exit Sub" in source
    assert "ThisWorkbook.ReadOnly" in source


def test_injector_enforces_fresh_formula_cache_before_save():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")

    assert "excel.Calculation = XL_CALCULATION_AUTOMATIC" in source
    assert "excel.CalculateBeforeSave = True" in source
    assert "excel.CalculateFullRebuild()" in source
