from pathlib import Path

import pytest

from inject_pl import (
    MOD_NAME,
    NEGO_EVENT_CODE,
    PLPK_BUTTON_GEOMETRY,
    WORKBOOK_OPEN_DATE_CODE,
    _is_backup_workbook_path,
    _is_plpk_workbook,
    _is_template_workbook_path,
    _force_template_auto_calculation,
    _validate_vba_source,
    find_bapljkk_files,
)
from template_scrub import clean_workbook_donor_state


def test_snapshot_buttons_use_requested_labels_and_current_macros():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")

    assert '"Save Data", "SaveDataPL"' in source
    assert '"Load Data", "LoadDataPL"' in source
    assert '"Save Input Data"' not in source
    assert '"Load Input Data"' not in source
    assert '"btnCetakTimpangPL"' in source
    assert '"PrintPembuktianTimpangPDF"' in source
    assert "btnCetakTimpangPL" in PLPK_BUTTON_GEOMETRY


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


def test_template_path_is_distinguished_from_numbered_package():
    assert _is_template_workbook_path(
        r"D:\Template PL\Konstruksi\0. BAPLPK- Template.xlsm"
    )
    assert not _is_template_workbook_path(
        r"D:\PK\100. PLPK - Paket\0. BAPLPK - Paket.xlsm"
    )


def test_template_calc_patch_is_zip_minimal_and_preserves_vba(tmp_path):
    import zipfile

    workbook = tmp_path / "template.xlsm"
    with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "xl/workbook.xml",
            '<workbook><calcPr calcMode="manual" calcCompleted="0" calcOnSave="0"/></workbook>',
        )
        z.writestr("xl/vbaProject.bin", b"VBA-SENTINEL")
    _force_template_auto_calculation(workbook)
    with zipfile.ZipFile(workbook) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8")
        assert 'calcMode="auto"' in xml
        assert 'calcOnSave="1"' in xml
        assert z.read("xl/vbaProject.bin") == b"VBA-SENTINEL"
        from xml.etree import ElementTree as ET

        ET.fromstring(xml)


def test_template_injection_does_not_scrub_or_reset_donor_data():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")
    assert "clean_workbook_donor_state" not in source
    assert "_reset_template_number_formulas" not in source
    assert "def _harden_template_udf_formulas" in source
    assert 'IFERROR(terbilang1' in source


def test_donor_cleanup_breaks_links_deletes_stale_names_and_clears_lists():
    class LinkName:
        def __init__(self, name, refers):
            self.Name = name
            self.RefersTo = refers
            self.deleted = False

        def Delete(self):
            self.deleted = True

    class Names:
        def __init__(self, items):
            self.items = items
            self.Count = len(items)

        def Item(self, index):
            return self.items[index - 1]

    class Range:
        def __init__(self):
            self.cleared = False

        def ClearContents(self):
            self.cleared = True

    class Sheet:
        def __init__(self):
            self.range = Range()

        def Range(self, _address):
            return self.range

    class Sheets:
        def __init__(self):
            self.items = {"_DraftPaketList": Sheet(), "_DraftPaketPLList": Sheet()}

        def __call__(self, name):
            return self.items[name]

    class Workbook:
        def __init__(self):
            self.names = Names([
                LinkName("external", "=[1]Sheet!$A$1"),
                LinkName("broken", "=#REF!"),
                LinkName("valid", "=Sheet!$A$1"),
            ])
            self.Names = self.names
            self.Worksheets = Sheets()
            self.links = ["donor.xlsm"]
            self.broken = []

        def LinkSources(self, _link_type):
            return self.links

        def BreakLink(self, Name, Type):
            self.broken.append((Name, Type))

    workbook = Workbook()
    logs = clean_workbook_donor_state(workbook, clear_draft_lists=True)

    assert workbook.broken == [("donor.xlsm", 1)]
    assert workbook.names.items[0].deleted
    assert workbook.names.items[1].deleted
    assert not workbook.names.items[2].deleted
    assert workbook.Worksheets("_DraftPaketList").range.cleared
    assert workbook.Worksheets("_DraftPaketPLList").range.cleared
    assert any("External link diputus" in line for line in logs)


def test_plpk_geometry_matches_konstruksi_template_layout():
    assert PLPK_BUTTON_GEOMETRY["btnBukaDokpil_PL"] == (657.9, 175.0, 130.2, 40.0)
    assert PLPK_BUTTON_GEOMETRY["btnBukaBA_PL"] == (657.9, 218.5, 130.2, 27.0)
    assert PLPK_BUTTON_GEOMETRY["btnMuatHPS_PL"] == (929.0, 249.8, 129.9, 28.2)
    assert PLPK_BUTTON_GEOMETRY["btnSaveInputData"] == (926.7, 280.4, 130.2, 40.0)
    assert PLPK_BUTTON_GEOMETRY["btnLoadInputData"] == (925.7, 323.3, 130.4, 39.4)


def test_plpk_detection_does_not_misclassify_generic_nego_sheet():
    class Cell:
        Value = None

    class Sheet:
        def __init__(self, name):
            self.Name = name

        def Cells(self, _row, _column):
            return Cell()

    class Sheets:
        def __init__(self, names):
            self.items = [Sheet(name) for name in names]
            self.Count = len(self.items)

        def __call__(self, index):
            if isinstance(index, int):
                return self.items[index - 1]
            return next(sheet for sheet in self.items if sheet.Name == index)

    class Workbook:
        pass

    Workbook.Sheets = Sheets(["@ Master Data", "7.2 Dengan Nego"])

    assert not _is_plpk_workbook(Workbook())


def test_plpk_detection_accepts_construction_only_marker():
    class Sheet:
        def __init__(self, name):
            self.Name = name

        def Cells(self, _row, _column):
            return type("Cell", (), {"Value": None})()

    class Sheets:
        Count = 2

        def __call__(self, index):
            return [Sheet("@ Master Data"), Sheet("Harga Timpang")][index - 1]

    class Workbook:
        pass

    Workbook.Sheets = Sheets()

    assert _is_plpk_workbook(Workbook())


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
    assert "Application.CalculateFullRebuild" not in source
    assert "Application.CalculateBeforeSave = False" in source
    assert 'Array("@ Master Data", "5. HPS", "6. Penawaran"' in source
    assert "If Not PrepareWorkbookForMailMerge() Then Exit Sub" in source
    assert "ThisWorkbook.ReadOnly" in source


def test_all_word_merge_paths_prepare_cached_values():
    source = Path(__file__).with_name("ModDraftPaketPL.bas").read_text(encoding="utf-8")

    assert source.count("If Not PrepareWorkbookForMailMerge() Then Exit Sub") == 5
    assert "Private Sub RunMergePL" in source
    assert "Public Sub CetakReviuPlJkkPDF()" in source
    assert "Public Sub CetakBAReviuPLPDF()" in source


def test_injector_never_recalculates_formula_cache_during_structural_injection():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")

    assert "XL_AUTOMATION_SECURITY_LOW = 1" in source
    assert "excel.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW" in source
    assert "excel.EnableEvents = False" in source
    assert "excel.Calculation = XL_CALCULATION_MANUAL" in source
    assert "excel.CalculateBeforeSave = True" not in source
    assert "excel.CalculateFullRebuild()" not in source


def test_nego_event_recalculates_dependent_values_in_manual_mode():
    assert 'Set changed = Intersect(Target, Me.Range("J8:M26"))' in NEGO_EVENT_CODE
    assert 'Me.Range("M8:T29").Calculate' in NEGO_EVENT_CODE
    assert "Private Sub Worksheet_Activate()" in NEGO_EVENT_CODE


def test_source_hardens_blank_master_date_helpers():
    source = Path(__file__).with_name("inject_pl.py").read_text(encoding="utf-8")

    assert "def _harden_master_date_helpers" in source
    assert 'H10=0' in source
    assert 'Helper tanggal @ Master Data dibuat blank-safe' in source


def test_refresh_chain_is_scoped_and_dependency_ordered():
    source = Path(__file__).with_name("ModDraftPaketPL.bas").read_text(encoding="utf-8")
    assert "Public Sub RefreshDerivedPL()" in source
    assert '"7.2 Dengan Nego", "@ Evaluasi", "satu_data"' in source
    assert '"database_reviu", "database_dokpil"' in source
    assert '"6. Penawaran", _' in source
    assert "Application.CalculateFullRebuild" not in source
    assert "ModDraftPaketPL.RefreshDerivedPL" in WORKBOOK_OPEN_DATE_CODE
