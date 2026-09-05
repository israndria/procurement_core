from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_pl_injector_includes_dynamic_nego_layout_modules_and_events():
    source = (ROOT / "inject_pl.py").read_text(encoding="utf-8")
    assert '"modBarisItem"' in source
    assert '"modAutoLayoutNego"' in source
    assert "BEGIN POKJA_AUTO_BARIS_ITEM" in source
    assert "BEGIN POKJA_AUTO_LAYOUT_NEGO" in source
    assert "modAutoLayoutNego.RapikanDaftarNego True, False" in source
    assert "IFERROR(terbilang1" in source
    assert "terbilang(YEAR" in source


def test_pl_injector_saves_xml_snapshot_on_ctrl_s_and_fails_closed():
    source = (ROOT / "inject_pl.py").read_text(encoding="utf-8")
    assert "Private Sub Workbook_BeforeSave" in source
    assert "ModDraftPaketPL.SaveDataPL" in source
    assert "Cancel = True" in source
    assert "Penyimpanan dibatalkan: snapshot XML gagal dibuat." in source


def test_pl_injector_layout_module_sources_are_validated_and_imported():
    source = (ROOT / "inject_pl.py").read_text(encoding="utf-8")
    assert 'layout_bas = SCRIPT_DIR / f"{layout_name}.bas"' in source
    assert "_validate_vba_source(layout_content, layout_name)" in source
    assert "imported_layout.Name = layout_name" in source


def test_nego_layout_modules_do_not_use_tender_only_fixed_rows():
    for name in ("modBarisItem.bas", "modAutoLayoutNego.bas"):
        source = (ROOT / name).read_text(encoding="utf-8")
        assert "Private Const FIRST_ITEM_ROW As Long = 9" not in source
        assert "Private Const LAST_ITEM_ROW As Long = 508" not in source


def test_nego_layout_skips_duplicate_measurement_when_signature_is_unchanged():
    source = (ROOT / "modAutoLayoutNego.bas").read_text(encoding="utf-8")
    assert "If Not ForceRefresh Then" in source
    assert "If BuildLayoutSignature(ws) = mLastSignature Then GoTo CleanExit" in source


def test_nego_layout_never_treats_total_and_signature_footer_as_item_rows():
    for name in ("modBarisItem.bas", "modAutoLayoutNego.bas"):
        source = (ROOT / name).read_text(encoding="utf-8")
        assert "FindFooterRow" in source
        assert "lastRow = footerRow - 1" in source
        assert "EnsureFooterVisible" in source
        assert "lastRow = firstRow + MAX_ITEMS - 1" not in source
