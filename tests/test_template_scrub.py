from template_scrub import scrub_excel_pl_copy


def test_plpk_scrub_does_not_touch_label_or_panel_columns(monkeypatch, tmp_path):
    captured = {}

    def fake_clear(path, targets):
        captured["path"] = path
        captured["targets"] = targets
        return []

    monkeypatch.setattr("template_scrub._clear_constants_com", fake_clear)
    workbook = tmp_path / "0. BAPLPK- Paket.xlsm"
    workbook.write_bytes(b"placeholder")

    scrub_excel_pl_copy(workbook, is_pk=True)

    ranges = captured["targets"]["@ Master Data"]
    assert ranges == [
        "C3:C10", "C13:C28", "C30:C31", "C33:C64",
        "C66:C75", "C77:C80", "C87:C89", "H8:H10",
    ]
    assert all(not item.startswith("F") for item in ranges)
    assert all(not item.startswith("B") for item in ranges)


def test_pljkk_scrub_does_not_touch_label_or_panel_columns(monkeypatch, tmp_path):
    captured = {}

    def fake_clear(path, targets):
        captured["targets"] = targets
        return []

    monkeypatch.setattr("template_scrub._clear_constants_com", fake_clear)
    workbook = tmp_path / "0. BAPLJKK- Paket.xlsm"
    workbook.write_bytes(b"placeholder")

    scrub_excel_pl_copy(workbook, is_pk=False)

    ranges = captured["targets"]["@ Master Data"]
    assert ranges == ["C3:C10", "C13:C28", "C30:C54", "C61:C63", "H8:H10"]
    assert all(not item.startswith("F") for item in ranges)
    assert all(not item.startswith("B") for item in ranges)
