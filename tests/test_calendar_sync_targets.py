import calendar_sync_targets
from openpyxl import Workbook


class _Response:
    def __init__(self, payload):
        self.payload = payload
        self.status_code = 200

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


def test_load_targets_reads_only_active_scope(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", "https://example.supabase.co")
    monkeypatch.setenv("SUPABASE_KEY", "test-key")
    calls = []

    def fake_request(method, url, **kwargs):
        calls.append((method, url, kwargs))
        return _Response([{"jenis_paket": "tender", "kode_paket": "123", "enabled": True}])

    monkeypatch.setattr(calendar_sync_targets.requests, "request", fake_request)
    rows = calendar_sync_targets.load_targets("tender")

    assert rows[0]["kode_paket"] == "123"
    assert calls[0][2]["params"]["scope"] == "eq.POKJA2026"
    assert calls[0][2]["params"]["enabled"] == "eq.true"
    assert calls[0][2]["params"]["jenis_paket"] == "eq.tender"


def test_code_must_be_numeric():
    try:
        calendar_sync_targets.upsert_target("pl", "bukan-kode")
    except ValueError as exc:
        assert "numerik" in str(exc)
    else:
        raise AssertionError("kode non-numerik seharusnya ditolak")


def test_folder_identity_matches_tender_master_code(tmp_path):
    folder = tmp_path / "Paket Tender"
    folder.mkdir()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "@ Master Data"
    sheet["C4"] = "123456"
    workbook.save(folder / "Paket.xlsm")

    assert calendar_sync_targets.folder_identity_matches(
        "Paket Tender", "123456", [str(tmp_path)], "@ Master Data", ("C4",)
    )
    assert not calendar_sync_targets.folder_identity_matches(
        "Paket Tender", "999999", [str(tmp_path)], "@ Master Data", ("C4",)
    )


def test_folder_identity_accepts_pl_legacy_f2_fallback(tmp_path):
    folder = tmp_path / "Paket PL"
    folder.mkdir()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "@ Master Data"
    sheet["F2"] = "987654"
    workbook.save(folder / "Paket.xlsx")

    assert calendar_sync_targets.folder_identity_matches(
        "Paket PL", "987654", [str(tmp_path)], "@ Master Data", ("C3", "F2")
    )
