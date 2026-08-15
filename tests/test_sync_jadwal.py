import pandas as pd

import sync_jadwal


class _Request:
    def __init__(self, payload=None, error=None):
        self.payload = payload or {}
        self.error = error

    def execute(self):
        if self.error:
            raise RuntimeError(self.error)
        return self.payload


class _Events:
    def __init__(self, events=None):
        self.items = {event["id"]: event for event in (events or [])}
        self.next_id = 1
        self.update_calls = 0
        self.insert_calls = 0
        self.delete_calls = 0

    def list(self, **kwargs):
        items = list(self.items.values())
        private = kwargs.get("privateExtendedProperty", "")
        if private:
            key, value = private.split("=", 1)
            items = [
                event for event in items
                if event.get("extendedProperties", {}).get("private", {}).get(key) == value
            ]
        query = kwargs.get("q")
        if query:
            items = [event for event in items if query in (event.get("description") or "")]
        return _Request({"items": items})

    def update(self, *, eventId, body, **kwargs):
        self.update_calls += 1
        self.items[eventId] = {"id": eventId, **body}
        return _Request({"id": eventId})

    def insert(self, *, body, **kwargs):
        self.insert_calls += 1
        event_id = f"new-{self.next_id}"
        self.next_id += 1
        self.items[event_id] = {"id": event_id, **body}
        return _Request({"id": event_id})

    def delete(self, *, eventId, **kwargs):
        self.delete_calls += 1
        self.items.pop(eventId, None)
        return _Request({})


class _Service:
    def __init__(self, events=None):
        self.api = _Events(events)

    def events(self):
        return self.api


def _frame():
    return pd.DataFrame([
        {
            "Tahap": "Pengumuman",
            "Mulai": "14 Agustus 2026 11:00",
            "Sampai": "15 Agustus 2026 11:00",
            "Perubahan": "Tidak Ada",
            "Nama_Paket": "Paket Uji",
        }
    ])


def test_reconcile_updates_existing_legacy_event_without_delete_first():
    url = "https://spse.inaproc.id/tapinkab/lelang/123/jadwal"
    old = {
        "id": "old-1",
        "summary": "Pengumuman - Paket Uji",
        "description": f"Link: {url}",
        "start": {"dateTime": "2026-08-14T10:00:00"},
    }
    service = _Service([old])

    result = sync_jadwal.reconcile_tender_events(service, _frame(), url, "045")

    assert result["ok"] is True
    assert result["updated"] == 1
    assert result["inserted"] == 0
    assert result["deleted"] == 0
    assert service.api.update_calls == 1
    assert service.api.delete_calls == 0
    assert service.api.items["old-1"]["extendedProperties"]["private"]["source_tender"] == "123"


def test_owned_tender_rows_only_contains_allowlisted_codes(monkeypatch):
    db = pd.DataFrame([{
        "url": "https://spse.inaproc.id/tapinkab/lelang/1/jadwal",
        "members": "001",
        "nama_paket": "Lama",
        "last_sync": "",
        "content_hash": "",
    }, {
        "url": "https://spse.inaproc.id/tapinkab/lelang/2/jadwal",
        "members": "002",
        "nama_paket": "Bukan target",
        "last_sync": "",
        "content_hash": "",
    }])
    monkeypatch.setattr(sync_jadwal, "_load_supabase_tender_rows", lambda codes=None: [{
        "kode_tender": "1",
        "kode_pokja": "045",
        "nama_tender": "Target",
    }])

    result = sync_jadwal._owned_tender_rows(db, [{
        "jenis_paket": "tender",
        "kode_paket": "1",
        "nama_paket": "Target",
    }])

    assert result["url"].tolist() == [
        "https://spse.inaproc.id/tapinkab/lelang/1/jadwal"
    ]
    assert result.iloc[0]["nama_paket"] == "Target"
