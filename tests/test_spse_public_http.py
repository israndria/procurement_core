import requests

import spse_public_http


class _Response:
    def __init__(self, status_code, headers=None):
        self.status_code = status_code
        self.headers = headers or {}


class _Session:
    def __init__(self, responses):
        self.responses = iter(responses)
        self.calls = 0

    def get(self, *args, **kwargs):
        self.calls += 1
        response = next(self.responses)
        if isinstance(response, Exception):
            raise response
        return response


def test_get_public_retries_transient_status(monkeypatch):
    session = _Session([_Response(503, {"Retry-After": "7"}), _Response(200)])
    delays = []
    monkeypatch.setattr(spse_public_http.time, "sleep", delays.append)

    response = spse_public_http.get_public(session, "https://example.test/jadwal")

    assert response.status_code == 200
    assert session.calls == 2
    assert delays == [7]


def test_get_public_uses_plain_requests_once_after_cloudscraper_block(monkeypatch):
    primary = _Session([_Response(403), _Response(403), _Response(403)])
    fallback = _Session([_Response(200)])
    delays = []
    monkeypatch.setattr(spse_public_http.time, "sleep", delays.append)

    response = spse_public_http.get_public(
        primary,
        "https://example.test/jadwal",
        fallback=fallback,
    )

    assert response.status_code == 200
    assert primary.calls == 3
    assert fallback.calls == 1
    assert delays == [2, 4]


def test_get_public_retries_network_error(monkeypatch):
    session = _Session([requests.ConnectionError("offline"), _Response(200)])
    delays = []
    monkeypatch.setattr(spse_public_http.time, "sleep", delays.append)

    response = spse_public_http.get_public(session, "https://example.test/jadwal")

    assert response.status_code == 200
    assert session.calls == 2
    assert delays == [2]
