"""HTTP helper terbatas untuk endpoint publik SPSE."""

from __future__ import annotations

import time
from collections.abc import Callable

import requests


RETRYABLE_HTTP_STATUS = frozenset({403, 408, 425, 429, 500, 502, 503, 504})
MAX_ATTEMPTS = 3


def _delay(response, attempt: int) -> int:
    retry_after = getattr(response, "headers", {}).get("Retry-After", "")
    try:
        return max(2, min(30, int(float(retry_after))))
    except (TypeError, ValueError):
        return min(30, 2 ** (attempt + 1))


def _notify(log_fn: Callable[[str], object] | None, message: str) -> None:
    if log_fn is None:
        return
    try:
        log_fn(message)
    except Exception:
        pass


def get_public(
    session,
    url: str,
    *,
    fallback=None,
    log_fn: Callable[[str], object] | None = None,
    attempts: int = MAX_ATTEMPTS,
    **kwargs,
):
    """GET endpoint publik SPSE dengan retry ringan dan fallback satu kali."""
    attempts = max(1, int(attempts))
    last_response = None
    last_error = None

    for attempt in range(attempts):
        try:
            response = session.get(url, **kwargs)
        except requests.RequestException as exc:
            last_error = exc
            if attempt + 1 >= attempts:
                break
            delay = min(30, 2 ** (attempt + 1))
            _notify(
                log_fn,
                f"  HTTP network error {url} - retry {attempt + 1}/{attempts - 1} dalam {delay}s",
            )
            time.sleep(delay)
            continue

        last_response = response
        if response.status_code not in RETRYABLE_HTTP_STATUS:
            return response
        if attempt + 1 >= attempts:
            break
        delay = _delay(response, attempt)
        _notify(
            log_fn,
            f"  HTTP {response.status_code} {url} - retry {attempt + 1}/{attempts - 1} dalam {delay}s",
        )
        time.sleep(delay)

    # cloudscraper kadang gagal di host tertentu; coba requests biasa sekali.
    # Retry tetap dibatasi agar tidak memperburuk rate-limit SPSE.
    if (
        fallback is not None
        and fallback is not session
        and (last_response is None or last_response.status_code in (403, 429))
    ):
        try:
            return fallback.get(url, **kwargs)
        except requests.RequestException as exc:
            if last_response is None:
                raise exc from last_error

    if last_response is not None:
        return last_response
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"GET SPSE tidak menghasilkan respons: {url}")
