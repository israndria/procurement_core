"""Jalankan sync jadwal Tender + PL dari Task Scheduler lokal."""

from __future__ import annotations

import os
import sys
from datetime import datetime


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNTIME_ROOT = os.path.normpath(
    os.environ.get(
        "POKJA_RUNTIME_ROOT",
        os.path.join(
            os.environ.get("LOCALAPPDATA", os.path.expanduser("~/AppData/Local")),
            "POKJA2026",
            "Asisten_Pokja",
        ),
    )
)
LOG_PATH = os.path.join(RUNTIME_ROOT, "logs", "sync_jadwal_all.log")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def log(message: str) -> None:
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    print(line)
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _asisten_root() -> str:
    configured = os.environ.get("POKJA_CODE_ROOT", "").strip().strip('"')
    if configured and os.path.isfile(os.path.join(configured, "gcal_pl_helper.py")):
        return os.path.normpath(configured)
    return os.path.normpath(os.path.join(os.path.dirname(BASE_DIR), "Asisten_Pokja"))


def run() -> int:
    """Best-effort sync dua sumber; return 1 jika ada kegagalan nyata."""
    failures = 0

    try:
        from sync_jadwal import sync_all

        tender = sync_all()
        failures += int(tender.get("failed", 0))
        log(
            "Tender selesai: "
            f"updated={tender.get('updated', 0)}, "
            f"unchanged={tender.get('unchanged', 0)}, "
            f"failed={tender.get('failed', 0)}"
        )
    except Exception as exc:
        failures += 1
        log(f"Tender gagal: {exc}")

    try:
        asisten_root = _asisten_root()
        if asisten_root not in sys.path:
            sys.path.insert(0, asisten_root)
        from gcal_pl_helper import sync_semua_paket_pl

        results = sync_semua_paket_pl(skip_unchanged=True)
        empty = [
            row for row in results
            if "Jadwal kosong di SPSE" in row.get("error", "")
        ]
        failed = [
            row for row in results
            if not row.get("ok") and "Jadwal kosong di SPSE" not in row.get("error", "")
        ]
        skipped = sum(1 for row in results if row.get("skipped"))
        # Satu paket SPSE error tidak boleh membatalkan sync paket lain.
        # Task gagal hanya jika tidak ada satu pun paket PL yang berhasil/skip.
        if failed and not any(row.get("ok") for row in results):
            failures += 1
        log(
            "PL PK/PLJKK selesai: "
            f"total={len(results)}, skipped={skipped}, "
            f"jadwal_kosong={len(empty)}, failed={len(failed)}"
        )
        for row in failed:
            log(f"PL gagal {row.get('kode_paket', '?')}: {row.get('error', '-')}")
    except Exception as exc:
        failures += 1
        log(f"PL PK/PLJKK gagal: {exc}")

    log("Run selesai: " + ("GAGAL" if failures else "OK"))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(run())
