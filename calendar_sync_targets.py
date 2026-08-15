"""Allowlist eksplisit paket yang boleh disinkronkan ke Google Calendar.

Fail-closed: jika tabel/secret tidak tersedia, scheduler tidak menebak paket
dari tabel sumber. Paket harus lebih dulu didaftarkan sebagai target aktif.
"""

from __future__ import annotations

import argparse
import os
import re
from pathlib import Path

import requests


VALID_KINDS = {"tender", "pl"}
DEFAULT_SCOPE = "POKJA2026"
TABLE = "calendar_sync_targets"


class TargetRegistryError(RuntimeError):
    """Registry tidak dapat dibaca atau ditulis."""


def folder_identity_matches(
    folder_name: str,
    package_code: str,
    roots: list[str] | tuple[str, ...],
    sheet_name: str,
    cells: tuple[str, ...],
) -> bool:
    """Validasi kode paket dari workbook Master Data di folder lokal."""
    expected = re.sub(r"\D", "", str(package_code or "").strip())
    raw_folder = os.path.normpath(str(folder_name or "").strip().strip('"'))
    if not expected or not raw_folder:
        return False

    if os.path.isabs(raw_folder):
        folders = [raw_folder]
    else:
        name = os.path.basename(raw_folder)
        if not name or name in (".", ".."):
            return False
        folders = [
            os.path.join(os.path.normpath(root), name)
            for root in roots
            if str(root or "").strip()
        ]

    try:
        from openpyxl import load_workbook
    except ImportError:
        return False

    for folder in dict.fromkeys(folders):
        if not os.path.isdir(folder):
            continue
        try:
            filenames = sorted(os.listdir(folder))
        except OSError:
            continue
        for filename in filenames:
            lower = filename.lower()
            if (
                lower.startswith("~$")
                or ".bak_" in lower
                or ".pre-" in lower
                or not lower.endswith((".xlsx", ".xlsm"))
            ):
                continue
            path = os.path.join(folder, filename)
            workbook = None
            try:
                workbook = load_workbook(
                    path,
                    read_only=True,
                    data_only=True,
                )
                sheet = workbook[sheet_name]
                values = [sheet[cell].value for cell in cells]
            except (OSError, KeyError, ValueError, TypeError):
                continue
            finally:
                if workbook is not None:
                    workbook.close()
            if any(re.sub(r"\D", "", str(value or "").strip()) == expected for value in values):
                return True
    return False


def _read_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    try:
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    except OSError:
        pass
    return values


def _supabase_config() -> tuple[str, str]:
    base_dir = Path(__file__).resolve().parent
    roots = [
        os.environ.get("POKJA_SECRET_ROOT", ""),
        str(base_dir.parent / "Secrets"),
        os.environ.get("LOCALAPPDATA", "")
        and str(Path(os.environ["LOCALAPPDATA"]) / "POKJA2026" / "Secrets"),
    ]
    env: dict[str, str] = {}
    for root in dict.fromkeys(os.path.normpath(r) for r in roots if r):
        env.update(_read_env_file(Path(root) / "secret_supabase.env"))
    url = (os.environ.get("SUPABASE_URL") or env.get("SUPABASE_URL", "")).strip().rstrip("/")
    key = (os.environ.get("SUPABASE_KEY") or env.get("SUPABASE_KEY", "")).strip()
    if not url or not key:
        raise TargetRegistryError("Supabase secret untuk calendar_sync_targets tidak tersedia.")
    return url, key


def _kind(kind: str) -> str:
    value = str(kind or "").strip().lower()
    if value not in VALID_KINDS:
        raise ValueError(f"jenis_paket harus salah satu dari: {', '.join(sorted(VALID_KINDS))}")
    return value


def _code(code: str) -> str:
    value = str(code or "").strip()
    if not value or not value.isdigit():
        raise ValueError("kode_paket harus berupa kode numerik SPSE.")
    return value


def _request(method: str, *, scope: str, **kwargs):
    url, key = _supabase_config()
    headers = kwargs.pop("headers", {})
    headers.update({"apikey": key, "Authorization": f"Bearer {key}"})
    response = requests.request(
        method,
        f"{url}/rest/v1/{TABLE}",
        headers=headers,
        timeout=30,
        **kwargs,
    )
    if response.status_code in (404, 406) or response.status_code >= 500:
        raise TargetRegistryError(
            f"Registry {TABLE} belum siap atau gagal di Supabase (HTTP {response.status_code})."
        )
    try:
        response.raise_for_status()
    except requests.RequestException as exc:
        raise TargetRegistryError(f"Registry {TABLE} ditolak Supabase: {exc}") from exc
    return response


def load_targets(
    kind: str | None = None,
    *,
    scope: str | None = None,
    enabled_only: bool = True,
) -> list[dict]:
    """Ambil target aktif dari allowlist; tidak pernah fallback ke semua paket."""
    params = {
        "select": "scope,jenis_paket,kode_paket,nama_paket,folder_name,enabled,source,note",
        "scope": f"eq.{scope or os.environ.get('POKJA_CALENDAR_SCOPE', DEFAULT_SCOPE)}",
        "order": "jenis_paket,kode_paket",
    }
    if enabled_only:
        params["enabled"] = "eq.true"
    if kind is not None:
        params["jenis_paket"] = f"eq.{_kind(kind)}"
    response = _request("GET", scope=scope or DEFAULT_SCOPE, params=params)
    rows = response.json()
    if not isinstance(rows, list):
        raise TargetRegistryError("Respons registry bukan array.")
    return rows


def upsert_target(
    kind: str,
    code: str,
    *,
    name: str = "",
    folder_name: str = "",
    source: str = "manual",
    note: str = "",
    enabled: bool = True,
    scope: str | None = None,
) -> dict:
    """Aktifkan satu paket secara eksplisit."""
    target_scope = scope or os.environ.get("POKJA_CALENDAR_SCOPE", DEFAULT_SCOPE)
    payload = {
        "scope": target_scope,
        "jenis_paket": _kind(kind),
        "kode_paket": _code(code),
        "nama_paket": str(name or "").strip() or None,
        "folder_name": str(folder_name or "").strip() or None,
        "enabled": bool(enabled),
        "source": str(source or "manual").strip(),
        "note": str(note or "").strip() or None,
    }
    response = _request(
        "POST",
        scope=target_scope,
        params={"on_conflict": "scope,jenis_paket,kode_paket"},
        headers={"Prefer": "resolution=merge-duplicates,return=representation"},
        json=payload,
    )
    body = response.json()
    return body[0] if isinstance(body, list) and body else payload


def disable_target(kind: str, code: str, *, scope: str | None = None) -> None:
    """Nonaktifkan target tanpa menghapus histori/identitasnya."""
    target_scope = scope or os.environ.get("POKJA_CALENDAR_SCOPE", DEFAULT_SCOPE)
    _request(
        "PATCH",
        scope=target_scope,
        params={
            "scope": f"eq.{target_scope}",
            "jenis_paket": f"eq.{_kind(kind)}",
            "kode_paket": f"eq.{_code(code)}",
        },
        headers={"Prefer": "return=minimal"},
        json={"enabled": False},
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Kelola allowlist scheduler Google Calendar")
    sub = parser.add_subparsers(dest="action", required=True)
    enable = sub.add_parser("enable")
    enable.add_argument("kind", choices=sorted(VALID_KINDS))
    enable.add_argument("code")
    enable.add_argument("--name", default="")
    enable.add_argument("--note", default="")
    disable = sub.add_parser("disable")
    disable.add_argument("kind", choices=sorted(VALID_KINDS))
    disable.add_argument("code")
    sub.add_parser("list")
    args = parser.parse_args()
    if args.action == "enable":
        print(upsert_target(args.kind, args.code, name=args.name, note=args.note))
    elif args.action == "disable":
        disable_target(args.kind, args.code)
        print(f"disabled {args.kind}/{args.code}")
    else:
        for row in load_targets():
            print(row)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
