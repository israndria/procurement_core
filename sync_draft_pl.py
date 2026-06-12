"""
sync_draft_pl.py — Simpan/ambil data_snapshot dari Supabase draft_paket_pl.
Dipanggil VBA via WScript.Shell.

Mode:
  python sync_draft_pl.py save   — baca _sync_draft_pl_input.json → PATCH data_snapshot
  python sync_draft_pl.py load   — baca kode_paket dari input → return data_snapshot as JSON

Input file : _sync_draft_pl_input.json  (ditulis VBA, dihapus setelah dibaca)
Output file: _sync_draft_pl_output.json (dibaca VBA)
"""

import sys
import os
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH  = os.path.join(SCRIPT_DIR, "_sync_draft_pl_input.json")
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "_sync_draft_pl_output.json")

from dotenv import load_dotenv
load_dotenv(os.path.join(SCRIPT_DIR, "secret_supabase.env"))
SB_URL = os.environ.get("SUPABASE_URL", "")
SB_KEY = os.environ.get("SUPABASE_KEY", "")


def _http_patch(kode_paket: str, payload: dict) -> dict:
    import urllib.request
    url = f"{SB_URL}/rest/v1/draft_paket_pl?kode_paket=eq.{kode_paket}"
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="PATCH")
    req.add_header("apikey", SB_KEY)
    req.add_header("Authorization", f"Bearer {SB_KEY}")
    req.add_header("Content-Type", "application/json")
    req.add_header("Prefer", "return=minimal")
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return {"ok": True, "status": resp.status}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _http_get_snapshot(kode_paket: str) -> dict | None:
    import urllib.request
    url = f"{SB_URL}/rest/v1/draft_paket_pl?kode_paket=eq.{kode_paket}&select=data_snapshot"
    req = urllib.request.Request(url)
    req.add_header("apikey", SB_KEY)
    req.add_header("Authorization", f"Bearer {SB_KEY}")
    req.add_header("Accept", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            rows = json.loads(resp.read().decode("utf-8"))
            if rows:
                return rows[0].get("data_snapshot")
            return None
    except Exception:
        return None


def _read_input() -> dict:
    with open(INPUT_PATH, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
    os.remove(INPUT_PATH)
    return data


def _write_output(result: dict):
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)


def cmd_save():
    inp = _read_input()
    kode_paket = inp.get("kode_paket", "")
    snapshot   = inp.get("snapshot", {})

    if not kode_paket:
        _write_output({"ok": False, "error": "kode_paket kosong"})
        return

    patch_payload = {"data_snapshot": snapshot}
    kode_unik = inp.get("kode_unik", "")
    if kode_unik:
        patch_payload["kode_unik"] = kode_unik
    result = _http_patch(kode_paket, patch_payload)
    _write_output(result)


def cmd_load():
    inp = _read_input()
    kode_paket = inp.get("kode_paket", "")

    if not kode_paket:
        _write_output({"ok": False, "error": "kode_paket kosong"})
        return

    snapshot = _http_get_snapshot(kode_paket)
    if snapshot is None:
        _write_output({"ok": True, "snapshot": {}})
    else:
        _write_output({"ok": True, "snapshot": snapshot})


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    try:
        if mode == "save":
            cmd_save()
        elif mode == "load":
            cmd_load()
        else:
            _write_output({"ok": False, "error": f"mode tidak dikenal: {mode}"})
    except Exception as e:
        _write_output({"ok": False, "error": str(e)})
