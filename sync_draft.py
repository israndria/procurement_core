"""
sync_draft.py — Simpan/ambil data_snapshot dari Supabase draft_paket.
Dipanggil VBA via WScript.Shell.

Mode:
  python sync_draft.py save   — baca _sync_draft_input.json → upsert data_snapshot
  python sync_draft.py load   — baca kode_tender dari input → return data_snapshot as JSON

Input file : _sync_draft_input.json  (ditulis VBA, dihapus setelah dibaca)
Output file: _sync_draft_output.json (dibaca VBA)
"""

import sys
import os
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH  = os.path.join(SCRIPT_DIR, "_sync_draft_input.json")
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "_sync_draft_output.json")

# Supabase env
from dotenv import load_dotenv
load_dotenv(os.path.join(SCRIPT_DIR, "secret_supabase.env"))
SB_URL = os.environ.get("SUPABASE_URL", "")
SB_KEY = os.environ.get("SUPABASE_KEY", "")


def _http_patch(kode_tender: str, payload: dict) -> dict:
    import urllib.request
    url = f"{SB_URL}/rest/v1/draft_paket?kode_tender=eq.{kode_tender}"
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


def _http_get_snapshot(kode_tender: str) -> dict | None:
    import urllib.request
    url = f"{SB_URL}/rest/v1/draft_paket?kode_tender=eq.{kode_tender}&select=data_snapshot"
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
    except Exception as e:
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
    """Baca snapshot dari input, upsert ke Supabase."""
    inp = _read_input()
    kode_tender = inp.get("kode_tender", "")
    snapshot    = inp.get("snapshot", {})

    if not kode_tender:
        _write_output({"ok": False, "error": "kode_tender kosong"})
        return

    result = _http_patch(kode_tender, {"data_snapshot": snapshot})
    _write_output(result)


def cmd_load():
    """Ambil snapshot dari Supabase, tulis ke output."""
    inp = _read_input()
    kode_tender = inp.get("kode_tender", "")

    if not kode_tender:
        _write_output({"ok": False, "error": "kode_tender kosong"})
        return

    snapshot = _http_get_snapshot(kode_tender)
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
