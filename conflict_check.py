"""
conflict_check.py — Cek status personil/alat lintas paket via Supabase + GCal.
Dipanggil VBA via WScript.Shell setelah MuatInputBA.

Usage: python conflict_check.py <kode_tender>
Output: _conflict_check.json di script dir

Format output:
{
  "personil": {
    "peserta_id_1": {
      "personil_1": {"nama": "X", "status": "warning|ok|unknown", "pesan": "..."},
      "personil_2": {...}
    },
    ...
  },
  "alat": { ... sama ... },
  "error": null
}
"""

import sys
import os
import json
from datetime import date, timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PATH   = os.path.join(SCRIPT_DIR, "_conflict_check.json")
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")

from dotenv import load_dotenv
load_dotenv(os.path.join(SCRIPT_DIR, "secret_supabase.env"))
SB_URL = os.environ.get("SUPABASE_URL", "")
SB_KEY = os.environ.get("SUPABASE_KEY", "")


# ── Supabase helpers ────────────────────────────────────────────────────────

def _sb_get(path: str) -> list:
    import urllib.request
    url = f"{SB_URL}/rest/v1/{path}"
    req = urllib.request.Request(url)
    req.add_header("apikey", SB_KEY)
    req.add_header("Authorization", f"Bearer {SB_KEY}")
    req.add_header("Accept", "application/json")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode("utf-8"))


def get_personil_target(kode_tender: str) -> list[dict]:
    """Ambil semua personil untuk paket target, sorted by peserta_id."""
    rows = _sb_get(
        f"paket_personil?kode_tender=eq.{kode_tender}"
        f"&select=peserta_id,nama_personil,posisi&order=peserta_id.asc"
    )
    return rows or []


def get_alat_target(kode_tender: str) -> list[dict]:
    rows = _sb_get(
        f"paket_alat?kode_tender=eq.{kode_tender}"
        f"&select=peserta_id,nama_alat,posisi&order=peserta_id.asc"
    )
    return rows or []


def get_konflik_personil(nama: str, kode_tender_target: str) -> list[dict]:
    """Cari nama personil di paket lain (bukan target)."""
    import urllib.parse
    nama_enc = urllib.parse.quote(nama)
    rows = _sb_get(
        f"paket_personil?nama_personil=eq.{nama_enc}"
        f"&kode_tender=neq.{kode_tender_target}"
        f"&select=kode_tender,nama_penyedia"
    )
    return rows or []


def get_konflik_alat(nama: str, kode_tender_target: str) -> list[dict]:
    import urllib.parse
    nama_enc = urllib.parse.quote(nama)
    rows = _sb_get(
        f"paket_alat?nama_alat=eq.{nama_enc}"
        f"&kode_tender=neq.{kode_tender_target}"
        f"&select=kode_tender,nama_penyedia"
    )
    return rows or []


def get_paket_info(kode_tender: str) -> dict:
    """Ambil nama_tender + data_snapshot (masa pelaksanaan) dari draft_paket."""
    rows = _sb_get(
        f"draft_paket?kode_tender=eq.{kode_tender}"
        f"&select=nama_tender,data_snapshot&limit=1"
    )
    if not rows:
        return {}
    row = rows[0]
    masa = None
    snap = row.get("data_snapshot") or {}
    # r12 = MD_E15 = Masa Pelaksanaan (hari) — row 12 di @ Master Data
    raw = snap.get("r12", "")
    if raw:
        try:
            masa = int(str(raw).split()[0])
        except Exception:
            pass
    return {
        "nama_tender": row.get("nama_tender", ""),
        "masa_pelaksanaan": masa,
    }


# ── GCal helper ─────────────────────────────────────────────────────────────

def _build_gcal_service():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from google.auth.transport.requests import Request

    with open(TOKEN_PATH, "r", encoding="utf-8") as f:
        token_data = json.load(f)

    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id"),
        client_secret=token_data.get("client_secret"),
        scopes=token_data.get("scopes", ["https://www.googleapis.com/auth/calendar"]),
    )
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)


_gcal_service = None
_sppbj_cache: dict[str, str | None] = {}  # kode_tender → "YYYY-MM-DD" | None


def get_tgl_sppbj(nama_tender: str) -> str | None:
    """Cari tanggal SPPBJ dari GCal. Return 'YYYY-MM-DD' atau None."""
    global _gcal_service
    if _gcal_service is None:
        try:
            _gcal_service = _build_gcal_service()
        except Exception:
            return None

    kata = nama_tender.split()
    keyword = " ".join(kata[:3]) if len(kata) >= 3 else nama_tender

    try:
        resp = _gcal_service.events().list(
            calendarId="primary",
            timeMin="2025-01-01T00:00:00Z",
            timeMax="2027-12-31T00:00:00Z",
            maxResults=50,
            singleEvents=True,
            orderBy="startTime",
            q=keyword,
        ).execute()
    except Exception:
        return None

    for ev in resp.get("items", []):
        summary = ev.get("summary", "").lower()
        if "penunjukan" in summary or "sppbj" in summary:
            dt_str = (ev.get("start", {}).get("date")
                      or ev.get("start", {}).get("dateTime", ""))[:10]
            if dt_str:
                return dt_str
    return None


# ── Status logic ─────────────────────────────────────────────────────────────

def cek_status(kode_konflik: str) -> dict:
    """
    Return {"kode_pokja_label": str, "status": "ok"|"warning"|"proses", "pesan": str}
    """
    info = get_paket_info(kode_konflik)
    nama = info.get("nama_tender", kode_konflik)
    masa = info.get("masa_pelaksanaan")  # hari, bisa None

    # Label singkat: ambil 3 kata pertama nama tender
    label = " ".join(nama.split()[:4]) if nama else kode_konflik

    # Coba ambil SPPBJ dari GCal (cached per nama)
    if kode_konflik not in _sppbj_cache:
        _sppbj_cache[kode_konflik] = get_tgl_sppbj(nama) if nama else None
    tgl_sppbj_str = _sppbj_cache[kode_konflik]

    if tgl_sppbj_str is None:
        # Belum SPPBJ — masih proses
        return {
            "label": label,
            "status": "proses",
            "pesan": f"Sedang proses di: {label}",
        }

    # SPPBJ ada — hitung masa kontrak
    try:
        tgl_sppbj = date.fromisoformat(tgl_sppbj_str)
    except Exception:
        return {"label": label, "status": "proses", "pesan": f"Sedang proses di: {label}"}

    if masa:
        tgl_selesai = tgl_sppbj + timedelta(days=masa)
        today = date.today()
        if today <= tgl_selesai:
            return {
                "label": label,
                "status": "warning",
                "pesan": f"Masih kontrak di: {label} (s.d. {tgl_selesai.strftime('%d/%m/%Y')})",
            }
        else:
            return {
                "label": label,
                "status": "ok",
                "pesan": f"Selesai kontrak di: {label} ({tgl_selesai.strftime('%d/%m/%Y')})",
            }
    else:
        # SPPBJ ada tapi masa pelaksanaan tidak diketahui
        return {
            "label": label,
            "status": "warning",
            "pesan": f"SPPBJ {tgl_sppbj_str} di: {label} (masa pelaksanaan ?)",
        }


# ── Main ─────────────────────────────────────────────────────────────────────

def run(kode_tender: str) -> dict:
    result = {"personil": {}, "alat": {}, "error": None}

    # Ambil semua personil & alat target
    personil_rows = get_personil_target(kode_tender)
    alat_rows     = get_alat_target(kode_tender)

    # Group by peserta_id
    from collections import defaultdict
    p_by_peserta: dict[str, list] = defaultdict(list)
    for r in personil_rows:
        p_by_peserta[r["peserta_id"]].append(r)

    a_by_peserta: dict[str, list] = defaultdict(list)
    for r in alat_rows:
        a_by_peserta[r["peserta_id"]].append(r)

    # Cek konflik personil
    for pid, items in p_by_peserta.items():
        result["personil"][pid] = {}
        for i, item in enumerate(items):
            nama = item["nama_personil"]
            posisi = item.get("posisi", f"personil_{i+1}")
            key = f"personil_{i+1}"

            konflik = get_konflik_personil(nama, kode_tender)
            if not konflik:
                result["personil"][pid][key] = {
                    "nama": nama, "posisi": posisi,
                    "status": "ok", "pesan": "Tidak ada riwayat",
                }
            else:
                # Ambil status dari paket konflik pertama (paling relevan)
                kode_k = konflik[0]["kode_tender"]
                st = cek_status(kode_k)
                result["personil"][pid][key] = {
                    "nama": nama, "posisi": posisi,
                    "status": st["status"],
                    "pesan": st["pesan"],
                }

    # Cek konflik alat
    for pid, items in a_by_peserta.items():
        result["alat"][pid] = {}
        for i, item in enumerate(items):
            nama = item["nama_alat"]
            posisi = item.get("posisi", f"alat_{i+1}")
            key = f"alat_{i+1}"

            konflik = get_konflik_alat(nama, kode_tender)
            if not konflik:
                result["alat"][pid][key] = {
                    "nama": nama, "posisi": posisi,
                    "status": "ok", "pesan": "Tidak ada riwayat",
                }
            else:
                kode_k = konflik[0]["kode_tender"]
                st = cek_status(kode_k)
                result["alat"][pid][key] = {
                    "nama": nama, "posisi": posisi,
                    "status": st["status"],
                    "pesan": st["pesan"],
                }

    return result


if __name__ == "__main__":
    kode = sys.argv[1] if len(sys.argv) > 1 else ""
    if not kode:
        out = {"error": "kode_tender kosong", "personil": {}, "alat": {}}
    else:
        try:
            out = run(kode)
        except Exception as e:
            out = {"error": str(e), "personil": {}, "alat": {}}

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(json.dumps(out, ensure_ascii=False))
