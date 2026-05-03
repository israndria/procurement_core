"""
sync_kalender.py — Cari tanggal pembukaan + pembuktian/penetapan dari GCal.
Dipanggil oleh VBA SyncKalender via WScript.Shell.
Output: _sync_kalender.json di script dir.
Usage: python sync_kalender.py <nama_tender>
"""
import sys
import os
import json
import re
from datetime import datetime, date

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")
OUT_PATH   = os.path.join(SCRIPT_DIR, "_sync_kalender.json")


def _build_service():
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


def cari_tanggal(nama_tender: str) -> dict:
    service = _build_service()

    # Ambil 3 kata pertama sebagai keyword
    kata = nama_tender.split()
    keyword = " ".join(kata[:3]) if len(kata) >= 3 else nama_tender

    resp = service.events().list(
        calendarId="primary",
        timeMin="2025-01-01T00:00:00Z",
        timeMax="2026-12-31T00:00:00Z",
        maxResults=200,
        singleEvents=True,
        orderBy="startTime",
        q=keyword,
    ).execute()

    events = resp.get("items", [])

    tgl_pembukaan  = None
    tgl_pembuktian = None

    for ev in events:
        summary = ev.get("summary", "").lower()
        dt_str = (ev.get("start", {}).get("date")
                  or ev.get("start", {}).get("dateTime", ""))[:10]
        if not dt_str:
            continue

        if tgl_pembukaan is None and "pembukaan" in summary:
            tgl_pembukaan = dt_str
        if tgl_pembuktian is None and any(k in summary for k in
                                          ("pembuktian", "penetapan", "klarifikasi")):
            tgl_pembuktian = dt_str

        if tgl_pembukaan and tgl_pembuktian:
            break

    return {
        "keyword":        keyword,
        "tgl_pembukaan":  tgl_pembukaan,
        "tgl_pembuktian": tgl_pembuktian,
        "error":          None,
    }


if __name__ == "__main__":
    # Baca nama dari file _sync_kalender_input.txt jika ada (hindari masalah special char di argv)
    inp_path = os.path.join(SCRIPT_DIR, "_sync_kalender_input.txt")
    if os.path.exists(inp_path):
        with open(inp_path, "r", encoding="utf-8-sig") as f:  # utf-8-sig strip BOM otomatis
            nama = f.read().strip()
        os.remove(inp_path)
    else:
        nama = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else ""

    if not nama:
        json.dump({"error": "nama_tender kosong"}, open(OUT_PATH, "w", encoding="utf-8"), ensure_ascii=False)
        sys.exit(1)

    try:
        hasil = cari_tanggal(nama)
    except Exception as e:
        hasil = {"error": str(e), "tgl_pembukaan": None, "tgl_pembuktian": None, "keyword": ""}

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(hasil, f, ensure_ascii=False, indent=2)

    print(json.dumps(hasil, ensure_ascii=False))
