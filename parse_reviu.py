"""
parse_reviu.py — Parse Draft_Pokja_XXX.pdf → isi database_reviu + database_dokpil
====================================================================================
Dipanggil dari VBA ModDraftPaket via WScript.Shell:
  python parse_reviu.py "<path_pdf>" "<kode_pokja>" "<folder_output>"

Output: <folder_output>/_parse_reviu.json
"""

import sys
import os
import json
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─── Load env Supabase ───────────────────────────────────────────────────────
def _load_env():
    env_file = os.path.join(BASE_DIR, "secret_supabase.env")
    if os.path.exists(env_file):
        with open(env_file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

_load_env()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

# ─── Fetch cara_pembayaran dari Supabase ─────────────────────────────────────
def fetch_cara_pembayaran():
    try:
        import httpx
        r = httpx.get(
            f"{SUPABASE_URL}/rest/v1/cara_pembayaran?select=id,keyword,teks",
            headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
            timeout=10
        )
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    # Fallback hardcode jika Supabase tidak tersedia
    return [
        {
            "id": "monthly_certificate",
            "keyword": ["monthly", "monthly certificate", "bulanan"],
            "teks": "(Monthly Certificate) Pembayaran dilakukan dengan cara didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara bulanan"
        },
        {
            "id": "termin",
            "keyword": ["termin", "pembayaran termin", "angsuran"],
            "teks": "(Termin) Pembayaran dilakukan dengan cara pembayarannya didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara cara angsuran"
        }
    ]

def deteksi_cara_pembayaran(daftar_cp, bidang=""):
    """
    Deteksi cara pembayaran berdasarkan bidang paket (dari Supabase).
    - Bina Marga (jalan/jembatan) → Monthly Certificate (MC)
    - Semua bidang lain (SDA, CK, irigasi, bangunan, drainase,
      pendidikan, perdagangan, dll) → Termin
    """
    bidang_lower = (bidang or "").lower().strip()

    # Kata kunci yang mengindikasikan Bina Marga / pekerjaan jalan
    KEYWORD_BINAMARGA = [
        "bina marga", "jalan", "jembatan", "perkerasan", "rigid", "flexible",
        "hotmix", "aspal", "overlay"
    ]
    is_binamarga = any(kw in bidang_lower for kw in KEYWORD_BINAMARGA)
    id_pilih = "monthly_certificate" if is_binamarga else "termin"

    for cp in daftar_cp:
        if cp["id"] == id_pilih:
            return cp["id"], cp["teks"]

    return daftar_cp[0]["id"], daftar_cp[0]["teks"]

# ─── Helpers teks ────────────────────────────────────────────────────────────
def bersihkan(s):
    return " ".join(s.split()).strip()

def normalisasi_pengalaman(s):
    m = re.search(r"(\d+)", s or "")
    return int(m.group(1)) if m else 0

# ─── Parser utama PDF ────────────────────────────────────────────────────────
def parse_pdf(path_pdf, bidang=""):
    try:
        import pdfplumber
    except ImportError:
        return None, "pdfplumber tidak tersedia"

    hasil = {
        # database_reviu
        "reviu": {
            "E2":  {"nilai": "", "status": "kosong", "label": "Fungsi Bangunan"},
            "E6":  {"nilai": "", "status": "kosong", "label": "SBU KBLI 2020"},
            "E7":  {"nilai": "", "status": "kosong", "label": "SBU KBLI 2015"},
            "E9":  {"nilai": "", "status": "kosong", "label": "Alat 1"},
            "E10": {"nilai": "", "status": "kosong", "label": "Alat 2"},
            "E11": {"nilai": "", "status": "kosong", "label": "Alat 3"},
            "E12": {"nilai": "", "status": "kosong", "label": "Alat 4"},
            "E13": {"nilai": "", "status": "kosong", "label": "Alat 5"},
            "E14": {"nilai": "", "status": "kosong", "label": "Alat 6"},
            "E15": {"nilai": "", "status": "kosong", "label": "Kapasitas 1"},
            "E16": {"nilai": "", "status": "kosong", "label": "Kapasitas 2"},
            "E17": {"nilai": "", "status": "kosong", "label": "Kapasitas 3"},
            "E18": {"nilai": "", "status": "kosong", "label": "Kapasitas 4"},
            "E19": {"nilai": "", "status": "kosong", "label": "Kapasitas 5"},
            "E20": {"nilai": "", "status": "kosong", "label": "Kapasitas 6"},
            "E21": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 1"},
            "E22": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 2"},
            "E23": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 3"},
            "E24": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 4"},
            "E25": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 5"},
            "E26": {"nilai": "", "status": "kosong", "label": "Jumlah Alat 6"},
            "E27": {"nilai": "", "status": "kosong", "label": "Jabatan Teknis"},
            "E28": {"nilai": 0,  "status": "kosong", "label": "Pengalaman Teknis (Tahun)"},
            "E29": {"nilai": "", "status": "kosong", "label": "Nama SKK/SKT"},
            "E30": {"nilai": "", "status": "kosong", "label": "Jabatan K3"},
            "E31": {"nilai": 0,  "status": "kosong", "label": "Pengalaman K3 (Tahun)"},
            "E32": {"nilai": "", "status": "kosong", "label": "Sertifikat K3"},
            "E33": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan RK3"},
            "E34": {"nilai": "", "status": "kosong", "label": "Risiko Tertinggi RK3"},
        },
        # database_dokpil
        "dokpil": {
            "E6":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 1"},
            "E7":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 2"},
            "E8":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 3"},
            "E9":  {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 4"},
            "E10": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 5"},
            "E11": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 6"},
            "E12": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 7"},
            "E13": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 8"},
            "E14": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 9"},
            "E15": {"nilai": "", "status": "kosong", "label": "Uraian Pekerjaan 10"},
            "E16": {"nilai": "", "status": "kosong", "label": "Cara Pembayaran", "id_cp": ""},
        },
        "error": ""
    }

    def set_val(sheet, cell, nilai, status="terisi"):
        hasil[sheet][cell]["nilai"] = nilai
        hasil[sheet][cell]["status"] = status

    with pdfplumber.open(path_pdf) as pdf:
        semua_halaman = [p.extract_text() or "" for p in pdf.pages]
        semua_teks = "\n".join(semua_halaman)

        # ── E2: Fungsi Bangunan (ambil dari Tujuan / Maksud) ─────────────────
        for teks in semua_halaman:
            m = re.search(
                r"(?:2\.2\s+Tujuan|Tujuan\s*\n+)(.*?)(?=\n\s*2\.\d|\n\s*\d+\.|\Z)",
                teks, re.DOTALL | re.IGNORECASE
            )
            if m:
                tujuan = bersihkan(m.group(1))
                if len(tujuan) > 20:
                    set_val("reviu", "E2", tujuan)
                    break
        # Fallback: seksi Maksud
        if not hasil["reviu"]["E2"]["nilai"]:
            for teks in semua_halaman:
                m = re.search(
                    r"(?:2\.1\s+Maksud|Maksud\s*\n+)(.*?)(?=\n\s*2\.\d|\n\s*\d+\.|\Z)",
                    teks, re.DOTALL | re.IGNORECASE
                )
                if m:
                    maksud = bersihkan(m.group(1))
                    if len(maksud) > 20:
                        set_val("reviu", "E2", maksud)
                        break

        # ── E6/E7: SBU — cari di PDF, fallback "perlu_keputusan" ─────────────
        # Pola umum di KAK: "SBU : BS001" atau "Kode SBU BS 001" atau "BS001"
        sbu_found = []
        for teks in semua_halaman:
            # Cari pola kode SBU: kombinasi 2 huruf + 3 digit (misal BS001, SI001, SP002)
            hits = re.findall(
                r'\bSBU\b[^A-Z0-9]*([A-Z]{2}\d{3})\b'   # "SBU BS001"
                r'|\bKode\s+SBU[^A-Z0-9]*([A-Z]{2}\d{3})\b'  # "Kode SBU BS001"
                r'|\b([A-Z]{2}\d{3})\b(?=[^)]*(?:SBU|Subkualifikasi|Bangunan Sipil|Bangunan Gedung))',
                teks
            )
            for h in hits:
                kode = (h[0] or h[1] or h[2]).strip()
                if kode and kode not in sbu_found:
                    sbu_found.append(kode)
            if sbu_found:
                break

        if len(sbu_found) >= 2:
            set_val("reviu", "E6", sbu_found[0])   # KBLI 2020
            set_val("reviu", "E7", sbu_found[1])   # KBLI 2015
        elif len(sbu_found) == 1:
            set_val("reviu", "E6", sbu_found[0])
            set_val("reviu", "E7", "", "perlu_keputusan")
        else:
            # Tidak ditemukan di PDF — perlu diisi manual
            set_val("reviu", "E6", "", "perlu_keputusan")
            set_val("reviu", "E7", "", "perlu_keputusan")

        # ── E9-E26: Tabel Peralatan ───────────────────────────────────────────
        alat_list = []
        for teks in semua_halaman:
            if "Peralatan" in teks and ("Kapasitas" in teks or "Kode" in teks or "unit" in teks.lower()):
                # Cari baris tabel peralatan: "1. NamaAlat KodeAlat Kapasitas Jumlah"
                baris_alat = re.findall(
                    r"(\d+)[.\)]\s+([A-Za-z][A-Za-z0-9 \-/]+?)\s+(E\d+|[A-Z]\d+)?\s*(\d+[-–]\d+\s*[Tt]on|\d+\s*[Tt]on|\d+\s*[Kk][Ww]|\d+\s*[Mm]3/[Jj]am|\d+\s*[A-Za-z]+)?\s*(\d+)\s*(?:[Uu]nit|unit)?",
                    teks
                )
                if baris_alat:
                    for b in baris_alat:
                        no, nama, kode, kapasitas, jumlah = b
                        alat_list.append({
                            "nama": bersihkan(nama),
                            "kapasitas": bersihkan(kapasitas) if kapasitas else "",
                            "jumlah": jumlah.strip() + " Unit" if jumlah else "1 Unit"
                        })
                    break

        # Fallback: cari baris lebih sederhana
        if not alat_list:
            for teks in semua_halaman:
                if "Tandem" in teks or "Vibratory" in teks or "Excavator" in teks or "Dump Truck" in teks:
                    baris = re.findall(
                        r"(\d+)[.\)]\s+(.+?)\s+(\d+[-–]\d+\s*[Tt]on|\d+\s*[Tt]on|\d+\s*[A-Za-z]+)?\s+(\d+)\s*$",
                        teks, re.MULTILINE
                    )
                    for b in baris:
                        no, nama, kapasitas, jumlah = b
                        nama_bersih = bersihkan(nama)
                        # Buang kode alat (E17, E19, dll) dari nama
                        nama_bersih = re.sub(r'\b[A-Z]\d{2,3}\b', '', nama_bersih).strip()
                        alat_list.append({
                            "nama": nama_bersih,
                            "kapasitas": bersihkan(kapasitas) if kapasitas else "",
                            "jumlah": jumlah.strip() + " Unit"
                        })
                    if alat_list:
                        break

        ALAT_CELLS  = ["E9","E10","E11","E12","E13","E14"]
        KAP_CELLS   = ["E15","E16","E17","E18","E19","E20"]
        JML_CELLS   = ["E21","E22","E23","E24","E25","E26"]

        for i, alat in enumerate(alat_list[:6]):
            set_val("reviu", ALAT_CELLS[i], alat["nama"])
            if alat["kapasitas"]:
                set_val("reviu", KAP_CELLS[i], alat["kapasitas"])
            set_val("reviu", JML_CELLS[i], alat["jumlah"])

        # Sel alat yang tidak terisi → status "tidak_ada"
        for i in range(len(alat_list), 6):
            hasil["reviu"][ALAT_CELLS[i]]["status"] = "tidak_ada"
            hasil["reviu"][KAP_CELLS[i]]["status"] = "tidak_ada"
            hasil["reviu"][JML_CELLS[i]]["status"] = "tidak_ada"

        # ── E27-E32: Tabel Personil ───────────────────────────────────────────
        for teks in semua_halaman:
            if "Personil" in teks and ("Jabatan" in teks or "Pelaksana" in teks or "K3" in teks):
                # Jabatan teknis (bukan K3)
                m_teknis = re.search(
                    r"(\d+)\s+(Pelaksana\s+[^\n]+?(?:Jalan|Gedung|Saluran|Bangunan|Lapangan)[^\n]*?)\s+"
                    r"(\d+)\s*[Tt]ahun\s+(SKK[^\n]+|SKT[^\n]+)",
                    teks, re.IGNORECASE
                )
                if m_teknis:
                    set_val("reviu", "E27", bersihkan(m_teknis.group(2)))
                    set_val("reviu", "E28", normalisasi_pengalaman(m_teknis.group(3)))
                    set_val("reviu", "E29", bersihkan(m_teknis.group(4)))
                else:
                    # Fallback lebih longgar
                    m_teknis2 = re.search(
                        r"1\s+(Pelaksana[^\n]+)\s+(\d+)\s*[Tt]ahun\s+(S\s*K\s*[KT][^\n]+)",
                        teks, re.IGNORECASE
                    )
                    if m_teknis2:
                        jabatan = bersihkan(m_teknis2.group(1))
                        set_val("reviu", "E27", jabatan)
                        set_val("reviu", "E28", normalisasi_pengalaman(m_teknis2.group(2)))
                        # Normalisasi SKK/SKT: hapus spasi berlebih di dalam kata
                        skk_raw = bersihkan(m_teknis2.group(3))
                        skk_norm = re.sub(r'\bS\s*K\s*K\b', 'SKK', skk_raw, flags=re.IGNORECASE)
                        skk_norm = re.sub(r'\bS\s*K\s*T\b', 'SKT', skk_norm, flags=re.IGNORECASE)
                        set_val("reviu", "E29", skk_norm)

                # Jabatan K3
                m_k3 = re.search(
                    r"(\d+)\s+(Petugas\s+K3|Ahli\s+K3[^\n]*)\s+(?:(\d+)\s*[Tt]ahun\s+)?(Sertifikat[^\n]+|SKA[^\n]+|SKK[^\n]+)",
                    teks, re.IGNORECASE
                )
                if m_k3:
                    set_val("reviu", "E30", bersihkan(m_k3.group(2)))
                    set_val("reviu", "E31", normalisasi_pengalaman(m_k3.group(3) or "0"))
                    set_val("reviu", "E32", bersihkan(m_k3.group(4)))
                else:
                    # Fallback: cari baris "Petugas K3" + baris berikutnya sebagai sertifikat
                    m_k3b = re.search(r"(Petugas\s+K3|Ahli\s+K3)\s+(Sertifikat[^\n]+)", teks, re.IGNORECASE)
                    if m_k3b:
                        set_val("reviu", "E30", bersihkan(m_k3b.group(1)))
                        set_val("reviu", "E32", bersihkan(m_k3b.group(2)))
                        set_val("reviu", "E31", 0)

                if hasil["reviu"]["E27"]["nilai"] or hasil["reviu"]["E30"]["nilai"]:
                    break

        # ── E33/E34: RK3K ─────────────────────────────────────────────────────
        for teks in semua_halaman:
            if "RK3" in teks or "Keselamatan" in teks and "Uraian Pekerjaan" in teks:
                # Kumpulkan semua uraian pekerjaan dari tabel RK3K
                uraian_list = re.findall(
                    r"^\s*\d+[.\)]\s+([A-Z][^\n]{10,}?)(?=\s*[-–]|\s*$)",
                    teks, re.MULTILINE
                )
                if uraian_list:
                    uraian_gabung = "\n".join(u.strip() for u in uraian_list[:4])
                    set_val("reviu", "E33", uraian_gabung)

                # Risiko tertinggi: cari blok yang mengandung "Paling Tinggi"
                # PDF struktur: "- Bahaya  Resiko\nPaling\nTinggi" atau inline
                m_risiko = re.search(
                    r"([-–]\s*[A-Z][^\n]{3,80}?)\s+Resiko[\s\n]*Paling[\s\n]*Tinggi",
                    teks, re.IGNORECASE
                )
                if m_risiko:
                    val = re.sub(r'\s+Resiko\s*$', '', bersihkan(m_risiko.group(1)), flags=re.IGNORECASE)
                    set_val("reviu", "E34", val)
                else:
                    # Fallback: cari "Paling Tinggi" sebagai kata terpisah di baris manapun
                    # lalu ambil teks identifikasi bahaya di baris/kolom sebelumnya
                    idx = teks.lower().find("paling\ntinggi")
                    if idx == -1:
                        idx = teks.lower().find("paling tinggi")
                    if idx > 0:
                        potongan = teks[max(0, idx-200):idx]
                        bahaya = re.findall(r"[-–]\s*([A-Z][^\n]{3,80})", potongan)
                        if bahaya:
                            set_val("reviu", "E34", "- " + bersihkan(bahaya[-1]))

                if hasil["reviu"]["E33"]["nilai"] or hasil["reviu"]["E34"]["nilai"]:
                    break

        # ── database_dokpil E6-E15: Uraian Pekerjaan dari Divisi RAB ─────────
        divisi_list = []
        for teks in semua_halaman:
            if "DIVISI" in teks and ("UMUM" in teks or "PERKERASAN" in teks or "REKAPITULASI" in teks):
                # Cari pola: "DIVISI 1 UMUM", "DIVISI 4 PEKERJAAN PREVENTIF", dll
                divisi_raw = re.findall(
                    r"DIVISI\s+\d+[.\s]*([A-Z][A-Z\s/,()]+?)(?=\s*\.\s*\.|\s+DIVISI\s+\d|\s+[A-Z]\s+Jumlah|Jumlah\s+Harga|\n\n|$)",
                    teks
                )
                for d in divisi_raw:
                    d_bersih = bersihkan(d).rstrip(". ").lower()
                    if d_bersih and d_bersih not in divisi_list and len(d_bersih) > 2:
                        divisi_list.append(d_bersih)
                if divisi_list:
                    break

        DOKPIL_CELLS = ["E6","E7","E8","E9","E10","E11","E12","E13","E14","E15"]
        for i, div in enumerate(divisi_list[:10]):
            set_val("dokpil", DOKPIL_CELLS[i], div)
        for i in range(len(divisi_list), 10):
            hasil["dokpil"][DOKPIL_CELLS[i]]["status"] = "tidak_ada"

        # ── database_dokpil E16: Cara Pembayaran ─────────────────────────────
        daftar_cp = fetch_cara_pembayaran()
        id_cp, teks_cp = deteksi_cara_pembayaran(daftar_cp, bidang)
        hasil["dokpil"]["E16"]["nilai"] = teks_cp
        hasil["dokpil"]["E16"]["id_cp"] = id_cp
        hasil["dokpil"]["E16"]["status"] = "terisi"

    return hasil, ""


# ─── Main ────────────────────────────────────────────────────────────────────
def main():
    # Mode --argfile: baca path_pdf, folder_output, bidang dari file teks
    bidang = ""
    if len(sys.argv) >= 3 and sys.argv[1] == "--argfile":
        argfile = sys.argv[2]
        with open(argfile, encoding="utf-8") as f:
            lines = [l.rstrip("\n").rstrip("\r") for l in f.readlines()]
        path_pdf = lines[0] if len(lines) > 0 else ""
        folder_output = lines[1] if len(lines) > 1 else ""
        bidang = lines[2] if len(lines) > 2 else ""
    elif len(sys.argv) >= 3:
        path_pdf = sys.argv[1]
        folder_output = sys.argv[2]
        bidang = sys.argv[3] if len(sys.argv) > 3 else ""
    else:
        print("Usage: python parse_reviu.py <path_pdf> <folder_output> [bidang]")
        sys.exit(1)

    if not os.path.exists(path_pdf):
        out = {"error": f"File PDF tidak ditemukan: {path_pdf}", "reviu": {}, "dokpil": {}}
        out_path = os.path.join(folder_output, "_parse_reviu.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        sys.exit(0)

    hasil, err = parse_pdf(path_pdf, bidang)
    if err:
        hasil = {"error": err, "reviu": {}, "dokpil": {}}

    out_path = os.path.join(folder_output, "_parse_reviu.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(hasil, f, ensure_ascii=False, indent=2)

    print(f"OK: {out_path}")


if __name__ == "__main__":
    main()
