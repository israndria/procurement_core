#!/usr/bin/env python3
"""
gabung_reviu_pl.py — Gabung BA Reviu (scan) + Isi Reviu + BA Reviu (hal terakhir)

Struktur output:
  Hal 1-(N-1) : ba_path hal 1 s/d (total-1)   (Berita Acara Reviu)
  Hal N-M     : Isi_Reviu_DPP_*.pdf            (semua halaman isi reviu)
  Hal terakhir: ba_path hal terakhir            (halaman tanda tangan)

Usage:
    # Mode lama — cari BA_REVIU_PL_*.pdf di folder paket
    python gabung_reviu_pl.py "D:\\path\\ke\\folder\\paket"

    # Mode baru — BA dari file scan eksternal (nama bebas)
    python gabung_reviu_pl.py "D:\\path\\ke\\folder\\paket" --ba "D:\\Download\\REVIU KONSULTAN PERENCANAAN 1.pdf"

Output: Gabung_Reviu_{kode_unik_atau_kode}.pdf di folder paket
"""
import sys
import os
import glob
import re
from pypdf import PdfReader, PdfWriter


def gabung_reviu(folder: str, ba_path_override: str = None) -> str:
    folder = os.path.abspath(folder)

    # Tentukan file BA
    if ba_path_override:
        ba_path = os.path.abspath(ba_path_override)
        if not os.path.exists(ba_path):
            raise FileNotFoundError(f"File BA scan tidak ditemukan: {ba_path}")
    else:
        ba_files = glob.glob(os.path.join(folder, "BA_REVIU_PL_*.pdf"))
        if not ba_files:
            raise FileNotFoundError(
                f"BA_REVIU_PL_*.pdf tidak ditemukan di {folder}. "
                "Gunakan --ba untuk tentukan file scan."
            )
        ba_path = ba_files[0]

    ba_reader = PdfReader(ba_path)
    ba_pages = len(ba_reader.pages)
    if ba_pages < 2:
        raise ValueError(f"BA Reviu hanya {ba_pages} halaman, butuh minimal 2")

    # Cari Isi_Reviu_DPP_*.pdf di folder paket
    isi_files = glob.glob(os.path.join(folder, "Isi_Reviu_DPP_*.pdf"))
    if not isi_files:
        raise FileNotFoundError(
            f"Isi_Reviu_DPP_*.pdf tidak ditemukan di {folder}. "
            "Cetak dulu via tombol 'Cetak Isi Reviu' di Excel."
        )
    isi_path = isi_files[0]
    isi_reader = PdfReader(isi_path)

    # Tentukan nama output — ambil kode_unik dari nama Isi_Reviu atau dari nama folder
    isi_basename = os.path.basename(isi_path)
    m = re.match(r"Isi_Reviu_DPP_(.+)\.pdf", isi_basename)
    if m and m.group(1) != "000":
        kode_out = m.group(1)
    else:
        # Fallback: ambil dari nama BA file lama
        ba_basename = os.path.basename(ba_path)
        kode_out = ba_basename.replace("BA_REVIU_PL_", "").replace(".pdf", "")
        if not kode_out or kode_out == "000":
            kode_out = os.path.basename(folder)

    writer = PdfWriter()

    # 1. BA hal 1 s/d (total-1)
    for i in range(ba_pages - 1):
        writer.add_page(ba_reader.pages[i])

    # 2. Semua halaman Isi Reviu
    for page in isi_reader.pages:
        writer.add_page(page)

    # 3. BA hal terakhir (tanda tangan)
    writer.add_page(ba_reader.pages[ba_pages - 1])

    out_path = os.path.join(folder, f"Gabung_Reviu_{kode_out}.pdf")
    with open(out_path, "wb") as f:
        writer.write(f)

    print(f"OK: {out_path}")
    print(
        f"   BA Reviu ({ba_pages} hal) + Isi Reviu ({len(isi_reader.pages)} hal) "
        f"= {len(writer.pages)} hal total"
    )
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python gabung_reviu_pl.py \"D:\\path\\ke\\folder\\paket\" [--ba \"D:\\path\\ba_scan.pdf\"]")
        sys.exit(1)

    _folder = sys.argv[1]
    _ba_override = None
    if "--ba" in sys.argv:
        idx = sys.argv.index("--ba")
        if idx + 1 < len(sys.argv):
            _ba_override = sys.argv[idx + 1]

    try:
        hasil = gabung_reviu(_folder, _ba_override)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
