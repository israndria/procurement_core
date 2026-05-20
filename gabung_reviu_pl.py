#!/usr/bin/env python3
"""
gabung_reviu_pl.py — Gabung BA Reviu (hal 1-2) + Isi Reviu + BA Reviu (hal 3)

Struktur output:
  Hal 1-2 : BA_REVIU_PL_{kode}.pdf hal 1-2   (Berita Acara Reviu)
  Hal 3-N : Isi_Reviu_DPP_{safe_name}.pdf     (semua halaman isi reviu)
  Hal N+1 : BA_REVIU_PL_{kode}.pdf hal 3      (halaman tanda tangan)

Usage:
    python gabung_reviu_pl.py "D:\\path\\ke\\folder\\paket"

Output: Gabung_Reviu_{kode_paket}.pdf di folder yang sama
"""
import sys
import os
import glob
from pypdf import PdfReader, PdfWriter


def gabung_reviu(folder: str) -> str:
    folder = os.path.abspath(folder)

    # Cari BA_REVIU_PL_*.pdf
    ba_files = glob.glob(os.path.join(folder, "BA_REVIU_PL_*.pdf"))
    if not ba_files:
        raise FileNotFoundError(f"BA_REVIU_PL_*.pdf tidak ditemukan di {folder}")
    ba_path = ba_files[0]
    ba_reader = PdfReader(ba_path)
    ba_pages = len(ba_reader.pages)
    if ba_pages < 2:
        raise ValueError(f"BA Reviu hanya {ba_pages} halaman, butuh minimal 2")

    # Cari Isi_Reviu_DPP_*.pdf
    isi_files = glob.glob(os.path.join(folder, "Isi_Reviu_DPP_*.pdf"))
    if not isi_files:
        raise FileNotFoundError(
            f"Isi_Reviu_DPP_*.pdf tidak ditemukan di {folder}. "
            "Cetak dulu via tombol 'Cetak Isi Reviu' di Excel."
        )
    isi_path = isi_files[0]
    isi_reader = PdfReader(isi_path)

    # Ambil kode_paket dari nama BA file: BA_REVIU_PL_{kode}.pdf
    ba_basename = os.path.basename(ba_path)
    kode_paket = ba_basename.replace("BA_REVIU_PL_", "").replace(".pdf", "")

    writer = PdfWriter()

    # 1. BA hal 1-2
    writer.add_page(ba_reader.pages[0])
    if ba_pages >= 2:
        writer.add_page(ba_reader.pages[1])

    # 2. Semua halaman Isi Reviu
    for page in isi_reader.pages:
        writer.add_page(page)

    # 3. BA hal 3 (tanda tangan) — skip jika BA hanya 2 hal
    if ba_pages >= 3:
        writer.add_page(ba_reader.pages[2])

    out_path = os.path.join(folder, f"Gabung_Reviu_{kode_paket}.pdf")
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
        print("Usage: python gabung_reviu_pl.py \"D:\\path\\ke\\folder\\paket\"")
        sys.exit(1)
    try:
        hasil = gabung_reviu(sys.argv[1])
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
