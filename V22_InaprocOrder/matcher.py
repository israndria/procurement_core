"""
matcher.py — Cocokkan nama barang dari Excel ke variasi yang tersedia di halaman produk.
"""

import re


def cocokkan_variasi(nama_barang: str, variasi_tersedia: list[str]) -> str | None:
    """
    Cari variasi yang paling cocok dengan nama_barang dari Excel.

    Contoh:
      "Tinta Printer Spesifikasi : Hitam", ["Hitam", "Cyan", "Magenta", "Yellow"]
      → "Hitam"

      "Amplop Coklat Besar", []
      → None (produk tanpa variasi — langsung isi kuantitas)

      "Tinta Printer Spesifikasi : Warna", ["Hitam", "Cyan", "Magenta", "Yellow"]
      → None (tidak cocok — flag ⚠️)
    """
    if not variasi_tersedia:
        return None  # tidak ada variasi → langsung isi qty

    nama_lower = nama_barang.lower()

    # Pass 1: exact substring match
    for variasi in variasi_tersedia:
        if variasi.lower() in nama_lower:
            return variasi

    # Pass 2: token-based match (tiap kata di nama cocok dengan kata di variasi)
    kata_nama = set(re.findall(r'\w+', nama_lower))
    for variasi in variasi_tersedia:
        kata_variasi = set(re.findall(r'\w+', variasi.lower()))
        if kata_variasi & kata_nama:  # ada irisan
            return variasi

    return None  # tidak cocok


def ada_variasi(variasi_tersedia: list[str]) -> bool:
    """Return True jika produk punya pilihan variasi."""
    return len(variasi_tersedia) > 0
