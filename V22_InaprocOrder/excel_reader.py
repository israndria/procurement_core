"""
excel_reader.py — Baca file Excel pesanan dari PPK.
Mengekstrak nama barang, kuantitas, dan hyperlink ke katalog.inaproc.id.
"""

import openpyxl


def baca_pesanan(file_path: str) -> list[dict]:
    """
    Baca file Excel PPK dan return list item pesanan.

    Setiap item: {
        'nama_barang': str,
        'kuantitas': int,
        'link': str,
        'baris': int  (nomor baris di Excel, untuk debugging)
    }

    Baris yang tidak punya hyperlink di kolom C dilewati otomatis
    (header, baris kosong, subtotal, dsb).
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    pesanan = []

    for row in ws.iter_rows():
        # Kolom C = index 2 (0-based)
        if len(row) < 3:
            continue

        sel_link = row[2]  # Kolom C — berisi hyperlink
        if sel_link.hyperlink is None:
            continue

        link = sel_link.hyperlink.target
        if not link or "katalog.inaproc.id" not in link:
            continue

        # Kolom B = nama barang
        nama_barang = str(row[1].value).strip() if row[1].value else ""

        # Kolom D = kuantitas
        kuantitas = 1
        if len(row) >= 4 and row[3].value is not None:
            try:
                kuantitas = int(row[3].value)
            except (ValueError, TypeError):
                kuantitas = 1

        if not nama_barang:
            continue

        pesanan.append({
            "nama_barang": nama_barang,
            "kuantitas": kuantitas,
            "link": link,
            "baris": sel_link.row,
        })

    wb.close()
    return pesanan
