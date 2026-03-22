"""
RELINK TEMPLATES - Update mail merge data source di Word templates
=================================================================
Scan folder Excel, cari semua .docx (bukan Merged), re-link ke file .xlsm
yang ada di folder yang sama.

Dipanggil dari tombol VBA "Relink Template" di Excel.

Cara pakai:
  python relink_templates.py "D:\\path\\to\\file.xlsm"
"""
import os
import sys
import ctypes

# Import fungsi link_word_to_excel dari setup_paket_baru
from setup_paket_baru import link_word_to_excel

# Mapping: keyword di nama file Word -> sheet name untuk mail merge
WORD_SHEET_MAP = {
    "BA PK":   "satu_data",
    "Reviu":   "list_reviu",
    "Dokpil":  "list_dokpil",
}


def detect_sheet(docx_name):
    """Deteksi sheet name berdasarkan nama file Word."""
    for keyword, sheet in WORD_SHEET_MAP.items():
        if keyword.lower() in docx_name.lower():
            return sheet
    return None


def show_message(msg, title="Relink Template", icon=0x40):
    """Tampilkan popup Windows."""
    ctypes.windll.user32.MessageBoxW(0, msg, title, icon)


def main():
    if len(sys.argv) < 2:
        show_message("Usage: relink_templates.py <excel_path>", icon=0x10)
        sys.exit(1)

    excel_path = os.path.abspath(sys.argv[1])
    folder = os.path.dirname(excel_path)

    if not os.path.exists(excel_path):
        show_message(f"File Excel tidak ditemukan:\n{excel_path}", icon=0x10)
        sys.exit(1)

    # Scan folder untuk .docx (skip Merged dan .bak)
    docx_files = [
        f for f in os.listdir(folder)
        if f.lower().endswith(".docx")
        and "(Merged)" not in f
        and not f.startswith("~")
    ]

    if not docx_files:
        show_message(f"Tidak ada file .docx ditemukan di:\n{folder}", icon=0x30)
        sys.exit(0)

    results = []
    for docx in docx_files:
        sheet = detect_sheet(docx)
        if sheet is None:
            results.append(f"[SKIP] {docx} (tidak dikenali)")
            continue

        docx_path = os.path.join(folder, docx)
        ok = link_word_to_excel(docx_path, excel_path, sheet)
        if ok:
            results.append(f"[OK] {docx} -> {sheet}")
        else:
            results.append(f"[GAGAL] {docx}")

    summary = "Relink selesai!\n\n" + "\n".join(results)
    show_message(summary)


if __name__ == "__main__":
    main()
