"""
CONFIG — Single Source of Truth untuk konstanta proyek POKJA 2026
=================================================================
Semua script Python import dari sini.
VBA tetap pakai Private Const di ModWordLink.bas (harus sinkron manual).
"""
import os
from urllib.parse import quote

# ===== PATH =====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Template folder (relatif terhadap BASE_DIR → naik 1 level ke @ POKJA 2026)
POKJA_ROOT = os.path.dirname(os.path.dirname(BASE_DIR))  # @ POKJA 2026
TEMPLATE_DIR = os.path.join(POKJA_ROOT, "Paket Experiment")

# Python exe (portable WinPython)
PYTHON_EXE = os.path.join(BASE_DIR, "python", "python.exe")
PYTHONW_EXE = os.path.join(BASE_DIR, "python", "pythonw.exe")

# ===== EXCEL TEMPLATE =====
EXCEL_TEMPLATE = "@ BA PK 2026 (Improved) v.1.xlsm"

# ===== SHEET NAMES (mail merge target) =====
# PENTING: harus sinkron dengan VBA Private Const di ModWordLink.bas
SHEET_BA = "satu_data"
SHEET_REVIU = "list_reviu"
SHEET_DOKPIL = "list_dokpil"

# ===== WORD TEMPLATE → SHEET MAPPING =====
# (nama_file_word, sheet_name)
WORD_SHEET_MAP = [
    ("1. Full Dokumen BA PK v.1.docx", SHEET_BA),
    ("2. Isi Reviu PK v.1.docx",       SHEET_REVIU),
    ("3. Dokpil Full PK v.1.docx",      SHEET_DOKPIL),
]

# Keyword mapping: untuk detect sheet dari nama file Word (dipakai relink)
WORD_KEYWORD_MAP = {
    "BA PK":   SHEET_BA,
    "Reviu":   SHEET_REVIU,
    "Dokpil":  SHEET_DOKPIL,
}

# ===== VBA PDF MODES =====
# mode_name → (word_const, sheet_const, status_template)
PDF_MODES = {
    "pdf_bareviu":            ("WORD_BA",     "SHEET_BA",     "BA_REVIU_DPP_{kode}.pdf"),
    "pdf_all":                ("WORD_REVIU",  "SHEET_REVIU",  "Isi_Reviu_{kode}.pdf"),
    "pdf_dokpil":             ("WORD_DOKPIL", "SHEET_DOKPIL", "DOKPIL_{kode}.pdf"),
    "pdf":                    ("WORD_BA",     "SHEET_BA",     "Undangan_{kode}.pdf"),
    "pdf_pembuktian":         ("WORD_BA",     "SHEET_BA",     "BA Pembuktian & Nego_ {kode}"),
    "pdf_revaluasi":          ("WORD_BA",     "SHEET_BA",     "REvaluasi_{kode}.pdf"),
    "pdf_pembuktian_timpang": ("WORD_BA",     "SHEET_BA",     "BA Pembuktian Timpang_{kode}.pdf"),
}


# ===== HELPER: URL ENCODING =====
def excel_to_file_uri(excel_path):
    """Convert Windows path ke file:/// URI dengan encoding yang benar."""
    # Normalize path separators
    path = excel_path.replace('\\', '/')
    # Encode setiap segmen path, tapi preserve drive letter (D:)
    parts = path.split('/')
    encoded_parts = []
    for i, part in enumerate(parts):
        if i == 0 and len(part) == 2 and part[1] == ':':
            encoded_parts.append(part)  # drive letter apa adanya
        else:
            encoded_parts.append(quote(part, safe=''))
    return 'file:///' + '/'.join(encoded_parts)
