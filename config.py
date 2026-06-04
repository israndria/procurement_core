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
EXCEL_TEMPLATE = "0. BAPK - Template.xlsm"

# ===== SHEET NAMES (mail merge target) =====
# PENTING: harus sinkron dengan VBA Private Const di ModWordLink.bas
SHEET_BA = "satu_data"
SHEET_REVIU = "list_reviu"
SHEET_DOKPIL = "list_dokpil"

# ===== WORD TEMPLATE → SHEET MAPPING =====
# (nama_file_word, sheet_name)
# Template BA tender DIPECAH per-dokumen (dulu monolitik "1. Full Dokumen BAPK").
# File monolitik lama diarsipkan di Paket Experiment\backup\.
WORD_SHEET_MAP = [
    ("1. Reviu Dok. Persiapan Pengadaan - Template.docx", SHEET_BA),
    ("2. Isi Reviu PK - Template.docm",                   SHEET_REVIU),
    ("3. Dokpil Full PK - Template.docx",                 SHEET_DOKPIL),
    ("4. Undangan Full PK - Template.docx",               SHEET_BA),
    ("5. Berita Acara Utama PK - Template.docx",          SHEET_BA),
    ("6. Ringkasan Evaluasi PK - Template.docx",          SHEET_BA),
    ("7. BA Dengan Timpang PK - Template.docx",           SHEET_BA),
]

# Keyword mapping: untuk detect sheet dari nama file Word (dipakai relink).
# Urutan penting: keyword spesifik dulu (Isi Reviu pakai list_reviu, selain itu satu_data).
WORD_KEYWORD_MAP = {
    "Isi Reviu": SHEET_REVIU,
    "Dokpil":    SHEET_DOKPIL,
    "Reviu Dok": SHEET_BA,      # BA Reviu DPP — sumber satu_data
    "Undangan":  SHEET_BA,
    "Berita Acara Utama": SHEET_BA,
    "Ringkasan Evaluasi": SHEET_BA,
    "Timpang":   SHEET_BA,
    "BAPK":      SHEET_BA,
}

# ===== PL (Pengadaan Langsung) TEMPLATE =====
TEMPLATE_DIR_PL = os.path.join(
    POKJA_ROOT,
    "Paket Experiment - Pengadaan Langsung",
    "Development - PL - JKK",
)
TEMPLATE_DIR_PL_PK = os.path.join(
    POKJA_ROOT,
    "Paket Experiment - Pengadaan Langsung",
    "Development - PL - PK",
)
EXCEL_TEMPLATE_PL = "0. BAPLJKK - Template.xlsm"
SHEET_BA_PL    = "satu_data"
SHEET_REVIU_PL = "list_reviu"
SHEET_DOKPIL_PL = "list_dokpil"

WORD_SHEET_MAP_PL = [
    ("5. BA PLJKK - Template.docx",               SHEET_BA_PL),
    ("1. Reviu DPP PLJKK - Template.docx",        SHEET_REVIU_PL),
    ("3. Dokpil Full PLJKK - Template.docx",      SHEET_DOKPIL_PL),
]

# Output folder per jenis PL (folder tujuan buat folder baru)
OUTPUT_DIR_PL_JKK = os.path.join(
    POKJA_ROOT,
    "@ Pejabat Pengadaan 2026",
    "@ Pengadaan Langsung JKK",
)
OUTPUT_DIR_PL_PK = os.path.join(
    POKJA_ROOT,
    "@ Pejabat Pengadaan 2026",
    "@ Pengadaan Langsung PK",
)

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
