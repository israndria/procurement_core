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
def _is_pokja_root(path):
    """True jika path terlihat sebagai root dokumen POKJA yang tersync Drive."""
    return bool(path) and any(
        os.path.isdir(os.path.join(path, marker))
        for marker in ("@ Tender 2026", "@ Pejabat Pengadaan 2026", "memory")
    )


def _discover_pokja_root():
    """Cari root dokumen per-PC; source code tidak lagi berada di Google Drive."""
    configured = os.environ.get("POKJA_DRIVE_ROOT", "").strip().strip('"')
    candidates = [
        configured,
        r"D:\Dokumen\@ POKJA 2026",
        r"G:\Other computers\My Laptop\@ POKJA 2026",
        r"C:\POKJA2026",
        os.path.dirname(BASE_DIR),
    ]
    for candidate in candidates:
        if _is_pokja_root(candidate):
            return os.path.normpath(candidate)
    raise RuntimeError(
        "Root dokumen POKJA tidak ditemukan. Set POKJA_DRIVE_ROOT ke folder "
        "'@ POKJA 2026' di Google Drive."
    )


POKJA_ROOT = _discover_pokja_root()
TEMPLATE_DIR = os.path.join(POKJA_ROOT, "Paket Experiment")

# Python exe (portable WinPython)
PYTHON_EXE = os.path.normpath(
    os.environ.get("POKJA_PYTHON", os.path.join(BASE_DIR, "python", "python.exe"))
)
PYTHONW_EXE = os.path.join(os.path.dirname(PYTHON_EXE), "pythonw.exe")

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
    ("8. Berita Acara Minimalis - Template.docx",          SHEET_BA),
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
    "Berita Acara Minimalis": SHEET_BA,
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
    ("1. BA Reviu DPP PLJKK - Template.docx",     SHEET_REVIU_PL),
    ("2. Isi Reviu PLJKK - Template.docm",        SHEET_REVIU_PL),
    ("3. Dokpil Full PLJKK - Template.docx",      SHEET_DOKPIL_PL),
    ("4. Undangan Full PLJK - Template.docx",     SHEET_BA_PL),
    ("5. BA PLJKK - Template.docx",               SHEET_BA_PL),
]

# ===== PL WORKFLOW V2 =====
# Donor lama tetap dipertahankan. Registry ini hanya menjadi kontrak untuk
# paket baru: satu workbook per subjenis, isi Word per domain, header dipilih
# saat output dibuat. Folder donor boleh diganti user tanpa mengubah kode.
PL_WORKFLOW_REGISTRY = {
    "PL_PERENCANAAN": {
        "jenis_pl": "JKK",
        "label": "PL Konsultansi Perencanaan",
        "folder_name": "Perencanaan",
        "excel_template": "0. BAPLJKK - Template Perencanaan.xlsm",
        "word_map": [
            ("1. BA Reviu DPP PLJKK - Template Perencanaan.docx", SHEET_REVIU_PL),
            ("2. Isi Reviu PLJKK - Template Perencanaan.docm", SHEET_REVIU_PL),
            ("3. Dokpil Full PLJKK - Template Perencanaan.docx", SHEET_DOKPIL_PL),
            ("5. BA PLJKK - Template Perencanaan.docx", SHEET_BA_PL),
        ],
    },
    "PL_PENGAWASAN": {
        "jenis_pl": "JKK",
        "label": "PL Konsultansi Pengawasan",
        "folder_name": "Pengawasan",
        "excel_template": "0. BAPLJKK - Template Pengawasan.xlsm",
        "word_map": [
            ("1. BA Reviu DPP PLJKK - Template Pengawasan.docx", SHEET_REVIU_PL),
            ("2. Isi Reviu PLJKK - Template Pengawasan.docm", SHEET_REVIU_PL),
            ("3. Dokpil Full PLJKK - Template Pengawasan.docx", SHEET_DOKPIL_PL),
            ("5. BA PLJKK - Template Pengawasan.docx", SHEET_BA_PL),
        ],
    },
    "PL_KONSTRUKSI": {
        "jenis_pl": "PK",
        "label": "PL Pekerjaan Konstruksi",
        "folder_name": "Konstruksi",
        "excel_template": "0. BAPLPK- Template.xlsm",
        "word_map": [
            ("1. BA Reviu PLPK - Template.docx", SHEET_BA_PL),
            ("2. Isi Reviu PLPK - Template.docm", SHEET_REVIU_PL),
            ("3. Dokpil Full PK - Template.docx", SHEET_DOKPIL_PL),
            ("5. BA PLPK - Template.docx", SHEET_BA_PL),
            ("7. BA Dengan Timpang PLPK - Template.docx", SHEET_BA_PL),
        ],
    },
}


def detect_pl_workflow(row=None, jenis_pl=None):
    """Deteksi subjenis PL dari metadata paket tanpa input manual tambahan."""
    row = row or {}
    text = " ".join(str(row.get(k) or "") for k in (
        "nama_paket", "uraian_pekerjaan", "metode_pengadaan", "jenis_pekerjaan",
    )).lower()
    jenis = str(jenis_pl or row.get("jenis_pl") or "").upper().strip()
    # jenis_pl dari SPSE adalah sumber utama. Paket JKK bisa menyebut
    # "konstruksi" di nama pekerjaan (mis. pengawasan konstruksi), tetapi
    # workbook-nya tetap JKK, bukan PLPK.
    if jenis in {"PK", "PLPK"}:
        return "PL_KONSTRUKSI"
    if jenis in {"JKK", "PLJKK"}:
        if any(k in text for k in ("pengawasan", "supervisi", "manajemen konstruksi", "mk ")):
            return "PL_PENGAWASAN"
        return "PL_PERENCANAAN"
    # Fallback hanya untuk row lama yang belum memiliki jenis_pl.
    if any(k in text for k in ("konstruksi", "pembangunan", "pagar", "paving", "gapura", "los ", "pengurugan", "normalisasi")):
        return "PL_KONSTRUKSI"
    if any(k in text for k in ("pengawasan", "supervisi", "manajemen konstruksi", "mk ")):
        return "PL_PENGAWASAN"
    return "PL_PERENCANAAN"


def pl_workflow_config(workflow):
    key = str(workflow or "").upper()
    if key not in PL_WORKFLOW_REGISTRY:
        raise KeyError(f"Workflow PL tidak dikenal: {workflow}")
    return PL_WORKFLOW_REGISTRY[key]


def _pl_template_set_complete(template_dir, workflow_cfg):
    """True jika donor memiliki semua file yang diminta registry workflow."""
    required = [workflow_cfg["excel_template"]]
    required.extend(name for name, _sheet in workflow_cfg["word_map"])
    return all(os.path.isfile(os.path.join(template_dir, name)) for name in required)


def pl_workflow_template_dir(workflow, root=None):
    """Resolve donor V2 lengkap; fallback ke donor legacy bila belum lengkap.

    Folder V2 boleh sudah dibuat bertahap. Jangan memilihnya hanya karena
    direktorinya ada: setup membutuhkan seluruh file pada ``word_map``.
    """
    root = root or os.path.join(POKJA_ROOT, "Paket Experiment - Pengadaan Langsung")
    cfg = pl_workflow_config(workflow)
    v2 = os.path.join(root, "V2 - Template PL", cfg["folder_name"])
    legacy_name = "Development - PL - PK" if cfg["jenis_pl"] == "PK" else "Development - PL - JKK"
    legacy = os.path.join(root, legacy_name)
    if not os.path.isdir(legacy):
        legacy = TEMPLATE_DIR_PL_PK if cfg["jenis_pl"] == "PK" else TEMPLATE_DIR_PL
    if _pl_template_set_complete(v2, cfg):
        return v2
    if _pl_template_set_complete(legacy, cfg):
        return legacy
    # Kembalikan kandidat yang paling informatif agar preflight setup dapat
    # melaporkan file mana yang hilang, bukan menyamarkan masalah konfigurasi.
    return v2 if os.path.isdir(v2) else legacy

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
    "pdf_minimalis":           ("WORD_BA",     "SHEET_BA",     "BA_Minimalis_{kode}.pdf"),
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
