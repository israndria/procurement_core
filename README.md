# Procurement Core

An open-source automation toolkit for Indonesian government procurement working groups (*Kelompok Kerja / Pokja*). Automates the entire procurement workflow — from scraping tender data off government portals, to generating legally-compliant procurement documents, to syncing schedules with Google Calendar.

Built for procurement officers under Indonesia's national procurement framework (Perpres 16/2018 jo. 12/2021 jo. 46/2025), where document accuracy and audit trails are legally required.

## Components

### 🔍 SPSE Scraper (`V20_Scrapper.py`)
Streamlit app that scrapes procurement data from LPSE (Layanan Pengadaan Secara Elektronik) portals across Indonesia. Stores tender and non-tender package data to Supabase for tracking and analysis.

### 📅 Tender Scheduler (`V19_Scheduler.py`)
Streamlit app for managing procurement schedules. Syncs tender timelines directly to Google Calendar so procurement officers never miss a deadline.

### 📄 Document Generator (`word_merge.py`)
Python bridge called by Excel VBA to merge procurement data into Word templates — generating official documents like *Berita Acara Pembukaan*, *Reviu*, and *Dokpil* as PDFs without blocking Excel.

### 🌐 Web Data Importer (`import_web_data.py`)
Extracts structured data from LPSE HTML pages and PDFs, outputs a JSON intermediary file consumed by Excel VBA — bridging government web portals and local document generation.

### 📁 Package Setup (`setup_paket_baru.py`)
CLI tool that creates a new procurement package folder with the correct template files and automatically links Word mail merge data sources to the Excel workbook via XML manipulation — no manual relinking needed.

### 💉 VBA Injector (`inject_buttons.py`)
Injects 14 VBA modules and UI buttons into Excel workbooks (`.xlsm`) via COM automation. Ensures all procurement packages always have the latest macro code.

## VBA Modules

| Module | Purpose |
|---|---|
| `ModWordLink` | Open/print/export Word documents from Excel; import LPSE data |
| `ModTerbilang` | Convert numbers to Indonesian words (*terbilang*) — built-in, no add-in needed |
| `ModUtilitas` | PDF export, direct printing, date conversion utilities |
| `ModNavigator` | Sheet navigation, table of contents generation |
| `ModKodeUnik` | Auto-generate unique procurement package codes |
| `ModAutoFit` | Auto-fit row heights across all sheets |
| `ModDraftPaket` | Sync draft tender package data |
| `ModDraftPaketPL` | Sync direct procurement (*Pengadaan Langsung*) package data |
| `ModKKEvaluasi` | Evaluation committee scoring sheet automation |
| `ModInputBA` | Centralized input hub for *Berita Acara* documents |
| `ModSyncDraft` | Two-way sync between Excel and Supabase draft data |
| `ModJawabanReviu` | Save/load procurement review answers without closing documents |
| `ModKodeUnikPL` | Unique code generation for direct procurement packages |

## Architecture

```
LPSE Website (HTML/PDF)
  → import_web_data.py       # extract fields via regex/BeautifulSoup
  → _import_lpse.json        # temp intermediary
  → VBA ImportHTML           # reads JSON → fills Excel "1. Input Data"
  → Excel formulas           # propagate to output sheets
  → word_merge.py            # reads Excel via COM → merges Word templates
  → Final PDFs               # BA PK, Reviu, Dokpil
```

## Requirements

- Windows only (uses `win32com` for Excel/Word COM automation)
- Python 3.10+ (WinPython portable recommended)
- Microsoft Excel + Word
- Dependencies: `supabase`, `streamlit`, `python-docx`, `openpyxl`, `pdfplumber`, `requests`, `beautifulsoup4`

## Running

```bash
# Tender Scheduler
python -m streamlit run V19_Scheduler.py

# SPSE Scraper
python -m streamlit run V20_Scrapper.py

# Setup new procurement package
python setup_paket_baru.py "19. Pokja 091"

# Inject VBA into all .xlsm files
"Inject All VBA.bat"
```

## Context

Indonesia processes **IDR 800+ trillion (~$50B USD)** in government procurement annually across thousands of agencies. Each procurement package requires multiple legally-mandated documents with specific formats. This toolkit automates that paper trail — reducing document preparation time from days to hours for procurement officers who otherwise do everything manually in Excel and Word.

## License

See `license.txt`
