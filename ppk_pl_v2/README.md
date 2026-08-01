# PPK PL V2 — deployment portable

Desain deployment:

- `procurement_core/ppk_pl_v2` = source generator, injector, dan launcher yang disimpan di Git.
- Workbook paket, template Word, output PDF/Word, snapshot, dan log = tetap di Google Drive.
- Workbook `.xlsm` tidak dimasukkan ke repo publik karena dapat berisi data personal dan layout manual user.
- VBA tetap memanggil `ThisWorkbook.Path\__ppk_runtime\generate_dokumen_ppk.bat`; launcher meneruskan paket ke source clone lokal.

## Setup PC kantor

1. Clone/pull `procurement_core` ke disk lokal, bukan Google Drive.
2. Pastikan Python portable lokal memiliki `openpyxl`, `python-docx`, `pywin32`, dan Microsoft Excel/Word terpasang.
3. Dari PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\ppk_pl_v2\setup_ppk_office.ps1 -DriveRoot "G:\Other computers\My Laptop\@ POKJA 2026"
```

Untuk simulasi/verifikasi tanpa perubahan persisten:

```powershell
.\ppk_pl_v2\setup_ppk_office.ps1 -DriveRoot "G:\Other computers\My Laptop\@ POKJA 2026" -VerifyOnly
```

Script akan:

- memilih `POKJA_PYTHON` lokal;
- menetapkan `POKJA_V19_ROOT`, `POKJA_DRIVE_ROOT`, `POKJA_SECRET_ROOT`, dan `POKJA_PPK_ROOT` per user;
- memvalidasi dependency Python/COM;
- memvalidasi workbook, folder `Konstruksi`, `Perencanaan`, `Pengawasan`, routing `PK/KP/KPWAS`, dan template aktif;
- memasang launcher portable ke folder `__ppk_runtime` paket, dengan backup otomatis jika launcher lama berbeda.

Setelah setup, tutup-buka Excel. Tombol VBA lama tetap dipakai. Saat generate, source dibaca dari clone lokal dan workbook/template dibaca dari folder paket Google Drive.

## Injector workbook

Workbook yang sudah aktif tidak perlu diinjeksi ulang. Jika membuat master baru, Excel harus ditutup lalu jalankan:

```powershell
& "$env:POKJA_PYTHON" .\ppk_pl_v2\inject_vba_ppk.py `
  "G:\Other computers\My Laptop\@ POKJA 2026\Paket Experiment - Pengadaan Langsung\V2 - Template PPK PL\0. Master_Data_PL_PPK.xlsm"
```

Jangan menjalankan `openpyxl.save()` terhadap `.xlsm`. Injector memakai Excel COM agar VBA, merge, warna, dan tombol tetap dipertahankan.

## Operasional antar-PC

- Git: commit/push di satu PC, lalu `git pull --ff-only` manual di PC lain.
- Google Drive: workbook/template/output tersinkron otomatis.
- Secret Supabase: simpan lokal di `%LOCALAPPDATA%\POKJA2026\Secrets`, jangan di Git/Drive.
- Launcher tidak melakukan `git pull` otomatis.
