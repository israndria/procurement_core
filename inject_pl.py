"""
Inject VBA ModDraftPaketPL ke 0. BAPLJKK - Template.xlsm (dan semua copy paket PL).
Terpisah dari inject_buttons.py (PK/tender) — tidak ada dependency silang.

Usage:
    python inject_pl.py                           # inject ke semua .xlsm BAPLJKK di POKJA root
    python inject_pl.py "path/to/file.xlsm"       # inject ke file spesifik
"""
import win32com.client
import pythoncom
import os
import hashlib
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
BAS_FILE = SCRIPT_DIR / "ModDraftPaketPL.bas"
MOD_NAME = "ModDraftPaketPL"
WORDLINK_BAS_FILE = SCRIPT_DIR / "ModWordLink.bas"
WORDLINK_MOD_NAME = "ModWordLink"
LAYOUT_MODULES = ("modBarisItem", "modAutoLayoutNego")
XL_CALCULATION_MANUAL = -4135
XL_AUTOMATION_SECURITY_LOW = 1

_BACKUP_DIRECTORY_NAMES = {
    ".vba-backup",
    "_backup",
    "_backup_archive",
    "backup",
    "backups",
    "archive",
    "archives",
}
_BACKUP_FILE_MARKERS = (
    ".bak",
    ".backup",
    "backup_",
    "-backup",
    ".before-",
)

# Geometry resmi PLPK dari template Konstruksi. Jangan memakai geometry
# PLJKK/legacy: injector ini dipakai bersama untuk dua keluarga PL, sehingga
# layout dipilih setelah workbook dikenali sebagai PLPK.
PLPK_BUTTON_GEOMETRY = {
    "btnBukaDokpil_PL": (657.9, 175.0, 130.2, 40.0),
    "btnRelinkPL": (793.2, 175.0, 129.9, 40.0),
    "btnRefreshDataPL": (927.6, 175.0, 129.9, 40.0),
    "btnBukaBA_PL": (657.9, 218.5, 130.2, 27.0),
    "btnBukaReviu_PL": (793.2, 218.5, 129.9, 27.0),
    "btnCetakBAReviu_PL": (657.9, 248.5, 130.2, 27.7),
    "btnCetakDokpil_PL": (793.2, 248.5, 129.9, 27.7),
    "btnCetakReviu_PL": (657.9, 280.6, 130.2, 28.5),
    "btnGabungReviu_PL": (793.2, 280.6, 129.9, 28.5),
    "btnMuatHPS_PL": (929.0, 249.8, 129.9, 28.2),
    "btnIsiEvaluasiPL": (927.7, 218.6, 129.9, 27.0),
    "btnCetakBAPLJKK": (657.9, 312.5, 130.2, 28.0),
    "btnGabungBAPLJKK": (793.2, 312.5, 129.9, 28.0),
    "btnSaveInputData": (926.7, 280.4, 130.2, 40.0),
    "btnLoadInputData": (925.7, 323.3, 130.4, 39.4),
}

# Event workbook untuk BAPLJKK — relink tetap manual, input tanggal dipermudah.
WORKBOOK_OPEN_DATE_CODE = (
    "Private Sub Workbook_Open()\n"
    "    ' Workbook_Open dimatikan — relink manual lewat tombol Relink Word\n"
    "End Sub\n"
    "\n"
    "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n"
    "    On Error GoTo SafeExit\n"
    "    If Sh.Name <> \"@ Master Data\" And Sh.Name <> \"@ Evaluasi\" Then Exit Sub\n"
    "    If Target.CountLarge <> 1 Or Target.Column <> 3 Then Exit Sub\n"
    "    If InStr(1, CStr(Sh.Cells(Target.Row, 2).Value), \"tanggal\", vbTextCompare) = 0 Then Exit Sub\n"
    "\n"
    "    ' Nilai tanggal Excel asli harus dibiarkan numeric.\n"
    "    If IsNumeric(Target.Value2) And IsDate(Target.Value) Then\n"
    "        Target.NumberFormat = \"dd mmmm yyyy\"\n"
    "        Exit Sub\n"
    "    End If\n"
    "\n"
    "    Dim raw As String\n"
    "    raw = Trim$(CStr(Target.Value2))\n"
    "    If raw = \"\" Then Exit Sub\n"
    "\n"
    "    Dim hasil As Variant\n"
    "    hasil = ParseTanggalPL(raw)\n"
    "    If IsEmpty(hasil) Then Exit Sub\n"
    "\n"
    "    Application.EnableEvents = False\n"
    "    Target.NumberFormat = \"dd mmmm yyyy\"\n"
    "    Target.Value = hasil\n"
    "\n"
    "SafeExit:\n"
    "    Application.EnableEvents = True\n"
    "End Sub\n"
    "\n"
    "Private Function ParseTanggalPL(ByVal raw As String) As Variant\n"
    "    Dim bulanArr As Variant\n"
    "    bulanArr = Array(\"Januari\", \"Februari\", \"Maret\", \"April\", \"Mei\", \"Juni\", _\n"
    "                     \"Juli\", \"Agustus\", \"September\", \"Oktober\", \"November\", \"Desember\")\n"
    "    ParseTanggalPL = Empty\n"
    "\n"
    "    Dim s As String: s = Trim$(raw)\n"
    "    Dim tgl As Long, bln As Long, thn As Long\n"
    "    Dim sep As String, parts() As String\n"
    "    If InStr(s, \"/\") > 0 Then sep = \"/\"\n"
    "    If InStr(s, \"-\") > 0 Then sep = \"-\"\n"
    "    If InStr(s, \".\") > 0 Then sep = \".\"\n"
    "    If sep = \"\" And InStr(s, \" \" ) > 0 Then sep = \" \"\n"
    "\n"
    "    On Error GoTo InvalidDate\n"
    "    If sep <> \"\" Then\n"
    "        parts = Split(s, sep)\n"
    "        If UBound(parts) < 2 Then GoTo InvalidDate\n"
    "        tgl = CLng(Trim$(parts(0)))\n"
    "        If IsNumeric(Trim$(parts(1))) Then\n"
    "            bln = CLng(Trim$(parts(1)))\n"
    "        Else\n"
    "            Dim iBulan As Long\n"
    "            For iBulan = 0 To 11\n"
    "                If LCase$(Trim$(parts(1))) = LCase$(CStr(bulanArr(iBulan))) Then\n"
    "                    bln = iBulan + 1\n"
    "                    Exit For\n"
    "                End If\n"
    "            Next iBulan\n"
    "        End If\n"
    "        thn = CLng(Trim$(parts(2)))\n"
    "    Else\n"
    "        If Not IsNumeric(s) Then GoTo InvalidDate\n"
    "        If Len(s) = 8 Then\n"
    "            tgl = CLng(Left$(s, 2))\n"
    "            bln = CLng(Mid$(s, 3, 2))\n"
    "            thn = CLng(Right$(s, 4))\n"
    "        ElseIf Len(s) = 6 Then\n"
    "            tgl = CLng(Left$(s, 2))\n"
    "            bln = CLng(Mid$(s, 3, 2))\n"
    "            thn = 2000 + CLng(Right$(s, 2))\n"
    "        ElseIf Len(s) >= 5 Then\n"
    "            thn = CLng(Right$(s, 4))\n"
    "            s = Left$(s, Len(s) - 4)\n"
    "            If Len(s) = 3 Then\n"
    "                tgl = CLng(Left$(s, 2))\n"
    "                bln = CLng(Right$(s, 1))\n"
    "            ElseIf Len(s) = 4 Then\n"
    "                tgl = CLng(Left$(s, 2))\n"
    "                bln = CLng(Right$(s, 2))\n"
    "            Else\n"
    "                GoTo InvalidDate\n"
    "            End If\n"
    "        Else\n"
    "            GoTo InvalidDate\n"
    "        End If\n"
    "    End If\n"
    "\n"
    "    If tgl < 1 Or tgl > 31 Or bln < 1 Or bln > 12 Or thn < 2000 Or thn > 2099 Then GoTo InvalidDate\n"
    "    Dim dt As Date\n"
    "    dt = DateSerial(thn, bln, tgl)\n"
    "    If Day(dt) <> tgl Or Month(dt) <> bln Or Year(dt) <> thn Then GoTo InvalidDate\n"
    "    ParseTanggalPL = dt\n"
    "    Exit Function\n"
    "\n"
    "InvalidDate:\n"
    "    ParseTanggalPL = Empty\n"
    "End Function\n"
)

# Reuse the guarded date-change event, while adding the dynamic layout hooks.
# The hooks are deliberately incremental; no full-workbook recalculation occurs
# when a user opens a PL workbook.
WORKBOOK_OPEN_CODE = WORKBOOK_OPEN_DATE_CODE.replace(
    "Private Sub Workbook_Open()\n",
    "Private Sub Workbook_Open()\n"
    "    On Error Resume Next\n"
    "    modAutoLayoutNego.ResetCacheLayout\n"
    "    modAutoLayoutNego.PasangShortcutRapikan\n"
    "    modAutoLayoutNego.RapikanDaftarNego True, False\n"
    "    modBarisItem.RefreshBarisItem True\n"
    "    On Error GoTo SafeExit\n"
    "SafeExit:\n",
    1,
)

HPS_EVENT_CODE = (
    "' BEGIN POKJA_AUTO_BARIS_ITEM\n"
    "Private Sub Worksheet_Calculate()\n"
    "    modBarisItem.RefreshBarisItem False\n"
    "End Sub\n"
    "\n"
    "Private Sub Worksheet_Change(ByVal Target As Range)\n"
    "    If Intersect(Target, Me.Range(\"A2:A501\")) Is Nothing Then Exit Sub\n"
    "    modBarisItem.RefreshBarisItem False\n"
    "End Sub\n"
    "' END POKJA_AUTO_BARIS_ITEM"
)

NEGO_EVENT_CODE = (
    "' BEGIN POKJA_AUTO_LAYOUT_NEGO\n"
    "Private Sub Worksheet_Calculate()\n"
    "    On Error GoTo SafeExit\n"
    "    modAutoLayoutNego.AutoRapikanJikaPerlu Me\n"
    "SafeExit:\n"
    "End Sub\n"
    "\n"
    "Private Sub Worksheet_Activate()\n"
    "    On Error GoTo SafeExit\n"
    "    modAutoLayoutNego.PasangShortcutRapikan\n"
    "    modAutoLayoutNego.AutoRapikanJikaPerlu Me\n"
    "SafeExit:\n"
    "End Sub\n"
    "' END POKJA_AUTO_LAYOUT_NEGO"
)


def _validate_vba_source(content: str, module_name: str = MOD_NAME) -> None:
    """Tolak source BAS rusak sebelum menyentuh workbook."""
    if f'Attribute VB_Name = "{module_name}"' not in content:
        raise ValueError(f"Attribute VB_Name {module_name} tidak ditemukan")
    if "%%SUPABASE_URL%%" in content or "%%SUPABASE_KEY%%" in content:
        raise ValueError("Placeholder secret VBA belum tersubstitusi")

    malformed_formula = [
        (line_no, line)
        for line_no, line in enumerate(content.splitlines(), 1)
        if (".Formula" in line or ".FormulaLocal" in line) and '\\"' in line
    ]
    if malformed_formula:
        line_no, _ = malformed_formula[0]
        raise ValueError(
            f"VBA formula memakai escape Python/JSON (\\\\\\\") di baris {line_no}"
        )


def _create_backup(filepath: str) -> Path:
    """Backup unik dan recoverable sebelum VBA workbook diubah."""
    source = Path(filepath)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    source_id = hashlib.sha1(str(source).encode("utf-8")).hexdigest()[:10]
    safe_stem = source.stem[:80].rstrip(" .")
    backup_name = f"{safe_stem}.{source_id}.before-{MOD_NAME}-{stamp}{source.suffix}"

    # Ukur path tujuan sebenarnya. Mengukur source.name saja tidak cukup:
    # suffix timestamp + nama modul dapat mendorong path paket melewati
    # MAX_PATH walaupun path source masih di bawah ambang.
    backup_dir = source.parent / ".vba-backup"
    backup = backup_dir / backup_name
    if len(str(backup)) >= 240:
        configured_root = os.environ.get("POKJA_DRIVE_ROOT", "").strip()
        pokja_root = Path(configured_root) if configured_root else None
        if not pokja_root or not pokja_root.exists():
            pokja_root = next(
                (parent for parent in source.parents if parent.name == "@ POKJA 2026"),
                SCRIPT_DIR,
            )
        backup_dir = pokja_root / ".vba-backup"
        backup = backup_dir / backup_name
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, backup)
    return backup


def _is_backup_workbook_path(filepath: str | os.PathLike) -> bool:
    """Tolak workbook backup/archive sebagai target injector VBA."""
    path = Path(filepath)
    if any(part.casefold() in _BACKUP_DIRECTORY_NAMES for part in path.parts[:-1]):
        return True

    name = path.name.casefold()
    stem = path.stem.casefold()
    return (
        name.startswith("~$")
        or any(marker in stem for marker in _BACKUP_FILE_MARKERS)
    )


def _harden_evaluasi_date_formulas(ws_eval) -> int:
    """Hardening formula turunan tanggal agar memakai nilai tanggal Excel."""
    weekday = '"Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"'
    date_labels = {
        "tanggal pembukaan penawaran",
        "tanggal pembuktian kualifikasi",
        "tanggal klarifikasi & negosiasi",
        "tanggal ba hasil pengadaan langsung",
    }
    patched = 0
    for row in range(1, 100):
        label = str(ws_eval.Cells(row, 2).Value or "").strip().casefold()
        if label not in date_labels:
            continue
        date_ref = f"C{row}"
        day_ref = f"C{row + 1}"
        current = ws_eval.Cells(row, 3).Value
        serial = _coerce_eval_date_serial(current)
        if serial is not None:
            ws_eval.Cells(row, 3).Value = serial
        ws_eval.Cells(row, 3).NumberFormat = "dd mmmm yyyy"
        ws_eval.Cells(row + 1, 3).Formula = f'=IF({date_ref}="","",DAY({date_ref}))'
        ws_eval.Cells(row + 2, 3).Formula = (
            f'=IF({date_ref}="","",CHOOSE(WEEKDAY({date_ref}),{weekday}))'
        )
        ws_eval.Cells(row + 2, 4).Formula = f'=IF({date_ref}="","",{date_ref})'
        ws_eval.Cells(row + 2, 4).NumberFormat = "dd mmmm yyyy"
        # Donor workbook tidak seragam: sebagian memiliki UDF terbilang1,
        # sebagian (termasuk template PL pusat) hanya memiliki terbilang.
        # IFERROR membuat hasil lintas-template tetap valid tanpa menyimpan
        # #NAME? ke PDF saat salah satu nama UDF tidak tersedia.
        ws_eval.Cells(row + 3, 3).Formula = (
            f'=IF({day_ref}="","",IFERROR(terbilang1({day_ref}),terbilang({day_ref})))'
        )
        ws_eval.Cells(row + 4, 3).Formula = (
            f'=IF({date_ref}="","",CHOOSE(MONTH({date_ref}),'
            '"Januari","Februari","Maret","April","Mei","Juni",'
            '"Juli","Agustus","September","Oktober","November","Desember"))'
        )
        ws_eval.Cells(row + 4, 4).Formula = f'=IF({date_ref}="","",MONTH({date_ref}))'
        ws_eval.Cells(row + 4, 4).NumberFormat = "General"
        next_label = str(ws_eval.Cells(row + 5, 2).Value or "").strip().casefold()
        if next_label == "tahun":
            ws_eval.Cells(row + 5, 3).Formula = f'=IF({date_ref}="","",YEAR({date_ref}))'
        elif next_label == "tahun terbilang":
            ws_eval.Cells(row + 5, 3).Formula = (
                f'=IF({date_ref}="","",IFERROR(terbilang1(YEAR({date_ref})),terbilang(YEAR({date_ref}))))'
            )
        patched += 1
    return patched


def _coerce_eval_date_serial(value):
    """Konversi tanggal sumber template ke serial Excel tanpa locale COM."""
    if isinstance(value, datetime):
        value = value.date()
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return (value - datetime(1899, 12, 30).date()).days
    text = str(value or "").strip()
    if not text:
        return None
    months = {
        "januari": 1, "februari": 2, "maret": 3, "april": 4,
        "mei": 5, "juni": 6, "juli": 7, "agustus": 8,
        "september": 9, "oktober": 10, "november": 11, "desember": 12,
        "january": 1, "february": 2, "march": 3, "may": 5,
        "june": 6, "july": 7, "august": 8, "october": 10,
        "december": 12,
    }
    import re
    match = re.fullmatch(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if not match:
        return None
    month = months.get(match.group(2).casefold())
    if month is None:
        return None
    try:
        parsed = datetime(int(match.group(3)), month, int(match.group(1)))
    except ValueError:
        return None
    return (parsed.date() - datetime(1899, 12, 30).date()).days


def _inject_marked_sheet_event(vb_project, workbook, sheet_name, event_code,
                               start_marker, end_marker, *, legacy_event=False):
    """Pasang event sheet secara idempotent dan fail-closed.

    Event custom milik paket tidak boleh ditimpa diam-diam. Hanya blok marker
    milik injector atau pola legacy layout yang dikenal yang boleh diganti.
    """
    ws = workbook.Sheets(sheet_name)
    cm = vb_project.VBComponents(ws.CodeName).CodeModule
    current = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
    start = current.find(start_marker)
    if start >= 0:
        end = current.find(end_marker, start)
        if end < 0:
            raise RuntimeError(f"Marker event {sheet_name} tidak lengkap")
        current = current[:start] + current[end + len(end_marker):]
    elif legacy_event and (
        "Private Sub Worksheet_Change" in current
        and "Private Sub FixRowHeight" in current
        and "mergeArea" in current
    ):
        current = ""
    elif any(marker in current for marker in (
        "Private Sub Worksheet_Change",
        "Private Sub Worksheet_Calculate",
        "Private Sub Worksheet_Activate",
    )):
        raise RuntimeError(f"Sheet {sheet_name} memiliki event custom di luar pola legacy")

    new_code = current.rstrip()
    if new_code:
        new_code += "\n\n"
    new_code += event_code
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(new_code)
    return cm.CountOfLines


def _remove_marked_sheet_event(vb_project, workbook, sheet_name,
                               start_marker, end_marker):
    """Hapus blok event milik injector tanpa menyentuh event custom paket."""
    ws = workbook.Sheets(sheet_name)
    cm = vb_project.VBComponents(ws.CodeName).CodeModule
    current = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
    start = current.find(start_marker)
    if start < 0:
        return cm.CountOfLines
    end = current.find(end_marker, start)
    if end < 0:
        raise RuntimeError(f"Marker event {sheet_name} tidak lengkap")
    current = current[:start] + current[end + len(end_marker):]
    current = current.rstrip()
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    if current:
        cm.AddFromString(current)
    return cm.CountOfLines


def inject_pl(filepath: str):
    filepath = os.path.abspath(filepath)
    print(f"\nInjecting PL module to: {filepath}")

    if _is_backup_workbook_path(filepath):
        print("  [ERROR] Target berada di backup/archive; injector dihentikan.")
        return False

    if not os.path.exists(filepath):
        print(f"  [ERROR] File tidak ditemukan: {filepath}")
        return False

    if not BAS_FILE.exists():
        print(f"  [ERROR] ModDraftPaketPL.bas tidak ditemukan: {BAS_FILE}")
        return False

    # Baca + substitusi secret
    content = BAS_FILE.read_text(encoding="utf-8")
    if "%%SUPABASE_URL%%" in content or "%%SUPABASE_KEY%%" in content:
        from dotenv import load_dotenv
        canonical_env = SCRIPT_DIR.parent / "Secrets" / "secret_supabase.env"
        configured_root = os.environ.get("POKJA_SECRET_ROOT", "").strip()
        configured_env = Path(configured_root) / "secret_supabase.env" if configured_root else None
        env_path = canonical_env if canonical_env.exists() else (configured_env or canonical_env)
        load_dotenv(env_path)
        sb_url = os.environ.get("SUPABASE_URL", "").strip('"')
        sb_key = os.environ.get("SUPABASE_KEY", "").strip('"')
        if not sb_url or not sb_key:
            print("  [ERROR] SUPABASE_URL / SUPABASE_KEY tidak ditemukan di secret_supabase.env")
            return False
        content = content.replace("%%SUPABASE_URL%%", sb_url)
        content = content.replace("%%SUPABASE_KEY%%", sb_key)

    try:
        _validate_vba_source(content)
    except ValueError as exc:
        print(f"  [ERROR] Preflight VBA gagal: {exc}")
        return False

    # Attribute VB_Name menentukan nama module hasil Import, BUKAN nama file
    # temp. Pakai nama sementara dulu agar tidak bentrok dgn module lama yang
    # masih ada saat proses import (baru direname ke MOD_NAME setelah module
    # lama dihapus) -- tahan interupsi di tengah proses.
    tmp_mod_name = f"{MOD_NAME}_NEW"
    content_tmp = content.replace(f'Attribute VB_Name = "{MOD_NAME}"', f'Attribute VB_Name = "{tmp_mod_name}"')

    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False, mode="w", encoding="utf-8")
    tmp.write(content_tmp)
    tmp.close()
    tmp_path = tmp.name

    pythoncom.CoInitialize()
    excel = None
    wb = None
    backup_path = None

    try:
        backup_path = _create_backup(filepath)
        print(f"  [BACKUP] {backup_path}")

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        # UDF seperti terbilang1 harus terdaftar saat workbook dibuka. Jangan
        # pernah membuka workbook untuk ditulis dalam mode ForceDisable (3):
        # save setelah formula tersentuh dapat menyimpan cache UDF sebagai
        # #NAME?. Event workbook tetap dimatikan agar Workbook_Open tidak jalan.
        excel.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
        excel.EnableEvents = False
        try:
            wb = excel.Workbooks.Open(filepath, 0, False)
        except Exception as open_error:
            # Workbook hasil patch manual kadang valid sebagai ZIP tetapi
            # ditolak parser Excel. xlRepairFile=1 meminta Excel memulihkan
            # workbook saat dibuka; tetap dipakai hanya sebagai fallback.
            print(f"  [WARN] Open normal gagal, coba Excel repair: {open_error}")
            wb = excel.Workbooks.Open(filepath, 0, False, None, None, None, None, None, 1)
            print("  [OK] Excel repair open berhasil")
        # Excel menolak mengubah Calculation sebelum workbook terbuka pada
        # sebagian versi. Set setelah Open, sebelum perubahan struktural.
        # Injector hanya mengubah VBA/shape; cached formula dipertahankan.
        excel.Calculation = XL_CALCULATION_MANUAL
        vb = wb.VBProject

        # PLPK dikenali dari struktur workbook, bukan nama file/folder.
        # PLJKK Pengawasan/Perencanaan harus tetap mendapat fitur umum PL,
        # tetapi tidak boleh menerima otomasi 7.2 Dengan Nego.
        master_ws = wb.Sheets("@ Master Data")
        is_pk = str(master_ws.Cells(76, 1).Value or "").strip() == "5. DATA PESERTA"
        active_layout_modules = LAYOUT_MODULES if is_pk else ()

        # PL tidak memakai generator kode unik otomatis. Bersihkan modul/button
        # legacy yang bisa ikut terbawa dari injector tender/umum.
        for legacy_name in ("ModKodeUnik", "ModKodeUnikPL"):
            for comp in list(vb.VBComponents):
                if comp.Name == legacy_name:
                    vb.VBComponents.Remove(comp)
                    print(f"  {legacy_name} legacy dihapus")
                    break

        # Import module baru DULU dengan nama sementara (tahan interupsi:
        # kalau proses mati di sini, module lama MOD_NAME masih utuh)
        imported = vb.VBComponents.Import(tmp_path)
        old_comp = None
        for comp in vb.VBComponents:
            if comp.Name == MOD_NAME:
                old_comp = comp
                break
        if old_comp:
            vb.VBComponents.Remove(old_comp)
            print(f"  {MOD_NAME} lama dihapus")
        imported.Name = MOD_NAME
        print(f"  [OK] {imported.Name} imported ({imported.CodeModule.CountOfLines} baris)")

        imported_text = imported.CodeModule.Lines(1, imported.CodeModule.CountOfLines)
        # VBE menormalkan kapitalisasi identifier dan inline-comment mengikuti
        # symbol table workbook. Karena itu validasi identitas byte-per-byte
        # tidak stabil; pastikan jumlah baris + entry point utama utuh.
        expected_lines = len(content.splitlines()) - 1  # Attribute VB_Name
        if imported.CodeModule.CountOfLines != expected_lines:
            raise ValueError(
                f"Jumlah baris {MOD_NAME} berubah setelah import: "
                f"{imported.CodeModule.CountOfLines} != {expected_lines}"
            )
        for marker in (
            "Public Sub IsiDataPLByKode(",
            "Public Sub RefreshDataPL()",
            "Public Sub IsiEvaluasiPLStandalone()",
        ):
            if marker.casefold() not in imported_text.casefold():
                raise ValueError(f"Entry point VBA hilang setelah import: {marker}")
        _validate_vba_source(
            f'Attribute VB_Name = "{MOD_NAME}"\n{imported_text}'
        )
        print(f"  [OK] {MOD_NAME} lolos verifikasi source pasca-import")

        # Tombol Gabung BA Reviu memakai ModWordLink. Paket PL lama sering
        # membawa modul ini dari donor lama, sehingga resolver Python-nya
        # tidak mengenali clone procurement_core lokal. Sinkronkan modul inti
        # di injector PL agar workbook baru maupun existing konsisten.
        if not WORDLINK_BAS_FILE.exists():
            raise FileNotFoundError(f"{WORDLINK_BAS_FILE} tidak ditemukan")
        wordlink_content = WORDLINK_BAS_FILE.read_text(encoding="utf-8")
        _validate_vba_source(wordlink_content, WORDLINK_MOD_NAME)
        wordlink_tmp = tempfile.NamedTemporaryFile(
            suffix=".bas", delete=False, mode="w", encoding="utf-8"
        )
        wordlink_tmp.write(
            wordlink_content.replace(
                f'Attribute VB_Name = "{WORDLINK_MOD_NAME}"',
                f'Attribute VB_Name = "{WORDLINK_MOD_NAME}_NEW"',
            )
        )
        wordlink_tmp.close()
        wordlink_tmp_path = wordlink_tmp.name
        imported_wordlink = vb.VBComponents.Import(wordlink_tmp_path)
        old_wordlink = None
        for comp in vb.VBComponents:
            if comp.Name == WORDLINK_MOD_NAME:
                old_wordlink = comp
                break
        if old_wordlink:
            vb.VBComponents.Remove(old_wordlink)
            print(f"  {WORDLINK_MOD_NAME} lama dihapus")
        imported_wordlink.Name = WORDLINK_MOD_NAME
        imported_wordlink_text = imported_wordlink.CodeModule.Lines(
            1, imported_wordlink.CodeModule.CountOfLines
        )
        expected_wordlink_lines = len(wordlink_content.splitlines()) - 1
        if imported_wordlink.CodeModule.CountOfLines != expected_wordlink_lines:
            raise ValueError(
                f"Jumlah baris {WORDLINK_MOD_NAME} berubah setelah import: "
                f"{imported_wordlink.CodeModule.CountOfLines} != {expected_wordlink_lines}"
            )
        _validate_vba_source(
            f'Attribute VB_Name = "{WORDLINK_MOD_NAME}"\n{imported_wordlink_text}',
            WORDLINK_MOD_NAME,
        )
        print(f"  [OK] {WORDLINK_MOD_NAME} imported ({imported_wordlink.CodeModule.CountOfLines} baris)")
        os.unlink(wordlink_tmp_path)

        # Auto-layout 7.2 Dengan Nego hanya untuk PLPK Konstruksi.
        # Import menggunakan nama sementara agar module lama tetap utuh bila
        # proses terhenti sebelum module baru tervalidasi.
        for layout_name in active_layout_modules:
            layout_bas = SCRIPT_DIR / f"{layout_name}.bas"
            if not layout_bas.exists():
                raise FileNotFoundError(f"{layout_bas} tidak ditemukan")
            layout_content = layout_bas.read_text(encoding="utf-8")
            _validate_vba_source(layout_content, layout_name)
            layout_tmp = tempfile.NamedTemporaryFile(
                suffix=".bas", delete=False, mode="w", encoding="utf-8"
            )
            layout_tmp.write(
                layout_content.replace(
                    f'Attribute VB_Name = "{layout_name}"',
                    f'Attribute VB_Name = "{layout_name}_NEW"',
                )
            )
            layout_tmp.close()
            try:
                imported_layout = vb.VBComponents.Import(layout_tmp.name)
                old_layout = None
                for comp in vb.VBComponents:
                    if comp.Name == layout_name:
                        old_layout = comp
                        break
                if old_layout:
                    vb.VBComponents.Remove(old_layout)
                    print(f"  {layout_name} lama dihapus")
                imported_layout.Name = layout_name
                layout_text = imported_layout.CodeModule.Lines(
                    1, imported_layout.CodeModule.CountOfLines
                )
                expected_layout_lines = len(layout_content.splitlines()) - 1
                if imported_layout.CodeModule.CountOfLines != expected_layout_lines:
                    raise ValueError(
                        f"Jumlah baris {layout_name} berubah setelah import: "
                        f"{imported_layout.CodeModule.CountOfLines} != {expected_layout_lines}"
                    )
                _validate_vba_source(
                    f'Attribute VB_Name = "{layout_name}"\n{layout_text}',
                    layout_name,
                )
                print(
                    f"  [OK] {layout_name} imported "
                    f"({imported_layout.CodeModule.CountOfLines} baris)"
                )
            finally:
                try:
                    os.unlink(layout_tmp.name)
                except OSError:
                    pass

        # Bersihkan modul auto-layout lama dari template/paket PLJKK bila
        # pernah diinjeksi oleh versi injector sebelumnya.
        if not is_pk:
            for layout_name in LAYOUT_MODULES:
                for comp in list(vb.VBComponents):
                    if comp.Name == layout_name:
                        vb.VBComponents.Remove(comp)
                        print(f"  {layout_name} auto-layout dihapus dari PLJKK")
                        break

        # Inject Workbook_Open ke ThisWorkbook
        this_wb_comp = None
        for comp in vb.VBComponents:
            if comp.Name == "ThisWorkbook":
                this_wb_comp = comp
                break
        if this_wb_comp:
            cm = this_wb_comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(WORKBOOK_OPEN_CODE if is_pk else WORKBOOK_OPEN_DATE_CODE)
            print(f"  [OK] Workbook_Open injected ({cm.CountOfLines} baris)")

        # Perubahan HPS dan kalkulasi 7.2 langsung memicu perapian. Event
        # custom di luar marker tidak ditimpa agar workbook paket tetap aman.
        try:
            if is_pk:
                count = _inject_marked_sheet_event(
                    vb, wb, "5. HPS", HPS_EVENT_CODE,
                    "' BEGIN POKJA_AUTO_BARIS_ITEM",
                    "' END POKJA_AUTO_BARIS_ITEM",
                )
                print(f"  [OK] Sheet 5. HPS auto-row events injected ({count} baris)")
            else:
                count = _remove_marked_sheet_event(
                    vb, wb, "5. HPS",
                    "' BEGIN POKJA_AUTO_BARIS_ITEM",
                    "' END POKJA_AUTO_BARIS_ITEM",
                )
                print(f"  [OK] Sheet 5. HPS auto-row events removed from PLJKK ({count} baris)")
        except Exception as event_error:
            print(f"  [WARN] Event sheet 5. HPS tidak dipasang: {event_error}")

        try:
            if is_pk:
                count = _inject_marked_sheet_event(
                    vb, wb, "7.2 Dengan Nego", NEGO_EVENT_CODE,
                    "' BEGIN POKJA_AUTO_LAYOUT_NEGO",
                    "' END POKJA_AUTO_LAYOUT_NEGO",
                    legacy_event=True,
                )
                print(f"  [OK] Sheet 7.2 Dengan Nego auto-layout events injected ({count} baris)")
            else:
                count = _remove_marked_sheet_event(
                    vb, wb, "7.2 Dengan Nego",
                    "' BEGIN POKJA_AUTO_LAYOUT_NEGO",
                    "' END POKJA_AUTO_LAYOUT_NEGO",
                )
                print(f"  [OK] Sheet 7.2 Dengan Nego auto-layout removed from PLJKK ({count} baris)")
        except Exception as event_error:
            print(f"  [WARN] Event sheet 7.2 Dengan Nego tidak dipasang: {event_error}")

        try:
            eval_count = _harden_evaluasi_date_formulas(wb.Sheets("@ Evaluasi"))
            print(f"  [OK] Formula tanggal @ Evaluasi di-hardening ({eval_count} blok)")
        except Exception as eval_error:
            print(f"  [WARN] Formula tanggal @ Evaluasi: {eval_error}")

        # Tombol di @ Master Data (Muat Paket PL + Isi Data PL sudah dihapus —
        # pengisian @ Master Data kini otomatis via COM saat buat folder).
        try:
            ws = wb.Sheets("@ Master Data")

            # Unprotect sheet sebelum modifikasi shape
            try:
                ws.Unprotect("pokja2026")
            except Exception:
                pass

            # Hapus tombol lama
            names_to_delete = []
            BTN_NAMES = ("btnMuatPL", "btnIsiPL", "btnKodeUnik", "btnBukaBA_PL", "btnBukaReviu_PL", "btnBukaDokpil_PL", "btnRelinkPL", "btnRefreshDataPL", "btnMuatHPS_PL", "btnCetakBAReviu_PL", "btnSyncDraftPL", "btnClearHighlightPL", "btnCetakDokpil_PL", "btnCetakReviu_PL", "btnGabungReviu_PL", "btnIsiEvaluasiPL", "btnCetakBAPLJKK", "btnGabungBAReviu", "btnGabungBAPLJKK", "btnSaveInputData", "btnLoadInputData")
            for shp in ws.Shapes:
                if shp.Name in BTN_NAMES:
                    names_to_delete.append(shp.Name)
            for name in names_to_delete:
                try:
                    ws.Shapes(name).Delete()
                    print(f"  Tombol lama {name} dihapus")
                except Exception:
                    pass

            BLUE    = (43, 87, 154)
            GREEN_C = (40, 167, 69)
            ORANGE  = (200, 100, 0)
            PURPLE  = (102, 51, 153)
            TEAL    = (0, 128, 128)

            # Layout tombol disimpan eksplisit agar injector tidak mengembalikan
            # tombol ke layout JKK saat workbook PLPK di-inject ulang.
            # Deteksi berdasarkan struktur @ Master Data, bukan nama file/folder.
            is_pk = str(ws.Cells(76, 1).Value or '').strip() == '5. DATA PESERTA'
            if is_pk:
                button_geometry = PLPK_BUTTON_GEOMETRY.copy()
                print('  Layout tombol: PLPK (template Konstruksi)')
            else:
                # Layout baku PLJKK.
                # Baseline manual PLJKK — disamakan dengan layout template
                # yang sudah dirapikan user (diverifikasi via Excel COM).
                _X = [661.4, 796.5, 930.9, 1065.0]
                _W = [130.1, 129.8, 130.2, 130.5]
                _Y = [180.0, 210.9, 241.1, 272.5, 302.9]
                _H = [27.3, 26.8, 27.6, 27.3, 26.3]
                button_geometry = {
                    'btnBukaDokpil_PL':   (_X[0], _Y[0], _W[0], _H[0]),
                    'btnRelinkPL':        (_X[1], _Y[0], _W[1], _H[0]),
                    'btnRefreshDataPL':   (_X[2], _Y[0], _W[2], _H[0]),
                    'btnBukaBA_PL':       (_X[0], _Y[1], _W[0], _H[1]),
                    'btnBukaReviu_PL':    (_X[1], _Y[1], _W[1], _H[1]),
                    'btnCetakBAReviu_PL': (_X[0], _Y[2], _W[0], _H[2]),
                    'btnCetakDokpil_PL':  (_X[1], _Y[2], _W[1], _H[2]),
                    'btnCetakReviu_PL':   (_X[0], _Y[3], _W[0], _H[3]),
                    'btnGabungReviu_PL':  (_X[1], _Y[3], _W[1], _H[3]),
                    'btnMuatHPS_PL':      (930.9, 242.5, 130.2, 27.8),
                    'btnIsiEvaluasiPL':   (931.0, 211.5, 130.5, 27.3),
                    'btnCetakBAPLJKK':    (_X[0], _Y[4], _W[0], _H[4]),
                    'btnGabungBAPLJKK':  (_X[1], _Y[4], _W[1], _H[4]),
                    'btnSaveInputData':  (_X[3], _Y[0], _W[3], _H[0]),
                    'btnLoadInputData':  (_X[3], _Y[1], _W[3], _H[1]),
                }
                print('  Layout tombol: PLJKK')

            ba_label = "Cetak BA PLPK" if is_pk else "Cetak BA PLJKK"
            ba_macro = "CetakBAPLPKPDF" if is_pk else "CetakBAPLJKKPDF"
            gabung_label = "Gabung BA PLPK" if is_pk else "Gabung BA PLJKK"
            gabung_macro = "GabungBAPLPK" if is_pk else "GabungBAPLJKK"

            def add_btn(name, label, macro, rgb):
                left, top, width, height = button_geometry[name]
                shp = ws.Shapes.AddShape(5, left, top, width, height)
                shp.Name = name
                r, g, b = rgb
                shp.Fill.ForeColor.RGB = r + (g * 256) + (b * 65536)
                shp.Line.Visible = False
                tf = shp.TextFrame2
                tf.TextRange.Text = label
                tf.TextRange.Font.Fill.ForeColor.RGB = 16777215
                tf.TextRange.Font.Size = 10
                tf.TextRange.Font.Bold = True
                tf.TextRange.ParagraphFormat.Alignment = 2
                tf.VerticalAnchor = 3
                shp.OnAction = macro
                print(f"  [OK] {name} ({label}) -> {macro}")

            RED_DARK   = (180, 0, 0)

            # Baris 0: Buka Dokpil | Relink Word
            # (Muat Paket PL + Isi Data PL dihapus — @ Master Data kini diisi otomatis
            #  via COM saat buat folder di Streamlit, lihat isi_master_data_pl.py)
            add_btn("btnBukaDokpil_PL",   "Buka Dokpil",       "BukaDokpilPlJkk",         TEAL)
            add_btn("btnRelinkPL",        "Relink Word",       "RelinkPL",                 (128, 0, 0))
            add_btn("btnRefreshDataPL",   "Refresh Data PL",   "RefreshDataPL",            (0, 150, 100))
            # Baris 1: Buka BA | Buka Reviu | (kosong) | (kosong)
            add_btn("btnBukaBA_PL",       "Buka BA",           "BukaBAPlJkk",             ORANGE)
            add_btn("btnBukaReviu_PL",    "Buka Reviu",        "BukaReviuPlJkk",          PURPLE)
            # Baris 2: Cetak BA Reviu PL | Cetak Dokpil PDF | (kosong) | (kosong)
            add_btn("btnCetakBAReviu_PL", "Cetak BA Reviu PL", "CetakBAReviuPLPDF",       RED_DARK)
            add_btn("btnCetakDokpil_PL",  "Cetak Dokpil PDF",  "CetakDokpilPlJkkPDF",     (0, 100, 180))
            # Baris 3: Cetak Isi Reviu (kolom 0) | Gabung Reviu (kolom 1) | Muat HPS (kolom 2) | Isi Evaluasi PL (kolom 3)
            add_btn("btnCetakReviu_PL",   "Cetak Isi Reviu",   "CetakReviuPlJkkPDF",       (0, 120, 80))
            add_btn("btnGabungReviu_PL",  "Gabung BA Reviu",   "GabungBAReviu",             (0, 128, 96))
            add_btn("btnMuatHPS_PL",      "Muat HPS",          "MuatHPSPL",                 (200, 100, 0))
            add_btn("btnIsiEvaluasiPL",   "Isi Evaluasi PL",   "IsiEvaluasiPLStandalone",   (160, 60, 0))
            # Baris 4: BA mengikuti konteks workbook, shape legacy tetap
            # dipertahankan agar injector tidak meninggalkan tombol duplikat.
            add_btn("btnCetakBAPLJKK",    ba_label,    ba_macro,    (140, 20, 20))
            add_btn("btnGabungBAPLJKK",   gabung_label, gabung_macro, (100, 20, 80))
            add_btn("btnSaveInputData",    "Save Data", "SaveDataPL", (0, 128, 96))
            add_btn("btnLoadInputData",    "Load Data", "LoadDataPL", (0, 96, 160))

            # Sengaja TIDAK re-protect @ Master Data — user butuh edit bebas
            # (Aturan PL: sheet @ Master Data harus selalu unprotected)

        except Exception as e:
            print(f"  [WARN] Tombol @ Master Data: {e}")

        # Jangan Calculate/CalculateFullRebuild di injector. Ini operasi
        # struktural; memaksa recalc di sini adalah jalur korupsi cache UDF
        # ketika workbook lama memiliki fungsi VBA yang belum ter-load.
        wb.Save()
        print(f"  [SAVED] {os.path.basename(filepath)}")
        return True

    except Exception as e:
        print(f"  [ERROR] {e}")
        return False
    finally:
        os.unlink(tmp_path)
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass
        try:
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


def find_bapljkk_files(root: str) -> list:
    """Cari workbook PL JKK dan PL PK di bawah root.

    Nama fungsi dipertahankan agar pemanggil lama tetap kompatibel.
    """
    files = []
    for current_root, directories, names in os.walk(root):
        directories[:] = [
            directory
            for directory in directories
            if directory.casefold() not in _BACKUP_DIRECTORY_NAMES
            and not directory.startswith(".")
        ]
        for name in names:
            upper_name = name.upper()
            if not upper_name.endswith(".XLSM"):
                continue
            if "BAPLJKK" not in upper_name and "BAPLPK" not in upper_name:
                continue
            candidate = os.path.join(current_root, name)
            if not _is_backup_workbook_path(candidate):
                files.append(candidate)
    return sorted(set(files))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # File spesifik dari argumen
        target = sys.argv[1]
        sys.exit(0 if inject_pl(target) else 1)
    else:
        # Auto-scan POKJA root
        pokja_root = str(SCRIPT_DIR.parent.parent)
        files = find_bapljkk_files(pokja_root)
        if not files:
            print(f"Tidak ada file BAPLJKK*.xlsm ditemukan di: {pokja_root}")
            print("Usage: python inject_pl.py \"path/to/0. BAPLJKK - Template.xlsm\"")
            sys.exit(1)

        print(f"Ditemukan {len(files)} file BAPLJKK:")
        for f in files:
            print(f"  {f}")
        print()

        ok = 0
        for f in files:
            if inject_pl(f):
                ok += 1

        print(f"\n{'='*50}")
        print(f"Selesai: {ok}/{len(files)} file berhasil diinjeksi.")
