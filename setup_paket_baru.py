"""
SETUP PAKET BARU - Otomasi Copy Template + Auto-Link Mail Merge
================================================================
Workflow:
  1. Input nama folder paket baru
  2. Copy 4 file template ke folder baru
  3. Auto-link 3 Word template ke file Excel (via XML, tanpa buka Word)
  4. Siap digunakan!

Cara pakai:
  python setup_paket_baru.py
  python setup_paket_baru.py "19. Pokja 091"
"""
import os
import shutil
import sys
import zipfile
import re
import json
import time
from datetime import datetime

from config import (
    POKJA_ROOT, TEMPLATE_DIR, EXCEL_TEMPLATE, WORD_SHEET_MAP,
    TEMPLATE_DIR_PL, EXCEL_TEMPLATE_PL, WORD_SHEET_MAP_PL,
    OUTPUT_DIR_PL_JKK, OUTPUT_DIR_PL_PK,
    detect_pl_workflow, pl_workflow_config,
    excel_to_file_uri,
)
from pl_snapshot_revision import XML_DATA_SUBFOLDER

# Output base default = root POKJA folder (untuk mode Tender)
OUTPUT_BASE = POKJA_ROOT


def _win_extended_path(path):
    """Gunakan namespace Windows extended bila path melewati MAX_PATH."""
    path = os.path.abspath(os.fspath(path))
    if os.name != "nt" or len(path) < 240 or path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path.lstrip("\\")
    return "\\\\?\\" + path


def _copy2_retry(source, destination, attempts=3, delay=0.75):
    """Copy satu template secara atomik agar race/partial copy dapat diulang."""
    source = os.path.abspath(os.fspath(source))
    destination = os.path.abspath(os.fspath(destination))
    source_io = _win_extended_path(source)
    destination_io = _win_extended_path(destination)
    last_error = None
    for attempt in range(1, attempts + 1):
        # Jangan menambahkan suffix ke destination: nama paket PL dapat sudah
        # dekat batas MAX_PATH. Temp pendek di parent tetap satu volume,
        # sehingga os.replace() masih atomik tanpa memperpanjang path.
        part = os.path.join(
            os.path.dirname(destination),
            f".p{os.getpid()}-{attempt}",
        )
        part_io = _win_extended_path(part)
        try:
            if not os.path.isfile(source_io):
                raise FileNotFoundError(
                    f"Sumber template hilang saat copy: {source}"
                )
            if os.path.exists(part_io):
                os.remove(part_io)
            shutil.copy2(source_io, part_io)
            os.replace(part_io, destination_io)
            return
        except (OSError, shutil.Error) as exc:
            last_error = exc
            try:
                if os.path.exists(part_io):
                    os.remove(part_io)
            except OSError:
                pass
            if attempt < attempts:
                time.sleep(delay * attempt)
    raise OSError(
        f"Gagal copy template setelah {attempts} percobaan: "
        f"{source} -> {destination}; {last_error}"
    ) from last_error


def _quarantine_new_setup(target_dir, output_base):
    """Pindahkan folder setup baru yang gagal ke lokasi recoverable."""
    if not os.path.isdir(_win_extended_path(target_dir)):
        return ""
    quarantine_root = os.path.join(output_base, "_setup-failed")
    os.makedirs(_win_extended_path(quarantine_root), exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = os.path.join(quarantine_root, f"{stamp}-{os.getpid()}")
    suffix = 1
    while os.path.exists(_win_extended_path(candidate)):
        candidate = os.path.join(quarantine_root, f"{stamp}-{os.getpid()}-{suffix}")
        suffix += 1
    os.replace(_win_extended_path(target_dir), _win_extended_path(candidate))
    return candidate


def _write_setup_status(target_dir, status, **extra):
    """Tulis status setup ringan untuk membedakan folder complete/partial."""
    meta_path = os.path.join(target_dir, ".template-meta.json")
    meta_io = _win_extended_path(meta_path)
    if not os.path.isfile(meta_io):
        return
    try:
        with open(meta_io, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        data["setup_status"] = status
        if status == "complete":
            data.pop("failed_at", None)
        elif status == "failed":
            data.pop("completed_at", None)
        data.update(extra)
        temp_path = f"{meta_path}.part-{os.getpid()}"
        temp_io = _win_extended_path(temp_path)
        with open(temp_io, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
        os.replace(temp_io, meta_io)
    except Exception as exc:
        print(f"  [WARN] Status setup gagal ditulis: {exc}")


def link_word_to_excel(word_path, excel_path, sheet_name="data_tender"):
    """Link Word mail merge ke Excel dengan mengedit XML di dalam .docx."""
    word_io_path = _win_extended_path(word_path)
    parent_dir = os.path.dirname(os.path.abspath(os.fspath(word_path)))
    backup = _win_extended_path(os.path.join(parent_dir, f".b{os.getpid()}"))
    temp_path = _win_extended_path(os.path.join(parent_dir, f".t{os.getpid()}"))

    try:
        # Baca settings.xml dan settings.xml.rels dari dalam docx
        with zipfile.ZipFile(word_io_path, 'r') as zf:
            if 'word/settings.xml' not in zf.namelist():
                return False
            settings_xml = zf.read('word/settings.xml')
            _ = zf.namelist()  # preload for later iteration

        settings_str = settings_xml.decode('utf-8')

        # Hapus mailMerge element yang ada (jika ada)
        settings_str = re.sub(r'<w:mailMerge>.*?</w:mailMerge>', '', settings_str, flags=re.DOTALL)

        # Build mailMerge XML baru
        # Escape & dan " untuk XML attribute value
        excel_escaped = excel_path.replace('&', '&amp;')
        mail_merge = (
            '<w:mailMerge>'
            '<w:mainDocumentType w:val="formLetters"/>'
            '<w:linkToQuery/>'
            '<w:dataType w:val="native"/>'
            '<w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;'
            'User ID=Admin;'
            f'Data Source={excel_escaped};'
            'Mode=Read;'
            'Extended Properties=&quot;HDR=YES;IMEX=1&quot;;"/>'
            f'<w:query w:val="SELECT * FROM `{sheet_name}$`"/>'
            '</w:mailMerge>'
        )

        # Sisipkan sebelum </w:settings>
        settings_str = settings_str.replace('</w:settings>', mail_merge + '</w:settings>')

        new_settings = settings_str.encode('utf-8')

        # Build settings.xml.rels dengan path Excel yang benar
        excel_uri = excel_to_file_uri(excel_path)
        new_settings_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_uri}" TargetMode="External"/>'
            f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_uri}" TargetMode="External"/>'
            '</Relationships>'
        ).encode('utf-8')

        # Repack docx - replace settings.xml + settings.xml.rels
        shutil.copy2(word_io_path, backup)

        with zipfile.ZipFile(backup, 'r') as zf_in:
            has_settings_rels = 'word/_rels/settings.xml.rels' in zf_in.namelist()
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == 'word/settings.xml':
                        zf_out.writestr(item, new_settings)
                    elif item.filename == 'word/_rels/settings.xml.rels':
                        zf_out.writestr(item, new_settings_rels)
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))
                # Buat settings.xml.rels kalau belum ada
                if not has_settings_rels:
                    zf_out.writestr('word/_rels/settings.xml.rels', new_settings_rels)

        # Replace original dengan file baru
        os.replace(temp_path, word_io_path)
        os.remove(backup)
        return True

    except Exception as e:
        print(f"      Error: {e}")
        if os.path.exists(backup):
            shutil.copy2(backup, word_io_path)
            os.remove(backup)
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False


def _nama_output_template(nama_template, suffix):
    """Ganti label Template + domain dengan nama paket tanpa duplikasi."""
    if not suffix or "Template" not in nama_template:
        return nama_template
    # V2: ``Template Perencanaan/Pengawasan`` menjadi suffix paket saja.
    return re.sub(
        r"Template(?:\s+(?:Perencanaan|Pengawasan|Konstruksi))?",
        suffix,
        nama_template,
        count=1,
        flags=re.IGNORECASE,
    )


def _setup_folder(folder_name, template_dir, excel_template, word_sheet_map, output_base=None, workflow=None):
    """Inti setup: copy template + auto-link mail merge ke folder baru."""
    print("=" * 60)
    print("  SETUP PAKET BARU")
    print("  Otomasi Copy Template + Auto-Link Mail Merge")
    print("=" * 60)

    if not folder_name:
        print("[BATAL] Nama folder kosong.")
        return

    template_dir = os.path.abspath(os.fspath(template_dir))
    required_sources = [excel_template] + [wf_tpl for wf_tpl, _sheet in word_sheet_map]
    missing_sources = [
        os.path.join(template_dir, name)
        for name in required_sources
        if not os.path.isfile(os.path.join(template_dir, name))
    ]
    if missing_sources:
        details = "\n".join(f"  - {path}" for path in missing_sources)
        raise FileNotFoundError(
            "Template workflow tidak lengkap; setup dibatalkan sebelum folder "
            f"dibuat.\n{details}"
        )

    base = output_base or OUTPUT_BASE
    folder_name = re.sub(r'[<>:"/\\|?*]', '-', folder_name).strip()
    target_dir = os.path.join(base, folder_name)
    created_target = False

    if os.path.exists(_win_extended_path(target_dir)):
        print(f"\n[WARN] Folder '{folder_name}' sudah ada!")
        files_exist = os.listdir(_win_extended_path(target_dir))
        if files_exist:
            print(f"  Isi: {len(files_exist)} file")
            for f in files_exist[:5]:
                print(f"    - {f}")
        if sys.stdin.isatty():
            try:
                jawab = input("\nLanjutkan? File yang sudah ada tidak akan di-overwrite. (y/n): ").strip().lower()
            except EOFError:
                jawab = 'y'
            if jawab != 'y':
                print("[BATAL]")
                return
        else:
            print("[AUTO] Non-interaktif — lanjut, file existing tidak di-overwrite.")
    else:
        os.makedirs(_win_extended_path(target_dir))
        created_target = True

    # Auto-create subfolder (untuk semua tipe paket, baru maupun existing)
    # Mode Tender: subfolder identik dengan PL JKK, kecuali no.9
    is_tender = bool(re.search(r"Pokja\s*[-]?\s*\d+", folder_name, re.IGNORECASE))
    _subfolder_9 = "9. Dokumen Penawaran Teknis & Biaya" if is_tender else "9. Dokumen Teknis Biaya"
    _subfolders = [
        "0. Draft Dokumen PPK",
        "1. KAK & Spesifikasi Teknis",
        "2. Rancangan Kontrak",
        "3. Uraian Singkat Pekerjaan",
        "4. Informasi Lainnya",
        "5. Evaluator Kualifikasi & Teknis",
        "6. BA Reviu Lengkap",
        "7. Berita Acara + Summary Non Tender",
        "8. Dokumen Kualifikasi",
        _subfolder_9,
    ]
    if workflow:
        _subfolders.extend(["10. Revisi Uploadan PPK", XML_DATA_SUBFOLDER])
    for _sub in _subfolders:
        os.makedirs(_win_extended_path(os.path.join(target_dir, _sub)), exist_ok=True)

    # Metadata ringan untuk audit/refresh otomatis; user tidak perlu mengisi.
    try:
        meta_path = os.path.join(target_dir, ".template-meta.json")
        if not os.path.exists(_win_extended_path(meta_path)):
            with open(_win_extended_path(meta_path), "w", encoding="utf-8") as _mf:
                json.dump({
                    "schema": 1,
                    "workflow": workflow or "legacy",
                    "template_dir": os.path.abspath(template_dir),
                    "dynamic_header": True,
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "setup_status": "in_progress",
                }, _mf, ensure_ascii=False, indent=2)
    except Exception as _meta_e:
        print(f"  [WARN] Metadata template gagal: {_meta_e}")


    # Extract suffix untuk rename Excel
    # Tender:  "Pokja 086" → "086"
    # PL JKK:  "1. PLJKK - Nama Paket" → "Nama Paket"
    # PL PK:   "1. PLPK - Nama Paket"  → "Nama Paket"
    pokja_suffix = ""
    m_pokja = re.search(r"Pokja\s*[-]?\s*(\d+)", folder_name, re.IGNORECASE)
    if m_pokja:
        pokja_suffix = m_pokja.group(1)
    else:
        # Mode PL: ambil bagian setelah "PLJKK - " atau "PLPK - "
        m_pl = re.search(r"PL(?:JKK|PK)\s+-\s+(.+)$", folder_name, re.IGNORECASE)
        if m_pl:
            pokja_suffix = m_pl.group(1).strip()

    print(f"\n[1/3] Folder: {target_dir}")

    # 1. Excel — rename "Template" → suffix
    excel_name_dst = _nama_output_template(excel_template, pokja_suffix)
    dst_excel = os.path.join(target_dir, excel_name_dst)
    excel_created = False

    def _copy_template(source, destination):
        try:
            _copy2_retry(source, destination)
        except Exception:
            _write_setup_status(
                target_dir,
                "failed",
                failed_at=datetime.now().isoformat(timespec="seconds"),
            )
            if created_target:
                try:
                    quarantined = _quarantine_new_setup(target_dir, base)
                    print(f"  [RECOVERABLE] Folder partial dipindah ke: {quarantined}")
                except Exception as quarantine_error:
                    print(f"  [WARN] Quarantine folder partial gagal: {quarantine_error}")
            raise

    print("\n[2/3] Copy & Rename template files...")
    if not os.path.exists(_win_extended_path(dst_excel)):
        _copy_template(os.path.join(template_dir, excel_template), dst_excel)
        excel_created = True
        print(f"  [OK] {excel_template} -> {excel_name_dst}")
    else:
        print(f"  [SKIP] {excel_name_dst} (sudah ada)")

    # 2. Word files — rename "Template" → suffix jika ada kata "Template" di nama file
    dst_word_map = []
    for wf_tpl, sheet_name in word_sheet_map:
        if pokja_suffix and "Template" in wf_tpl:
            wf_dst = _nama_output_template(wf_tpl, pokja_suffix)
        else:
            wf_dst = wf_tpl
        dst_path = os.path.join(target_dir, wf_dst)
        if not os.path.exists(_win_extended_path(dst_path)):
            _copy_template(os.path.join(template_dir, wf_tpl), dst_path)
            # V2 donor masih membawa header PUPR sebagai placeholder historis.
            # Salinan paket harus netral; profile instansi dipasang saat export.
            from document_profiles import is_official_header_document, strip_static_headers

            if is_official_header_document(dst_path):
                strip_static_headers(dst_path)
            print(f"  [OK] {wf_tpl} -> {wf_dst}")
        else:
            print(f"  [SKIP] {wf_dst} (sudah ada)")
        dst_word_map.append((dst_path, wf_dst, sheet_name))

    # 3. Auto-link Word → Excel
    print("\n[3/3] Auto-link Word Mail Merge -> Excel...")
    abs_excel = os.path.abspath(dst_excel)
    success_count = 0
    for dst_path, wf_dst, sheet_name in dst_word_map:
        ok = link_word_to_excel(dst_path, abs_excel, sheet_name)
        if ok:
            print(f"  [OK] {wf_dst} -> sheet '{sheet_name}'")
            success_count += 1
        else:
            print(f"  [FAIL] {wf_dst}")
    if success_count != len(dst_word_map):
        _write_setup_status(
            target_dir,
            "failed",
            failed_at=datetime.now().isoformat(timespec="seconds"),
        )
        if created_target:
            try:
                quarantined = _quarantine_new_setup(target_dir, base)
                print(f"  [RECOVERABLE] Folder partial dipindah ke: {quarantined}")
            except Exception as quarantine_error:
                print(f"  [WARN] Quarantine folder partial gagal: {quarantine_error}")
        raise RuntimeError(
            f"Auto-link mail merge tidak lengkap: "
            f"{success_count}/{len(dst_word_map)}"
        )

    # Scrub data donor setelah template disalin dan mail merge terhubung.
    # PL memakai range berbeda untuk JKK vs konstruksi; jangan bawa data contoh.
    if is_tender and excel_created:
        try:
            from template_scrub import scrub_package_copy

            scrub_log = scrub_package_copy(
                target_dir,
                dst_excel,
                [item[0] for item in dst_word_map],
            )
            print("\n[4/4] Scrub data donor...")
            for line in scrub_log:
                print(f"  {line}")
        except Exception as exc:
            print(f"  [WARN] Scrub template gagal: {exc}")
    elif is_tender:
        print("\n[4/4] Scrub data donor dilewati — workbook existing dipertahankan.")
    elif excel_created and workflow:
        try:
            from template_scrub import scrub_excel_pl_copy

            is_pk = "KONSTRUKSI" in str(workflow).upper() or "PLPK" in folder_name.upper()
            scrub_log = scrub_excel_pl_copy(dst_excel, is_pk=is_pk)
            print("\n[4/4] Scrub data donor PL...")
            for line in scrub_log:
                print(f"  {line}")
        except Exception as exc:
            print(f"  [WARN] Scrub template PL gagal: {exc}")

    print(f"\n{'='*60}")
    print(f"  SETUP SELESAI!")
    print(f"{'='*60}")
    print(f"\n  Folder : {target_dir}")
    print(f"  Excel  : {excel_name_dst}")
    print(f"  Word   : {success_count}/{len(word_sheet_map)} terhubung")
    _write_setup_status(
        target_dir,
        "complete",
        completed_at=datetime.now().isoformat(timespec="seconds"),
    )


def setup_paket_baru(folder_name=None, output_base=None):
    """Setup paket baru mode Tender (PK): copy template + auto-link mail merge."""
    if not folder_name:
        print("\nContoh: '19. Pokja 091'")
        folder_name = input("Nama folder paket baru: ").strip()
    _setup_folder(folder_name, TEMPLATE_DIR, EXCEL_TEMPLATE, WORD_SHEET_MAP, output_base=output_base)


def setup_paket_baru_pl(folder_name=None, output_base=None, template_dir=None, workflow=None):
    """Setup paket baru mode Pengadaan Langsung (PL): copy template BAPLJKK + auto-link.
    output_base: override folder tujuan (default: deteksi dari nama PLJKK/PLPK).
    """
    if not folder_name:
        print("\nContoh: '1. PLJKK - Perencanaan Pembangunan Jalan ...'")
        folder_name = input("Nama folder paket PL: ").strip()

    workflow = workflow or detect_pl_workflow({"nama_paket": folder_name})
    workflow_cfg = pl_workflow_config(workflow)

    # Prefix folder adalah kontrak family dari UI. Jangan pernah membuat
    # workbook PLPK di folder PLJKK, atau sebaliknya.
    _folder_upper = folder_name.upper()
    _folder_family = "PK" if re.search(r"\bPLPK\b", _folder_upper) else (
        "JKK" if re.search(r"\bPLJKK\b", _folder_upper) else workflow_cfg["jenis_pl"]
    )
    if workflow_cfg["jenis_pl"] != _folder_family:
        raise ValueError(
            f"Workflow {workflow} ({workflow_cfg['jenis_pl']}) tidak cocok "
            f"dengan family folder {_folder_family}. Proses dibatalkan sebelum copy."
        )

    if output_base is None:
        # Deteksi dari nama folder: PLPK → PK dir, default → JKK dir
        if workflow_cfg["jenis_pl"] == "PK":
            output_base = OUTPUT_DIR_PL_PK
        else:
            output_base = OUTPUT_DIR_PL_JKK

    resolved_template_dir = template_dir
    from config import _pl_template_set_complete, pl_workflow_template_dir
    if resolved_template_dir:
        resolved_template_dir = os.path.abspath(os.fspath(resolved_template_dir))
    if not resolved_template_dir or not _pl_template_set_complete(resolved_template_dir, workflow_cfg):
        fallback_template_dir = pl_workflow_template_dir(workflow)
        if resolved_template_dir and os.path.normcase(resolved_template_dir) != os.path.normcase(fallback_template_dir):
            print(
                "  [WARN] Template eksplisit tidak lengkap; memakai donor lengkap: "
                f"{fallback_template_dir}"
            )
        resolved_template_dir = fallback_template_dir
    _setup_folder(
        folder_name,
        resolved_template_dir,
        workflow_cfg["excel_template"],
        workflow_cfg["word_map"],
        output_base,
        workflow=workflow,
    )


if __name__ == "__main__":
    args = sys.argv[1:]
    mode_pl = "--mode" in args and args[args.index("--mode") + 1] == "pl" if "--mode" in args else False

    # --output-dir <path> → override output_base
    output_dir = None
    if "--output-dir" in args:
        idx = args.index("--output-dir")
        output_dir = args[idx + 1]
        args = args[:idx] + args[idx + 2:]

    template_dir = None
    if "--template-dir" in args:
        idx = args.index("--template-dir")
        template_dir = args[idx + 1]
        args = args[:idx] + args[idx + 2:]

    workflow = None
    if "--workflow" in args:
        idx = args.index("--workflow")
        workflow = args[idx + 1]
        args = args[:idx] + args[idx + 2:]

    name_args = [a for a in args if not a.startswith("--") and a != "pl"]
    folder_name = " ".join(name_args).strip() or None

    if mode_pl:
        setup_paket_baru_pl(folder_name, output_base=output_dir, template_dir=template_dir, workflow=workflow)
    else:
        setup_paket_baru(folder_name, output_base=output_dir)
