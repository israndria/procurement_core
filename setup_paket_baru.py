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

from config import (
    POKJA_ROOT, TEMPLATE_DIR, EXCEL_TEMPLATE, WORD_SHEET_MAP,
    excel_to_file_uri,
)

# Output base = root POKJA folder
OUTPUT_BASE = POKJA_ROOT


def link_word_to_excel(word_path, excel_path, sheet_name="data_tender"):
    """Link Word mail merge ke Excel dengan mengedit XML di dalam .docx."""
    

    try:
        # Baca settings.xml dan settings.xml.rels dari dalam docx
        with zipfile.ZipFile(word_path, 'r') as zf:
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
        backup = word_path + ".bak"
        shutil.copy2(word_path, backup)

        temp_path = word_path + ".tmp"
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
        os.replace(temp_path, word_path)
        os.remove(backup)
        return True

    except Exception as e:
        print(f"      Error: {e}")
        backup = word_path + ".bak"
        if os.path.exists(backup):
            shutil.copy2(backup, word_path)
            os.remove(backup)
        temp_path = word_path + ".tmp"
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False


def setup_paket_baru(folder_name=None):
    """Setup paket baru: copy template + auto-link mail merge."""
    
    print("=" * 60)
    print("  SETUP PAKET BARU")
    print("  Otomasi Copy Template + Auto-Link Mail Merge")
    print("=" * 60)
    
    # Input nama folder
    if not folder_name:
        print("\nContoh: '19. Pokja 091'")
        folder_name = input("Nama folder paket baru: ").strip()
    
    if not folder_name:
        print("[BATAL] Nama folder kosong.")
        return

    # Sanitize: ganti karakter tidak valid Windows dengan "-"
    folder_name = re.sub(r'[<>:"/\\|?*]', '-', folder_name).strip()

    target_dir = os.path.join(OUTPUT_BASE, folder_name)
    
    # Cek folder existing
    if os.path.exists(target_dir):
        print(f"\n[WARN] Folder '{folder_name}' sudah ada!")
        files_exist = os.listdir(target_dir)
        if files_exist:
            print(f"  Isi: {len(files_exist)} file")
            for f in files_exist[:5]:
                print(f"    - {f}")
        import sys
        if sys.stdin.isatty():
            jawab = input("\nLanjutkan? File yang sudah ada tidak akan di-overwrite. (y/n): ").strip().lower()
            if jawab != 'y':
                print("[BATAL]")
                return
        else:
            print("[AUTO] Non-interaktif — lanjut, file existing tidak di-overwrite.")
    else:
        os.makedirs(target_dir)
    
    # Extract Pokja Number untuk suffix (misal: "086")
    pokja_suffix = ""
    m_pokja = re.search(r"Pokja\s+(\d+)", folder_name, re.IGNORECASE)
    if m_pokja:
        pokja_suffix = m_pokja.group(1)
    
    print(f"\n[1/3] Folder: {target_dir}")
    if pokja_suffix:
        print(f"      Pokja detected: {pokja_suffix}")
    
    # Copy files with renaming
    print("\n[2/3] Copy & Rename template files...")
    
    # 1. Excel
    excel_name_dst = EXCEL_TEMPLATE.replace("Template", pokja_suffix) if pokja_suffix else EXCEL_TEMPLATE
    dst_excel = os.path.join(target_dir, excel_name_dst)
    
    if not os.path.exists(dst_excel):
        shutil.copy2(os.path.join(TEMPLATE_DIR, EXCEL_TEMPLATE), dst_excel)
        print(f"  [OK] {EXCEL_TEMPLATE} -> {excel_name_dst}")
    else:
        print(f"  [SKIP] {excel_name_dst} (sudah ada)")
    
    # 2. Word files
    dst_word_map = [] # list of (new_word_path, sheet_name)
    for wf_tpl, sheet_name in WORD_SHEET_MAP:
        wf_dst = wf_tpl.replace("Template", pokja_suffix) if pokja_suffix else wf_tpl
        dst_path = os.path.join(target_dir, wf_dst)
        
        # Copy Template Asli
        if not os.path.exists(dst_path):
            shutil.copy2(os.path.join(TEMPLATE_DIR, wf_tpl), dst_path)
            print(f"  [OK] {wf_tpl} -> {wf_dst}")
        else:
            print(f"  [SKIP] {wf_dst} (sudah ada)")
            
        dst_word_map.append((dst_path, wf_dst, sheet_name))
    
    # Auto-link setiap Word ke sheet yang benar
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
    
    # Summary
    print(f"\n{'='*60}")
    print(f"  SETUP SELESAI!")
    print(f"{'='*60}")
    print(f"\n  Folder : {target_dir}")
    print(f"  Excel  : {excel_name_dst}")
    print(f"  Word   : {success_count}/{len(WORD_SHEET_MAP)} terhubung")
    print(f"\n  Langkah selanjutnya:")
    print(f"  1. Buka Excel -> isi data paket baru")
    print(f"  2. Buka Word  -> Mailings -> Preview Results")
    print(f"  3. Finish & Merge -> Print/PDF")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        name = " ".join(sys.argv[1:])
        setup_paket_baru(name)
    else:
        setup_paket_baru()
