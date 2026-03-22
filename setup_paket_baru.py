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
    
    target_dir = os.path.join(OUTPUT_BASE, folder_name)
    
    # Cek folder existing
    if os.path.exists(target_dir):
        print(f"\n[WARN] Folder '{folder_name}' sudah ada!")
        files_exist = os.listdir(target_dir)
        if files_exist:
            print(f"  Isi: {len(files_exist)} file")
            for f in files_exist[:5]:
                print(f"    - {f}")
        jawab = input("\nLanjutkan? File yang sudah ada tidak akan di-overwrite. (y/n): ").strip().lower()
        if jawab != 'y':
            print("[BATAL]")
            return
    else:
        os.makedirs(target_dir)
    
    print(f"\n[1/3] Folder: {target_dir}")
    
    # Copy files
    print("\n[2/3] Copy template files...")
    
    # Excel
    dst_excel = os.path.join(target_dir, EXCEL_TEMPLATE)
    if not os.path.exists(dst_excel):
        shutil.copy2(os.path.join(TEMPLATE_DIR, EXCEL_TEMPLATE), dst_excel)
        print(f"  [OK] {EXCEL_TEMPLATE}")
    else:
        print(f"  [SKIP] {EXCEL_TEMPLATE} (sudah ada)")
    
    # Word files
    for wf, _ in WORD_SHEET_MAP:
        dst = os.path.join(target_dir, wf)
        if not os.path.exists(dst):
            shutil.copy2(os.path.join(TEMPLATE_DIR, wf), dst)
            print(f"  [OK] {wf}")
        else:
            print(f"  [SKIP] {wf} (sudah ada)")
    
    # Auto-link setiap Word ke sheet yang benar
    print("\n[3/3] Auto-link Word Mail Merge -> Excel...")
    abs_excel = os.path.abspath(dst_excel)
    
    success_count = 0
    for wf, sheet_name in WORD_SHEET_MAP:
        dst_word = os.path.join(target_dir, wf)
        ok = link_word_to_excel(dst_word, abs_excel, sheet_name)
        if ok:
            print(f"  [OK] {wf} -> sheet '{sheet_name}'")
            success_count += 1
        else:
            print(f"  [FAIL] {wf}")
    
    # Summary
    print(f"\n{'='*60}")
    print(f"  SETUP SELESAI!")
    print(f"{'='*60}")
    print(f"\n  Folder : {target_dir}")
    print(f"  Excel  : {EXCEL_TEMPLATE}")
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
