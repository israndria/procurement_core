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
import tempfile
import xml.etree.ElementTree as ET

# ===== KONFIGURASI =====
TEMPLATE_DIR = r"D:\Dokumen\@ POKJA 2026\Paket Experiment"
OUTPUT_BASE = r"D:\Dokumen\@ POKJA 2026"

EXCEL_TEMPLATE = "@ BA PK 2026 (Improved) v.1.xlsm"

# Setiap Word template terhubung ke sheet Excel BERBEDA
WORD_SHEET_MAP = [
    ("1. Full Dokumen BA PK v.1.docx", "1. Input Data"),
    ("2. Isi Reviu PK v.1.docx",       "database_reviu"),
    ("3. Dokpil Full PK v.1.docx",     "database_dokpil"),
]
# ========================


def link_word_to_excel(word_path, excel_path, sheet_name="data_tender"):
    """Link Word mail merge ke Excel dengan mengedit XML di dalam .docx."""
    
    # Register XML namespaces
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ET.register_namespace('w', ns_w)
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    ET.register_namespace('sl', 'http://schemas.openxmlformats.org/schemaLibrary/2006/main')
    ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office')
    ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
    ET.register_namespace('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex')
    ET.register_namespace('w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid')
    ET.register_namespace('w16', 'http://schemas.microsoft.com/office/word/2018/wordml')
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Extract docx (it's a zip)
        with zipfile.ZipFile(word_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            return False
        
        # Parse settings.xml
        tree = ET.parse(settings_path)
        root = tree.getroot()
        
        ns = f'{{{ns_w}}}'
        
        # Remove existing mailMerge element
        for mm in root.findall(f'{ns}mailMerge'):
            root.remove(mm)
        
        # Build connection string
        conn = (
            f"Provider=Microsoft.ACE.OLEDB.12.0;"
            f"User ID=Admin;"
            f"Data Source={excel_path};"
            f"Mode=Read;"
            f'Extended Properties="HDR=YES;IMEX=1";'
        )
        query = f"SELECT * FROM `{sheet_name}$`"
        
        # Create mailMerge element
        mm_elem = ET.SubElement(root, f'{ns}mailMerge')
        
        mdt = ET.SubElement(mm_elem, f'{ns}mainDocumentType')
        mdt.set(f'{ns}val', 'formLetters')
        
        ET.SubElement(mm_elem, f'{ns}linkToQuery')
        
        dt = ET.SubElement(mm_elem, f'{ns}dataType')
        dt.set(f'{ns}val', 'native')
        
        cs = ET.SubElement(mm_elem, f'{ns}connectString')
        cs.set(f'{ns}val', conn)
        
        q = ET.SubElement(mm_elem, f'{ns}query')
        q.set(f'{ns}val', query)
        
        # Save settings.xml
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8')
        
        # Repack docx
        backup = word_path + ".bak"
        shutil.copy2(word_path, backup)
        
        with zipfile.ZipFile(word_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arcname)
        
        os.remove(backup)
        return True
        
    except Exception as e:
        print(f"      Error: {e}")
        backup = word_path + ".bak"
        if os.path.exists(backup):
            shutil.copy2(backup, word_path)
            os.remove(backup)
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


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
