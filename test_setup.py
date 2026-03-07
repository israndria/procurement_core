"""
Setup Paket Baru v3 - Direct XML approach
Bypass Word COM OpenDataSource yang selalu hang.
Gunakan python-docx + XML manipulation untuk set mail merge connection.
"""
import os
import shutil
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from copy import deepcopy

TEMPLATE_DIR = r"D:\Dokumen\@ POKJA 2026\Paket Experiment"
OUTPUT_BASE = r"D:\Dokumen\@ POKJA 2026"

EXCEL_TEMPLATE = "@ BA PK 2026 (Improved) v.1.xlsm"
WORD_TEMPLATES = [
    "1. Full Dokumen BA PK v.1.docx",
    "2. Isi Reviu PK v.1.docx",
    "3. Dokpil Full PK v.1.docx",
]

# Word OOXML namespaces
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'o': 'urn:schemas-microsoft-com:office:office',
}

def update_mail_merge_in_settings(settings_xml, excel_path, sheet_name="data_tender"):
    """Update or insert mail merge settings in word/settings.xml."""
    
    # Register namespaces
    for prefix, uri in NSMAP.items():
        ET.register_namespace(prefix, uri)
    # Register additional namespaces commonly in settings.xml
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    ET.register_namespace('sl', 'http://schemas.openxmlformats.org/schemaLibrary/2006/main')
    ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
    
    tree = ET.parse(settings_xml)
    root = tree.getroot()
    
    # Remove existing mailMerge element if present
    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    for mm in root.findall(f'{ns_w}mailMerge'):
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
    mm_elem = ET.SubElement(root, f'{ns_w}mailMerge')
    
    # mainDocumentType
    mdt = ET.SubElement(mm_elem, f'{ns_w}mainDocumentType')
    mdt.set(f'{ns_w}val', 'formLetters')
    
    # linkToQuery
    ltq = ET.SubElement(mm_elem, f'{ns_w}linkToQuery')
    
    # dataType
    dt = ET.SubElement(mm_elem, f'{ns_w}dataType')
    dt.set(f'{ns_w}val', 'native')
    
    # connectString
    cs = ET.SubElement(mm_elem, f'{ns_w}connectString')
    cs.set(f'{ns_w}val', conn)
    
    # query
    q = ET.SubElement(mm_elem, f'{ns_w}query')
    q.set(f'{ns_w}val', query)
    
    tree.write(settings_xml, xml_declaration=True, encoding='UTF-8')
    return True


def link_word_to_excel(word_path, excel_path, sheet_name="data_tender"):
    """Link Word mail merge to Excel by editing the docx XML directly."""
    
    # docx is a zip file
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Extract
        with zipfile.ZipFile(word_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        # Find and update settings.xml
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            print(f"      [WARN] settings.xml not found!")
            return False
        
        # Update mail merge settings
        update_mail_merge_in_settings(settings_path, excel_path, sheet_name)
        
        # Repack into docx
        # First backup original
        backup = word_path + ".bak"
        if os.path.exists(backup):
            os.remove(backup)
        shutil.copy2(word_path, backup)
        
        # Create new zip
        with zipfile.ZipFile(word_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arcname)
        
        # Remove backup if success
        os.remove(backup)
        return True
        
    except Exception as e:
        print(f"      [ERROR] {e}")
        # Restore from backup
        backup = word_path + ".bak"
        if os.path.exists(backup):
            shutil.copy2(backup, word_path)
            os.remove(backup)
        return False
    finally:
        # Cleanup temp
        shutil.rmtree(temp_dir, ignore_errors=True)


def setup_paket(folder_name):
    """Setup paket baru: copy + link."""
    target_dir = os.path.join(OUTPUT_BASE, folder_name)
    
    print(f"\n{'='*60}")
    print(f"  SETUP: {folder_name}")
    print(f"{'='*60}")
    
    # Create folder
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
    print(f"\n[1] Folder: {target_dir}")
    
    # Copy files
    print("\n[2] Copy template files...")
    dst_excel = os.path.join(target_dir, EXCEL_TEMPLATE)
    if not os.path.exists(dst_excel):
        shutil.copy2(os.path.join(TEMPLATE_DIR, EXCEL_TEMPLATE), dst_excel)
    print(f"    [OK] {EXCEL_TEMPLATE}")
    
    dst_words = []
    for wf in WORD_TEMPLATES:
        dst = os.path.join(target_dir, wf)
        if not os.path.exists(dst):
            shutil.copy2(os.path.join(TEMPLATE_DIR, wf), dst)
        dst_words.append(dst)
        print(f"    [OK] {wf}")
    
    # Auto-link via XML
    print("\n[3] Auto-link Word -> Excel (via XML)...")
    abs_excel = os.path.abspath(dst_excel)
    
    for dst_word in dst_words:
        fname = os.path.basename(dst_word)
        print(f"\n    {fname}")
        
        success = link_word_to_excel(dst_word, abs_excel)
        if success:
            print(f"      [OK] Linked to {EXCEL_TEMPLATE}")
        else:
            print(f"      [FAIL]")
    
    # Verify
    print(f"\n[4] Isi folder:")
    for f in os.listdir(target_dir):
        size = os.path.getsize(os.path.join(target_dir, f))
        print(f"    {f} ({size:,} bytes)")
    
    print(f"\n{'='*60}")
    print("  SELESAI!")
    print(f"{'='*60}")
    print(f"\nBuka Word -> Mailings -> sudah terhubung ke Excel!")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        name = " ".join(sys.argv[1:])
    else:
        name = "TEST Pokja 997"
    
    setup_paket(name)
