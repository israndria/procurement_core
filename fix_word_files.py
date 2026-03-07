"""
Fix corrupted Word files by removing embedded mail merge settings.
The mail merge connection will be handled at runtime by word_merge.py instead.
"""
import os
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET

def fix_word_file(word_path):
    """Remove mail merge element from Word settings.xml to fix corruption."""
    print(f"  Fixing: {os.path.basename(word_path)}")
    
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # Register all common namespaces to avoid ns0, ns1 prefixes
    namespaces = {
        'w': ns_w,
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'sl': 'http://schemas.openxmlformats.org/schemaLibrary/2006/main',
        'o': 'urn:schemas-microsoft-com:office:office',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
        'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
        'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Extract docx
        with zipfile.ZipFile(word_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            print(f"    No settings.xml found, skipping")
            return False
        
        # Parse and remove mailMerge element
        tree = ET.parse(settings_path)
        root = tree.getroot()
        
        ns = f'{{{ns_w}}}'
        removed = False
        for mm in root.findall(f'{ns}mailMerge'):
            root.remove(mm)
            removed = True
        
        if not removed:
            print(f"    No mailMerge element found, file is clean")
            return True
        
        # Save settings.xml
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8')
        
        # Repack docx - PROPERLY preserve the original zip structure
        backup = word_path + ".bak"
        shutil.copy2(word_path, backup)
        
        # Read original zip to get compression info
        original_names = []
        with zipfile.ZipFile(backup, 'r') as zf:
            original_names = zf.namelist()
        
        with zipfile.ZipFile(word_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    # Use forward slashes (zip standard)
                    arcname = arcname.replace('\\', '/')
                    zf.write(file_path, arcname)
        
        os.remove(backup)
        print(f"    [OK] mailMerge removed, file restored")
        return True
        
    except Exception as e:
        print(f"    [ERROR] {e}")
        backup = word_path + ".bak"
        if os.path.exists(backup):
            shutil.copy2(backup, word_path)
            os.remove(backup)
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    folder = r"D:\Dokumen\@ POKJA 2026\Paket Experiment"
    
    word_files = [
        "1. Full Dokumen BA PK v.1.docx",
        "2. Isi Reviu PK v.1.docx",
        "3. Dokpil Full PK v.1.docx",
    ]
    
    print("Fixing Word files (removing embedded mail merge)...")
    for wf in word_files:
        wp = os.path.join(folder, wf)
        if os.path.exists(wp):
            fix_word_file(wp)
        else:
            print(f"  [SKIP] {wf} not found")
    
    print("\nDone! Word files should now open without corruption errors.")
    print("Mail merge connection will be handled at runtime by word_merge.py")
