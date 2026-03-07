"""
Link semua Word template di folder tertentu ke Excel di folder yang sama.
Dipanggil dari VBA via Shell atau langsung dari command line.
Usage: python link_word_here.py "D:\path\to\folder"
"""
import os
import sys
import shutil
import zipfile
import tempfile
import xml.etree.ElementTree as ET

EXCEL_NAME = "@ BA PK 2026 (Improved) v.1.xlsm"
WORD_SHEET_MAP = [
    ("1. Full Dokumen BA PK v.1.docx", "1. Input Data"),
    ("2. Isi Reviu PK v.1.docx",       "database_reviu"),
    ("3. Dokpil Full PK v.1.docx",     "database_dokpil"),
]

def link_word_to_excel(word_path, excel_path, sheet_name="data_tender"):
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for prefix, uri in [
        ('w', ns_w),
        ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
        ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
        ('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'),
        ('sl', 'http://schemas.openxmlformats.org/schemaLibrary/2006/main'),
        ('o', 'urn:schemas-microsoft-com:office:office'),
        ('w14', 'http://schemas.microsoft.com/office/word/2010/wordml'),
        ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
        ('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex'),
        ('w16cid', 'http://schemas.microsoft.com/office/word/2016/wordml/cid'),
        ('w16', 'http://schemas.microsoft.com/office/word/2018/wordml'),
    ]:
        ET.register_namespace(prefix, uri)
    
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(word_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            return False
        
        tree = ET.parse(settings_path)
        root = tree.getroot()
        ns = f'{{{ns_w}}}'
        
        for mm in root.findall(f'{ns}mailMerge'):
            root.remove(mm)
        
        conn = (f"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;"
                f"Data Source={excel_path};Mode=Read;"
                f'Extended Properties="HDR=YES;IMEX=1";')
        
        mm_elem = ET.SubElement(root, f'{ns}mailMerge')
        mdt = ET.SubElement(mm_elem, f'{ns}mainDocumentType')
        mdt.set(f'{ns}val', 'formLetters')
        ET.SubElement(mm_elem, f'{ns}linkToQuery')
        dt = ET.SubElement(mm_elem, f'{ns}dataType')
        dt.set(f'{ns}val', 'native')
        cs = ET.SubElement(mm_elem, f'{ns}connectString')
        cs.set(f'{ns}val', conn)
        q = ET.SubElement(mm_elem, f'{ns}query')
        q.set(f'{ns}val', f"SELECT * FROM `{sheet_name}$`")
        
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8')
        
        backup = word_path + ".bak"
        shutil.copy2(word_path, backup)
        with zipfile.ZipFile(word_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for rd, dirs, files in os.walk(temp_dir):
                for f in files:
                    fp = os.path.join(rd, f)
                    zf.write(fp, os.path.relpath(fp, temp_dir))
        os.remove(backup)
        return True
    except:
        backup = word_path + ".bak"
        if os.path.exists(backup):
            shutil.copy2(backup, word_path)
            os.remove(backup)
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    excel_path = os.path.join(folder, EXCEL_NAME)
    
    if not os.path.exists(excel_path):
        print(f"Excel tidak ditemukan: {excel_path}")
        sys.exit(1)
    
    print(f"Folder: {folder}")
    print(f"Excel: {EXCEL_NAME}")
    
    for wf, sheet_name in WORD_SHEET_MAP:
        wp = os.path.join(folder, wf)
        if os.path.exists(wp):
            ok = link_word_to_excel(wp, os.path.abspath(excel_path), sheet_name)
            print(f"  {'[OK]' if ok else '[FAIL]'} {wf} -> sheet '{sheet_name}'")
        else:
            print(f"  [SKIP] {wf} (tidak ada)")
    
    print("Selesai!")
    input("Tekan Enter untuk menutup...")
