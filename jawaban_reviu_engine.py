"""
jawaban_reviu_engine.py

simpan <docm_path>  -- baca CC via COM (dokumen TERBUKA di Word), simpan ke JSON
load   <docm_path>  -- inject JSON ke ZIP (dokumen HARUS TERTUTUP), buka via startfile
"""
import sys
import os
import json
import zipfile
import subprocess
from lxml import etree

NS       = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W        = '{' + NS + '}'
JSON_FILE = 'jawaban_reviu.json'

def json_path(docm_path):
    return os.path.join(os.path.dirname(docm_path), JSON_FILE)

# ──────────────────────────────────────────────
# SIMPAN via COM (dokumen terbuka)
# ──────────────────────────────────────────────

def simpan(docm_path):
    """
    Simpan CC dari ZIP langsung (dokumen harus sudah di-save dan tertutup).
    Skip CC yang menampilkan placeholder (w:showingPlcHdr).
    Merge dengan JSON lama — data yang ada tidak ditimpa jika CC kosong.
    """
    docm_path = os.path.abspath(docm_path)
    jp = json_path(docm_path)

    # Muat JSON lama sebagai base
    data = {}
    if os.path.exists(jp):
        with open(jp, encoding='utf-8') as f:
            data = json.load(f)

    with zipfile.ZipFile(docm_path, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    tree = etree.fromstring(files['word/document.xml'])
    body = tree.find('.//{%s}body' % NS)

    diperbarui = 0
    seen = set()
    for sdt in body.iter(W + 'sdt'):
        parent = sdt.getparent()
        if parent is not None and parent.tag.split('}')[-1] == 'sdtContent':
            continue  # skip nested
        sdtPr = sdt.find(W + 'sdtPr')
        if sdtPr is None:
            continue
        tagEl = sdtPr.find(W + 'tag')
        if tagEl is None:
            continue
        tag = tagEl.get(W + 'val', '')
        if not tag or tag in seen:
            continue
        seen.add(tag)

        # Skip CC yang masih menampilkan placeholder
        if sdtPr.find(W + 'showingPlcHdr') is not None:
            continue

        sdtContent = sdt.find(W + 'sdtContent')
        if sdtContent is None:
            continue

        # Skip jika teks kosong
        teks = ''.join(sdtContent.itertext()).strip()
        if not teks:
            continue

        data[tag] = etree.tostring(sdtContent, encoding='unicode')
        diperbarui += 1

    with open(jp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    size_kb = os.path.getsize(jp) / 1024
    print(f"[OK] Simpan: {diperbarui} CC diperbarui, {len(data)} total -> {jp} ({size_kb:.0f} KB)")
    return 0

# ──────────────────────────────────────────────
# LOAD via ZIP injection (dokumen tertutup)
# ──────────────────────────────────────────────

def load(docm_path):
    docm_path = os.path.abspath(docm_path)
    jp = json_path(docm_path)

    if not os.path.exists(jp):
        print(f"[ERROR] File JSON tidak ditemukan: {jp}")
        return 1

    with open(jp, encoding='utf-8') as f:
        data = json.load(f)

    # Baca ZIP
    with zipfile.ZipFile(docm_path, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    tree = etree.fromstring(files['word/document.xml'])
    body = tree.find('.//{%s}body' % NS)

    # Kumpulkan semua target dulu
    to_update = []
    for sdt in body.iter(W + 'sdt'):
        sdtPr = sdt.find(W + 'sdtPr')
        if sdtPr is None: continue
        tagEl = sdtPr.find(W + 'tag')
        if tagEl is None: continue
        tag = tagEl.get(W + 'val', '')
        if tag and tag in data:
            to_update.append((tag, sdt))

    ok = skip = 0
    for tag, sdt in to_update:
        inner_xml = data[tag]
        # inner_xml adalah fragment (bukan sdtContent wrapper) — bungkus dulu
        wrapped = f'<w:sdtContent xmlns:w="{NS}">{inner_xml}</w:sdtContent>'
        try:
            new_content = etree.fromstring(wrapped.encode('utf-8'))
        except Exception as e:
            print(f"  [WARN] {tag}: gagal parse — {e}")
            skip += 1
            continue
        old = sdt.find(W + 'sdtContent')
        if old is not None:
            sdt.remove(old)
        sdt.append(new_content)
        ok += 1

    # Tulis balik ke ZIP
    new_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['word/document.xml'] = new_xml

    tmp = docm_path + '.tmp_load'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, files[n])
    os.replace(tmp, docm_path)

    print(f"[OK] Load: {ok} CC diperbarui, {skip} skip")

    # Tunggu Word benar-benar melepas cache, lalu buka dari disk
    import time
    time.sleep(3)
    os.startfile(docm_path)
    return 0

# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Usage: jawaban_reviu_engine.py [simpan|load] <docm_path>")
        sys.exit(1)

    mode  = sys.argv[1].lower()
    docm  = sys.argv[2]

    if not os.path.exists(docm):
        print(f"[ERROR] File tidak ditemukan: {docm}")
        sys.exit(1)

    if mode == 'simpan':
        sys.exit(simpan(docm))
    elif mode == 'load':
        sys.exit(load(docm))
    else:
        print(f"[ERROR] Mode tidak dikenal: {mode}")
        sys.exit(1)

if __name__ == '__main__':
    main()
