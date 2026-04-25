"""
jawaban_reviu_engine.py — Engine Simpan & Load jawaban reviu via XML injection.

Simpan: baca inner XML tiap CC dari docm ZIP → simpan ke jawaban_reviu.json
Load  : baca JSON → inject inner XML ke docm ZIP (file harus tertutup)

Dipanggil dari VBA:
  python.exe jawaban_reviu_engine.py simpan <docm_path>
  python.exe jawaban_reviu_engine.py load   <docm_path>
"""
import sys
import os
import json
import zipfile
import shutil
from lxml import etree

NS  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W   = '{' + NS + '}'
JSON_FILE = 'jawaban_reviu.json'

# ──────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────

def json_path(docm_path):
    return os.path.join(os.path.dirname(docm_path), JSON_FILE)

def baca_zip(docm_path):
    with zipfile.ZipFile(docm_path, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}
    return names, files

def tulis_zip(docm_path, names, files):
    tmp = docm_path + '.tmp_wr'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, files[n])
    os.replace(tmp, docm_path)

def parse_doc_xml(files):
    return etree.fromstring(files['word/document.xml'])

def serialize_doc_xml(tree):
    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)

def iter_sdt(body):
    """
    Iterasi w:sdt yang punya tag, skip yang bersarang di dalam sdtContent CC lain.
    Setiap CC ditampilkan sekali saja (parent langsung bukan sdtContent).
    """
    for elem in body.iter(W + 'sdt'):
        parent = elem.getparent()
        if parent is not None:
            parent_local = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
            if parent_local == 'sdtContent':
                continue  # nested CC, skip
        sdtPr = elem.find(W + 'sdtPr')
        if sdtPr is None:
            continue
        tagEl = sdtPr.find(W + 'tag')
        if tagEl is None:
            continue
        tag = tagEl.get(W + 'val', '')
        if tag:
            yield tag, elem

# ──────────────────────────────────────────────
# SIMPAN
# ──────────────────────────────────────────────

def simpan(docm_path):
    names, files = baca_zip(docm_path)
    tree = parse_doc_xml(files)
    body = tree.find('.//{%s}body' % NS)

    data = {}
    # Ambil hanya dari CC outer (parent bukan sdtContent)
    for tag, sdt in iter_sdt(body):
        if tag in data:
            continue  # sudah ada, skip
        sdtContent = sdt.find(W + 'sdtContent')
        if sdtContent is None:
            continue
        inner_xml = etree.tostring(sdtContent, encoding='unicode')
        data[tag] = inner_xml

    jp = json_path(docm_path)
    with open(jp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=None)

    size_kb = os.path.getsize(jp) / 1024
    print(f"[OK] Simpan: {len(data)} CC -> {jp} ({size_kb:.0f} KB)")
    return 0

# ──────────────────────────────────────────────
# LOAD
# ──────────────────────────────────────────────

def load(docm_path):
    jp = json_path(docm_path)
    if not os.path.exists(jp):
        print(f"[ERROR] File JSON tidak ditemukan: {jp}")
        return 1

    with open(jp, encoding='utf-8') as f:
        data = json.load(f)

    names, files = baca_zip(docm_path)
    tree = parse_doc_xml(files)
    body = tree.find('.//{%s}body' % NS)

    # Kumpulkan dulu semua target, baru modifikasi (menghindari iterator instability)
    to_update = []
    for sdt in body.iter(W + 'sdt'):
        sdtPr = sdt.find(W + 'sdtPr')
        if sdtPr is None:
            continue
        tagEl = sdtPr.find(W + 'tag')
        if tagEl is None:
            continue
        tag = tagEl.get(W + 'val', '')
        if not tag or tag not in data:
            continue
        to_update.append((tag, sdt))

    ok = skip = 0
    for tag, sdt in to_update:
        new_content_xml = data[tag]
        try:
            new_sdt_content = etree.fromstring(new_content_xml.encode('utf-8'))
        except Exception as e:
            print(f"  [WARN] {tag}: gagal parse XML — {e}")
            skip += 1
            continue
        old_content = sdt.find(W + 'sdtContent')
        if old_content is not None:
            sdt.remove(old_content)
        sdt.append(new_sdt_content)
        ok += 1

    files['word/document.xml'] = serialize_doc_xml(tree)
    tulis_zip(docm_path, names, files)

    print(f"[OK] Load: {ok} CC diperbarui, {skip} skip")
    return 0

# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Usage: jawaban_reviu_engine.py [simpan|load] <docm_path>")
        sys.exit(1)

    mode = sys.argv[1].lower()
    docm = os.path.abspath(sys.argv[2])

    if not os.path.exists(docm):
        print(f"[ERROR] File tidak ditemukan: {docm}")
        sys.exit(1)

    if mode == 'simpan':
        sys.exit(simpan(docm))
    elif mode == 'load':
        sys.exit(load(docm))
    else:
        print(f"[ERROR] Mode tidak dikenal: {mode} (gunakan 'simpan' atau 'load')")
        sys.exit(1)

if __name__ == '__main__':
    main()
