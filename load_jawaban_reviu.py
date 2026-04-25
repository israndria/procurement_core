"""
Load jawaban reviu dari jawaban_reviu.json → inject XML ke Content Control di docm.
Dipanggil dari VBA setelah dokumen ditutup.
Argument: path ke docm (wajib)

Cara kerja:
- Baca jawaban_reviu.json dari folder yang sama dengan docm
- Buka docm sebagai ZIP
- Parse word/document.xml
- Untuk setiap CC yang punya tag di JSON, ganti isi paragraf dengan XML tersimpan
- Simpan kembali ke ZIP
"""
import sys
import os
import json
import zipfile
import shutil
import re
from lxml import etree
from copy import deepcopy

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{' + NS + '}'

def baca_json(json_path):
    with open(json_path, encoding='utf-8') as f:
        return json.load(f)

def xml_isi_dari_wordopenxml(word_open_xml: str):
    """
    WordOpenXML adalah dokumen lengkap. Ambil paragraph(s) di dalamnya
    — isi sebenarnya yang ada di dalam CC Range.
    Kembalikan list elemen lxml (w:p atau w:tbl).
    """
    try:
        root = etree.fromstring(word_open_xml.encode('utf-8'))
    except Exception:
        # Coba strip BOM / encoding declaration
        cleaned = re.sub(r'<\?xml[^>]+\?>', '', word_open_xml).strip()
        root = etree.fromstring(cleaned.encode('utf-8'))

    # WordOpenXML membungkus dalam w:wordDocument > w:body
    # Ambil semua child body kecuali w:sectPr
    body = root.find('.//{%s}body' % NS)
    if body is None:
        return []
    elems = []
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            continue
        elems.append(deepcopy(child))
    return elems

def temukan_cc_by_tag(body, tag):
    """
    Cari elemen w:sdt (Content Control) yang punya w:tag w:val == tag.
    Return elemen sdt dan parent-nya.
    """
    for parent in body.iter():
        for sdt in list(parent):
            if not hasattr(sdt, 'tag'):
                continue
            local = sdt.tag.split('}')[-1] if '}' in sdt.tag else sdt.tag
            if local != 'sdt':
                continue
            # Cek sdtPr > tag > val
            sdtPr = sdt.find(W + 'sdtPr')
            if sdtPr is None:
                continue
            tagEl = sdtPr.find(W + 'tag')
            if tagEl is None:
                continue
            val = tagEl.get(W + 'val', '')
            if val == tag:
                return sdt, parent
    return None, None

def ganti_isi_cc(sdt, new_elems):
    """
    Ganti konten w:sdtContent dengan elemen baru.
    Pertahankan sdtPr (properti CC).
    """
    sdtContent = sdt.find(W + 'sdtContent')
    if sdtContent is None:
        sdtContent = etree.SubElement(sdt, W + 'sdtContent')
    else:
        # Hapus semua anak lama
        for child in list(sdtContent):
            sdtContent.remove(child)

    for elem in new_elems:
        sdtContent.append(elem)

def inject(docm_path, json_path):
    data = baca_json(json_path)
    print(f"JSON: {len(data)} CC akan di-inject")

    # Baca ZIP
    tmp_path = docm_path + '.tmp_inject'
    shutil.copy2(docm_path, tmp_path)

    with zipfile.ZipFile(tmp_path, 'r') as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    doc_xml = files.get('word/document.xml')
    if doc_xml is None:
        print("[ERROR] word/document.xml tidak ditemukan dalam docm")
        sys.exit(1)

    tree = etree.fromstring(doc_xml)
    body = tree.find('.//{%s}body' % NS)
    if body is None:
        print("[ERROR] w:body tidak ditemukan")
        sys.exit(1)

    ok = 0
    skip = 0
    for tag, word_xml in data.items():
        if not word_xml:
            continue
        new_elems = xml_isi_dari_wordopenxml(word_xml)
        if not new_elems:
            print(f"  [WARN] {tag}: XML kosong setelah parse")
            skip += 1
            continue
        sdt, parent = temukan_cc_by_tag(body, tag)
        if sdt is None:
            print(f"  [SKIP] CC tag='{tag}' tidak ditemukan di dokumen")
            skip += 1
            continue
        ganti_isi_cc(sdt, new_elems)
        ok += 1

    print(f"Inject selesai: {ok} berhasil, {skip} skip")

    # Serialize kembali
    new_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['word/document.xml'] = new_xml

    # Tulis ZIP baru
    out_path = docm_path + '.injected'
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])

    # Ganti file asli
    os.replace(out_path, docm_path)
    if os.path.exists(tmp_path):
        os.remove(tmp_path)

    print(f"[OK] Dokumen diperbarui: {docm_path}")

def main():
    if len(sys.argv) < 2:
        print("[ERROR] Argumen: load_jawaban_reviu.py <path_docm>")
        sys.exit(1)

    docm_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(docm_path):
        print(f"[ERROR] File tidak ditemukan: {docm_path}")
        sys.exit(1)

    folder = os.path.dirname(docm_path)
    json_path = os.path.join(folder, 'jawaban_reviu.json')
    if not os.path.exists(json_path):
        print(f"[ERROR] File JSON tidak ditemukan: {json_path}")
        sys.exit(1)

    inject(docm_path, json_path)

if __name__ == '__main__':
    main()
