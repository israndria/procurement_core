"""
jawaban_reviu_engine.py

simpan <docm_path>  -- baca CC dari ZIP (tanpa close), simpan ke JSON
load   <docm_path>  -- inject XML ke ZIP -> doc.Close -> os.replace -> reopen di Word yg sama
"""
import sys
import os
import json
import zipfile
import shutil
from lxml import etree

NS        = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W         = '{' + NS + '}'
JSON_FILE = 'jawaban_reviu.json'

PLACEHOLDER_TEXTS = {
    'click or tap here to enter text.',
    'click or tap here to enter text',
    'klik atau ketuk di sini untuk memasukkan teks.',
    'klik atau ketuk di sini untuk memasukkan teks',
}

def json_path(docm_path):
    return os.path.join(os.path.dirname(os.path.abspath(docm_path)), JSON_FILE)


# ──────────────────────────────────────────────
# SIMPAN: baca ZIP (dokumen boleh tetap terbuka)
# Word membiarkan file terbaca walaupun dibuka;
# jika locked untuk read, fallback ke save+close+baca+reopen.
# ──────────────────────────────────────────────

def _baca_zip(docm_path):
    """Baca word/document.xml dari ZIP. Return bytes atau None jika gagal."""
    try:
        with zipfile.ZipFile(docm_path, 'r') as z:
            return z.read('word/document.xml')
    except Exception as e:
        print(f"[WARN] Tidak bisa baca ZIP langsung: {e}")
        return None


def simpan(docm_path):
    import win32com.client, pythoncom

    docm_path = os.path.abspath(docm_path)
    jp = json_path(docm_path)

    # Coba baca ZIP langsung (Word biasanya tidak lock untuk read)
    doc_xml = _baca_zip(docm_path)

    if doc_xml is None:
        # Fallback: simpan via COM dulu, lalu baca ZIP
        print("[INFO] Fallback: save via COM, baca ZIP")
        pythoncom.CoInitialize()
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            norm = os.path.normcase(docm_path)
            doc = None
            for d in word.Documents:
                if os.path.normcase(d.FullName) == norm:
                    doc = d
                    break
            if doc:
                doc.Save()
        finally:
            pythoncom.CoUninitialize()
        doc_xml = _baca_zip(docm_path)

    if doc_xml is None:
        print("[ERROR] Tidak bisa membaca dokumen XML")
        return 1

    # Muat JSON lama sebagai base (merge strategy)
    data = {}
    if os.path.exists(jp):
        with open(jp, encoding='utf-8') as f:
            data = json.load(f)

    tree = etree.fromstring(doc_xml)
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

        teks = ''.join(sdtContent.itertext()).strip()
        if not teks or teks.lower() in PLACEHOLDER_TEXTS:
            continue

        data[tag] = etree.tostring(sdtContent, encoding='unicode')
        diperbarui += 1

    with open(jp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    size_kb = os.path.getsize(jp) / 1024
    print(f"[OK] Simpan: {diperbarui} CC -> {jp} ({size_kb:.0f} KB)")
    return 0


# ──────────────────────────────────────────────
# LOAD: inject ZIP -> Close -> Replace -> Reopen
# Dokumen tutup sebentar lalu buka kembali (user
# bisa menyaksikan konten baru muncul).
# ──────────────────────────────────────────────

def _inject_zip(docm_path, data):
    """
    Baca ZIP, inject CC ke document.xml, tulis ke tmp, return tmp path.
    Tidak menyentuh file asli — caller yang os.replace.
    """
    tmp = docm_path + '.tmp_inject'
    shutil.copy2(docm_path, tmp)

    with zipfile.ZipFile(tmp, 'r') as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    doc_xml = files.get('word/document.xml')
    if doc_xml is None:
        os.remove(tmp)
        raise RuntimeError("word/document.xml tidak ditemukan dalam docm")

    tree = etree.fromstring(doc_xml)
    body = tree.find('.//{%s}body' % NS)
    if body is None:
        os.remove(tmp)
        raise RuntimeError("w:body tidak ditemukan")

    # Kumpulkan semua target dulu (hindari iterator instability)
    targets = []
    for sdt in body.iter(W + 'sdt'):
        parent = sdt.getparent()
        if parent is not None and parent.tag.split('}')[-1] == 'sdtContent':
            continue
        sdtPr = sdt.find(W + 'sdtPr')
        if sdtPr is None:
            continue
        tagEl = sdtPr.find(W + 'tag')
        if tagEl is None:
            continue
        tag = tagEl.get(W + 'val', '')
        if tag and tag in data:
            targets.append((tag, sdt))

    ok = skip = err = 0
    done_tags = set()
    for tag, sdt in targets:
        if tag in done_tags:
            continue
        done_tags.add(tag)
        inner_xml = data[tag]
        try:
            new_content = etree.fromstring(inner_xml)
        except Exception as e:
            print(f"  [WARN] {tag}: gagal parse XML: {e}")
            err += 1
            continue

        sdtContent = sdt.find(W + 'sdtContent')
        if sdtContent is None:
            sdtContent = etree.SubElement(sdt, W + 'sdtContent')
        else:
            for child in list(sdtContent):
                sdtContent.remove(child)
        for child in list(new_content):
            from copy import deepcopy
            sdtContent.append(deepcopy(child))

        # Hapus showingPlcHdr jika ada
        sdtPr = sdt.find(W + 'sdtPr')
        if sdtPr is not None:
            ph = sdtPr.find(W + 'showingPlcHdr')
            if ph is not None:
                sdtPr.remove(ph)
        ok += 1

    print(f"  Inject: {ok} CC ditulis, {skip} skip, {err} error")

    new_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['word/document.xml'] = new_xml

    out = docm_path + '.injected'
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])

    os.remove(tmp)
    return out


def load(docm_path):
    import win32com.client, pythoncom

    docm_path = os.path.abspath(docm_path)
    jp = json_path(docm_path)

    if not os.path.exists(jp):
        print(f"[ERROR] File JSON tidak ditemukan: {jp}")
        return 1

    with open(jp, encoding='utf-8') as f:
        data = json.load(f)

    if not data:
        print("[ERROR] JSON kosong, tidak ada yang diload")
        return 1

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            print("[ERROR] Word tidak terbuka. Buka dokumen di Word dulu.")
            return 1

        # Cari dokumen yang terbuka
        norm = os.path.normcase(docm_path)
        doc = None
        for d in word.Documents:
            if os.path.normcase(d.FullName) == norm:
                doc = d
                break
        if doc is None:
            print(f"[ERROR] Dokumen tidak terbuka di Word: {docm_path}")
            for d in word.Documents:
                print(f"  {d.FullName}")
            return 1

        print(f"[INFO] Dokumen: {doc.Name} | {doc.ContentControls.Count} CC")

        # Inject ke ZIP (file sementara)
        print("[INFO] Menyiapkan ZIP baru dengan jawaban...")
        injected = _inject_zip(docm_path, data)

        # Tutup dokumen tanpa save (akan diganti dengan file baru)
        print("[INFO] Menutup dokumen (sementara)...")
        doc.Close(SaveChanges=False)

        # Ganti file asli dengan yang sudah diinjeksi
        os.replace(injected, docm_path)
        print("[INFO] File diganti, membuka kembali...")

        # Buka kembali di instance Word yang sama
        word.Documents.Open(docm_path, False, False, False)
        print("[OK] Dokumen dibuka kembali dengan jawaban reviu.")
        return 0

    finally:
        pythoncom.CoUninitialize()


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Usage: jawaban_reviu_engine.py [simpan|load] <docm_path>")
        sys.exit(1)

    mode = sys.argv[1].lower()
    docm = sys.argv[2]

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
