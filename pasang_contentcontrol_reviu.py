"""
Pasang Content Control (RichText) ke 2. Isi Reviu PK v2.docm:
- Tag rekomen_1..8  : sel tabel rekomendasi tiap seksi
- Tag tanggapan_1..8: sel tabel tanggapan PPK tiap seksi
- Tag cat_A_1..N    : kolom catatan per baris tabel seksi A-H

Content Control tidak hilang walau konten dikosongkan/dihapus —
berbeda dengan Bookmark yang ikut terhapus bersama teksnya.

Jalankan sekali saja setelah docm baru dibuat/direname.
Script ini akan menghapus Content Control lama (by Tag) sebelum pasang ulang.
"""
import win32com.client
import pythoncom
import sys
import zipfile
from lxml import etree

DOCM_PATH = r"D:\Dokumen\@ POKJA 2026\Paket Experiment\2. Isi Reviu PK v2.docm"
SEKSI = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

LABELS_REKOMEN   = {'Rekomendasi / Catatan Hasil Reviu:', 'Rekomendasi Hasil Reviu:'}
LABELS_TANGGAPAN = {'Tanggapan PPK atas Rekomendasi / Catatan Hasil Reviu:',
                    'Tanggapan PPK atas Rekomendasi Hasil Reviu:'}

def get_text_xml(el, ns):
    return ''.join((t.text or '') for t in el.findall('.//w:t', ns)).strip()

def klasifikasi_tabel():
    """Pakai lxml untuk mapping tabel index (1-based) ke nama CC."""
    with zipfile.ZipFile(DOCM_PATH) as z:
        xml = z.read('word/document.xml')
    tree = etree.fromstring(xml)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = tree.find('.//w:body', ns)
    elements = list(body)

    mapping = {}
    tbl_counter = 0
    rek_count = 0
    tan_count = 0
    cat_count = 0

    for i, el in enumerate(elements):
        if el.tag.split('}')[-1] != 'tbl':
            continue
        tbl_counter += 1
        rows = el.findall('.//w:tr', ns)

        prev_text = ''
        for j in range(i-1, max(0, i-4), -1):
            pt = get_text_xml(elements[j], ns)
            if pt:
                prev_text = pt
                break

        if prev_text in LABELS_REKOMEN and rek_count < 8:
            rek_count += 1
            mapping[tbl_counter] = {'type': 'rekomen', 'name': f'rekomen_{rek_count}'}
        elif prev_text in LABELS_TANGGAPAN and tan_count < 8:
            tan_count += 1
            mapping[tbl_counter] = {'type': 'tanggapan', 'name': f'tanggapan_{tan_count}'}
        elif cat_count < 8 and len(rows) > 2:
            header = get_text_xml(rows[0], ns)
            if 'No' in header and 'Catatan' in header:
                sek = SEKSI[cat_count]
                n_rows = len(rows) - 1  # exclude header
                mapping[tbl_counter] = {'type': 'catatan', 'name': sek, 'rows': n_rows}
                cat_count += 1

    return mapping


def hapus_cc_lama(doc, prefix_list):
    """Hapus semua Content Control yang Tag-nya diawali prefix tertentu."""
    to_delete = []
    for cc in doc.ContentControls:
        for prefix in prefix_list:
            if cc.Tag.startswith(prefix):
                to_delete.append(cc)
                break
    for cc in to_delete:
        try:
            cc.LockContentControl = False
            cc.Delete(DeleteContents=False)
        except Exception as e:
            print(f"  [WARN] Gagal hapus CC '{cc.Tag}': {e}")
    print(f"  Hapus {len(to_delete)} CC lama.")


def pasang_cc(doc, rng, tag, title=""):
    """Pasang Rich Text Content Control ke range tertentu."""
    # wdContentControlRichText = 0
    cc = doc.ContentControls.Add(0, rng)
    cc.Tag = tag
    cc.Title = title if title else tag
    cc.LockContentControl = True   # CC tidak bisa dihapus oleh user
    cc.LockContents = False        # Isi tetap bisa diedit
    return cc


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    mapping = klasifikasi_tabel()

    print("Mapping tabel:")
    for idx, info in mapping.items():
        print(f"  tbl[{idx}] -> {info}")

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(DOCM_PATH, False, False, False)
        print(f"\nDibuka: {doc.Name} ({doc.Tables.Count} tabel, {doc.ContentControls.Count} CC existing)")

        # Hapus CC lama dulu
        prefixes = ['rekomen_', 'tanggapan_', 'cat_']
        hapus_cc_lama(doc, prefixes)

        cc_list = []

        for tbl_idx, info in mapping.items():
            if tbl_idx > doc.Tables.Count:
                print(f"  [SKIP] tbl[{tbl_idx}] melebihi jumlah tabel ({doc.Tables.Count})")
                continue

            tbl = doc.Tables(tbl_idx)
            t = info['type']

            if t in ('rekomen', 'tanggapan'):
                try:
                    rng = tbl.Cell(1, 1).Range
                    rng.MoveEnd(1, -1)
                    tag = info['name']
                    cc = pasang_cc(doc, rng, tag)
                    print(f"  [OK] CC '{tag}' @ tbl[{tbl_idx}]")
                    cc_list.append(tag)
                except Exception as e:
                    print(f"  [WARN] {info['name']}: {e}")

            elif t == 'catatan':
                sek = info['name']
                n_rows = info['rows']
                for r in range(1, n_rows + 1):
                    try:
                        if tbl.Rows.Count < r + 1:
                            break
                        if tbl.Rows(r + 1).Cells.Count < 3:
                            continue
                        rng = tbl.Cell(r + 1, 3).Range
                        rng.MoveEnd(1, -1)
                        tag = f"cat_{sek}_{r}"
                        teks_preview = (rng.Text or '')[:40].strip()
                        cc = pasang_cc(doc, rng, tag)
                        print(f"  [OK] CC '{tag}': {teks_preview}")
                        cc_list.append(tag)
                    except Exception as e:
                        print(f"  [WARN] cat_{sek}_{r}: {e}")

        try:
            doc.Save()
        except Exception:
            doc.SaveAs2(DOCM_PATH, FileFormat=13)

        print(f"\n[OK] {len(cc_list)} Content Control dipasang dan disimpan.")
        print("\nContent Control yang perlu DIHAPUS manual (dari mail merge):")
        print("  - cat_E_1 (ID RUP — dari Excel)")
        print("  - cat_F_1 (masa pelaksanaan — dari Excel)")
        print("  - cat_F_3 (bulan penggunaan — dari Excel)")
        print("  Hapus via Word: klik sel → Developer tab → pilih CC → Delete")

        doc.Close(SaveChanges=False)

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        try: word.Quit()
        except: pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
