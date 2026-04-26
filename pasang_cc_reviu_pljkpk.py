"""
Pasang Content Control (RichText) ke 2. Isi Reviu PLJKPK v2.docm:

Bagian I  (Reviu Dokumen Persiapan):
  - cat_A_N  : Tbl 1 (KAK, 25 baris), Tbl 2 (klarifikasi KAK)
  - rekomen_1: Tbl 3
  - cat_B_N  : Tbl 4 (HPS), Tbl 5 (klarifikasi HPS)
  - rekomen_2: Tbl 6
  - cat_C_N  : Tbl 7 (Kontrak), Tbl 8 (klarifikasi Kontrak)
  - rekomen_3: Tbl 9
  - cat_D_N  : Tbl 10 (Anggaran)
  - rekomen_4: Tbl 11
  - cat_E_N  : Tbl 12 (RUP)
  - rekomen_5: Tbl 13
  - cat_F_N  : Tbl 14 (Waktu)
  - rekomen_6: Tbl 15
  - cat_G_N  : Tbl 16 (Pasar)
  - rekomen_7: Tbl 17
  - cat_H_N  : Tbl 18 (Masukan PPK / Kualifikasi)

Bagian II (Penetapan Metode):
  - cat_I_N  : Tbl 19 (Metode Pemilihan)
  - cat_J_N  : Tbl 20 (Metode Kualifikasi)
  - cat_K_N  : Tbl 21 (Persyaratan Kualifikasi)
  - cat_L_N  : Tbl 22 (Persyaratan Teknis)
  - cat_M_N  : Tbl 23 (Metode Evaluasi)
  - cat_N_N  : Tbl 24 (Metode Penyampaian)
  - cat_O_N  : Tbl 25 (Jadwal)
  - cat_P_N  : Tbl 26 (Dokumen Pemilihan)

Tbl 27-28 (Lembar Kriteria Evaluasi) — TIDAK diberi CC (formulir terpisah).

Jalankan sekali saja setelah docm baru dibuat.
"""
import win32com.client
import pythoncom
import sys

DOCM_PATH = r"D:\Dokumen\@ POKJA 2026\Paket Experiment - Pengadaan Langsung - Konsultan Konstuksi\2. Isi Reviu PLJKPK v2.docm"

# Mapping tabel index (1-based) -> konfigurasi CC
# type: 'catatan' -> pasang CC di kolom 3, tiap baris data (skip baris header)
#        'rekomen' -> pasang CC di sel (1,1) tabel 1-baris
# Tabel multi-baris dengan catatan + klarifikasi dalam satu seksi digabung
# dengan tag sekuensial (A_1..A_N lanjut dari tabel sebelumnya)

# Bagian I
BAGIAN_I = [
    # (tbl_idx, tag_prefix, type)
    (1,  'cat_A', 'catatan'),
    (2,  'cat_A', 'catatan_lanjut'),   # lanjut nomor dari tbl 1
    (3,  'rekomen_1', 'rekomen'),
    (4,  'cat_B', 'catatan'),
    (5,  'cat_B', 'catatan_lanjut'),
    (6,  'rekomen_2', 'rekomen'),
    (7,  'cat_C', 'catatan'),
    (8,  'cat_C', 'catatan_lanjut'),
    (9,  'rekomen_3', 'rekomen'),
    (10, 'cat_D', 'catatan'),
    (11, 'rekomen_4', 'rekomen'),
    (12, 'cat_E', 'catatan'),
    (13, 'rekomen_5', 'rekomen'),
    (14, 'cat_F', 'catatan'),
    (15, 'rekomen_6', 'rekomen'),
    (16, 'cat_G', 'catatan'),
    (17, 'rekomen_7', 'rekomen'),
    (18, 'cat_H', 'catatan'),
]

# Bagian II (tidak ada rekomen, langsung catatan per tabel)
BAGIAN_II = [
    (19, 'cat_I', 'catatan'),
    (20, 'cat_J', 'catatan'),
    (21, 'cat_K', 'catatan'),
    (22, 'cat_L', 'catatan'),
    (23, 'cat_M', 'catatan'),
    (24, 'cat_N', 'catatan'),
    (25, 'cat_O', 'catatan'),
    (26, 'cat_P', 'catatan'),
    # Tbl 27-28: Lembar Kriteria — tidak diberi CC
]

ALL_CONFIG = BAGIAN_I + BAGIAN_II


def hapus_cc_lama(doc):
    prefixes = ('cat_', 'rekomen_')
    to_delete = []
    for cc in doc.ContentControls:
        if any(cc.Tag.startswith(p) for p in prefixes):
            to_delete.append(cc)
    for cc in to_delete:
        try:
            cc.LockContentControl = False
            cc.Delete(DeleteContents=False)
        except Exception as e:
            print(f"  [WARN] Gagal hapus CC '{cc.Tag}': {e}")
    print(f"  Hapus {len(to_delete)} CC lama.")


def pasang_cc(doc, rng, tag):
    cc = doc.ContentControls.Add(0, rng)   # 0 = wdContentControlRichText
    cc.Tag = tag
    cc.Title = tag
    cc.LockContentControl = True
    cc.LockContents = False
    return cc


def main():
    sys.stdout.reconfigure(encoding='utf-8')

    # Build counter per prefix untuk catatan_lanjut
    prefix_counter = {}

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(DOCM_PATH, False, False, False)
        print(f"Dibuka: {doc.Name} ({doc.Tables.Count} tabel, {doc.ContentControls.Count} CC existing)")

        hapus_cc_lama(doc)

        cc_list = []

        for tbl_idx, tag_prefix, tipe in ALL_CONFIG:
            if tbl_idx > doc.Tables.Count:
                print(f"  [SKIP] tbl[{tbl_idx}] melebihi jumlah tabel ({doc.Tables.Count})")
                continue

            tbl = doc.Tables(tbl_idx)

            if tipe == 'rekomen':
                try:
                    rng = tbl.Cell(1, 1).Range
                    rng.MoveEnd(1, -1)
                    cc = pasang_cc(doc, rng, tag_prefix)
                    print(f"  [OK] CC '{tag_prefix}' @ tbl[{tbl_idx}]")
                    cc_list.append(tag_prefix)
                except Exception as e:
                    print(f"  [WARN] {tag_prefix}: {e}")

            elif tipe in ('catatan', 'catatan_lanjut'):
                # Inisialisasi counter jika belum ada
                if tag_prefix not in prefix_counter:
                    prefix_counter[tag_prefix] = 0

                for r in range(2, tbl.Rows.Count + 1):  # skip baris header (baris 1)
                    try:
                        row = tbl.Rows(r)
                        if row.Cells.Count < 3:
                            continue
                        # Baris heading section (merged cell, 1 kolom) -- skip
                        # Cek: kalau kolom 1 ada teks non-angka dan tidak ada kolom 3 proper, skip
                        col1_text = tbl.Cell(r, 1).Range.Text.strip().rstrip('\r\x07')
                        col3_text = tbl.Cell(r, 3).Range.Text.strip().rstrip('\r\x07')

                        # Skip baris yang merupakan heading sub-seksi (merged, tanpa konten kolom 3)
                        # Heuristik: col1 ada teks panjang (>5 karakter) DAN col3 kosong
                        if len(col1_text) > 5 and not col3_text:
                            print(f"  [SKIP] tbl[{tbl_idx}] baris {r}: heading '{col1_text[:40]}'")
                            continue

                        prefix_counter[tag_prefix] += 1
                        tag = f"{tag_prefix}_{prefix_counter[tag_prefix]}"
                        rng = tbl.Cell(r, 3).Range
                        rng.MoveEnd(1, -1)
                        teks_preview = (rng.Text or '')[:50].strip()
                        cc = pasang_cc(doc, rng, tag)
                        print(f"  [OK] CC '{tag}': {teks_preview}")
                        cc_list.append(tag)
                    except Exception as e:
                        print(f"  [WARN] {tag_prefix}_{prefix_counter.get(tag_prefix,'?')+1}: row {r}: {e}")

        doc.Save()
        print(f"\n[OK] {len(cc_list)} Content Control dipasang dan disimpan.")
        print("\nDaftar CC yang dipasang:")
        for tag in cc_list:
            print(f"  {tag}")

        doc.Close(SaveChanges=False)

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        try:
            word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
