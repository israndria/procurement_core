# Protokol Revisi XML PL JKK/PLPK

## Tujuan

XML menjadi working copy untuk membandingkan revisi dokumen PPK tanpa membuka
atau mengubah workbook Excel secara langsung. Excel hanya disentuh ketika
user menjalankan `Load Data`.

## File per folder paket

- `input_data_baseline.xml` — snapshot pertama yang immutable.
- `input_data_snapshot.xml` — current state yang terakhir diterima Excel.
- `input_data_proposal.xml` — salinan kerja yang boleh diedit AI.
- `input_data_audit.jsonl` — audit setiap proposal yang dipromosikan.
- `input_data_snapshot.bak-*.xml` — backup otomatis sebelum current diganti.

Proposal yang dibuat melalui `seed-proposal` membawa atribut
`source_sha256`. Atribut ini mencegah proposal lama menimpa snapshot current
yang sudah berubah karena `Save Data` baru.

## Kontrak profile/layout

Snapshot baru memakai `version="2"` dan wajib menyimpan:

- `family="PLJKK"`, `layout_version="PLJKK-MASTER-DATA-v1"`; atau
- `family="PLPK"`, `layout_version="PLPK-MASTER-DATA-v1"`.

Map PLJKK dan PLPK tidak boleh digabung. Alamat yang sama mempunyai arti
berbeda pada kedua workbook, terutama `C51:C63`. Profile PLPK mencakup
personil `C33:C38`, alat/kapasitas/jumlah `C39:C56`, risiko K3 `C63:C64`,
uraian HPS `C66:C75`, peserta `C77:C80`, metadata Nota Dinas `C82:C89`,
serta blok tanggal Dokpil `H15:H18` dan `I17:I18`.

`C57:C62` pada PLPK adalah field helper/reserved dan sengaja tidak masuk
snapshot. XML v1 tanpa atribut profile tetap dibaca sebagai legacy PLJKK;
XML v1 tidak boleh dipakai untuk workbook PLPK.

`Save Data` membuat baseline hanya bila baseline belum ada. Save berikutnya
hanya memperbarui current dan tidak menimpa baseline.

## Alur kerja revisi

1. Setelah upload PPK pertama selesai, jalankan `Save Data`.
2. AI membaca `input_data_baseline.xml` dan dokumen PPK revisi.
3. Buat proposal dari current:

   ```powershell
   & $env:POKJA_PYTHON pl_snapshot_revision.py seed-proposal `
     input_data_snapshot.xml input_data_proposal.xml
   ```

4. AI membaca dokumen revisi, lalu mengubah hanya field yang terbukti berbeda
   pada `input_data_proposal.xml`. Pertahankan atribut root, terutama
   `kode_paket` dan `source_sha256`.
5. Buat laporan perbandingan:

   ```powershell
   & $env:POKJA_PYTHON pl_snapshot_revision.py compare `
     input_data_baseline.xml input_data_proposal.xml `
     --output revisi_snapshot_report.md
   ```

6. Setelah user menyetujui, proposal dipromosikan ke current:

   ```powershell
   & $env:POKJA_PYTHON pl_snapshot_revision.py promote `
     input_data_proposal.xml input_data_snapshot.xml `
     --baseline input_data_baseline.xml `
     --audit input_data_audit.jsonl `
     --expected-kode-paket 11000000000
   ```

   Validasi profile dapat diperketat dengan `--expected-family PLPK` atau
   `--expected-family PLJKK`.

7. Jalankan `Load Data` di Excel. VBA membaca current yang sudah divalidasi.

## Aturan AI

- Jangan menimpa `input_data_baseline.xml`.
- Jangan menghapus node `<cell>` untuk mengosongkan Excel. Gunakan
  `type="empty"`; node yang hilang berarti field tidak diubah.
- Jangan mengubah `C3` atau `F2`; keduanya identitas paket/kode unik.
- Jangan mengubah formula `C11`, `C12`, `C20`, `C22`, `C24`, `C26`, `H10`,
  `H11`, `I8`, `I9`, `I10`; ubah input sumbernya.
- Perbedaan angka dinormalisasi secara numerik, tetapi perubahan teks tetap
  dilaporkan agar user dapat menilai substansinya.
- Setiap perbedaan harus memiliki sumber dokumen dan halaman dalam laporan AI.
- Jika kode paket kosong, berbeda, atau proposal tidak lengkap, promosi ditolak.
- Proposal, current, dan baseline dengan family/layout berbeda ditolak.
- Jika current berubah setelah proposal dibuat, promosi ditolak dan proposal
  harus dibuat ulang. Ini mencegah revisi stale menimpa koreksi terbaru.
- File XML dapat dibaca, dibandingkan, dan diedit AI tanpa membuka Excel.
  Excel hanya disentuh pada tahap akhir `Load Data`; Python tidak memakai
  `openpyxl.save()` atau COM.

## Contoh laporan

```text
C15 - Jangka Waktu
Awal   : 180 Hari
Revisi : 90 Hari
Sumber : KAK revisi, halaman 3
Status : Perlu persetujuan
```

Modul `pl_snapshot_revision.py` hanya bekerja pada XML lokal. Tidak ada
`openpyxl.save()` dan tidak ada COM Excel di jalur validasi/promosi.
