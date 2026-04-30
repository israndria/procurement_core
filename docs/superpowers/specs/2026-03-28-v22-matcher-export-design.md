# V22 Inaproc Order Bot — Matcher Improvement & Export Hasil

**Tanggal:** 2026-03-28
**Versi dasar:** v1.2 STABLE (27/30 berhasil, 90%)
**Scope:** Fitur A (matcher multi-candidate) + Fitur C (export hasil 3 format)

---

## Konteks

Bot v1.2 sudah stabil dengan 3 kasus yang masih manual:
- **Variasi tidak cocok** (⚠️): matcher gagal sama sekali — nama barang tidak mengandung petunjuk variasi
- **Variasi ambigu**: matcher gagal karena nama di Excel memiliki angka yang bisa cocok ke beberapa variasi (contoh: "Staples No. 10-12" vs variasi "Max 10", "Max 1213", "GW10", "Kangaru 1210")

Fitur A memisahkan kedua kasus ini: ambigu tetap masuk keranjang (optimistic) dengan tanda, sedangkan tidak cocok tetap skip.
Fitur C memberikan 3 pilihan export setelah proses selesai.

---

## Fitur A — Matcher Multi-Candidate (Optimistic Flag)

### Tujuan
Menambah **Pass 3** di matcher: ekstrak angka dari nama barang dan cari variasi yang share angka yang sama. Jika ditemukan kandidat:
- Tepat 1 kandidat → pilih otomatis, masuk keranjang, status `variasi_tidak_pasti` (bukan `berhasil`)
- 2+ kandidat → pilih kandidat pertama, masuk keranjang, status `variasi_tidak_pasti`, keterangan berisi semua kandidat
- 0 kandidat → tetap `variasi_tidak_cocok`, tidak masuk keranjang

**Item dengan status `variasi_tidak_pasti` WAJIB di-crosscheck user di keranjang** — bisa dihapus atau dibiarkan tergantung kebenarannya.

### Perubahan `matcher.py`

Fungsi baru:
```python
def _extract_angka(teks: str) -> set[str]:
    """Extract semua sequence angka dari teks. "No. 10-12" → {"10", "12"}"""
    return set(re.findall(r'\d+', teks))
```

Return type `cocokkan_variasi` berubah dari `str | None` menjadi **tuple** `(str | None, list[str])`:
- `(variasi, [])` → match pasti (pass 1 atau 2), kandidat kosong
- `(kandidat_terpilih, semua_kandidat)` → ambigu (pass 3), kandidat tidak kosong
- `(None, [])` → tidak cocok sama sekali

Logic Pass 3:
1. Ekstrak angka dari `nama_barang`
2. Untuk setiap variasi, ekstrak angkanya juga
3. Kumpulkan variasi yang share setidaknya 1 angka → list kandidat
4. Jika kandidat tidak kosong → return `(kandidat[0], kandidat)`

### Perubahan `order_bot.py`

`proses_item()` di-update untuk handle tuple baru:
```python
cocok, kandidat = cocokkan_variasi(nama, variasi_tersedia)
if cocok is None and not kandidat:
    # variasi_tidak_cocok — skip seperti sebelumnya
    ...
elif kandidat:
    # variasi_tidak_pasti — pilih cocok, masuk keranjang, tapi flag
    self._pilih_variasi(cocok)
    # ... isi qty, klik tambah ...
    kandidat_lain = [k for k in kandidat if k != cocok]
    pesan = f"Dipilih: {cocok} (ambigu)."
    if kandidat_lain:
        pesan += f" Kandidat lain: {', '.join(kandidat_lain)}"
    return {**base, "status": "variasi_tidak_pasti", "pesan": pesan}
```

### Perubahan `app.py`

- Tambah ikon `🔄` untuk status `variasi_tidak_pasti`
- Tambah metric ke-5: **🔄 Perlu Cek**
- Di tabel hasil, kolom Keterangan untuk item `variasi_tidak_pasti` berisi daftar kandidat

---

## Fitur C — Export Hasil (3 Format)

### Tujuan
Setelah proses selesai, user bisa download tabel hasil dalam 3 format.

### Konten semua format
Kolom: **No, Status, Nama Barang, Qty, Keterangan, Link**

Status ditulis sebagai teks (bukan emoji) untuk CSV/Excel agar mudah di-filter:
`berhasil`, `variasi_tidak_pasti`, `variasi_tidak_cocok`, `skip_tidak_aktif`, `error`

### Format & Library

| Tombol | Format | Library | Catatan |
|---|---|---|---|
| `📥 Download CSV` | `.csv` | pandas `to_csv()` | Plain text, tanpa formatting |
| `📥 Download Excel` | `.xlsx` | openpyxl | Header bold, kolom auto-width |
| `📥 Download PDF` | `.pdf` | PyMuPDF (fitz) `Story` | Tabel dengan warna per status |

Nama file: `hasil_order_YYYY-MM-DD.csv/.xlsx/.pdf` (tanggal otomatis dari `datetime.date.today()`).

### Warna status di PDF

| Status | Warna baris |
|---|---|
| berhasil | Hijau muda `#d4f4dd` |
| variasi_tidak_pasti | Kuning muda `#fff3cd` |
| variasi_tidak_cocok | Oranye muda `#fde8d8` |
| skip_tidak_aktif | Abu-abu muda `#f0f0f0` |
| error | Merah muda `#fdd` |

### Perubahan `app.py`

Di Bagian 4, tambahkan 3 tombol berjajar (`st.columns(3)`) di atas tombol "Proses Ulang":
```python
col_csv, col_xlsx, col_pdf = st.columns(3)
with col_csv:
    st.download_button("📥 Download CSV", data=generate_csv(hasil), ...)
with col_xlsx:
    st.download_button("📥 Download Excel", data=generate_excel(hasil), ...)
with col_pdf:
    st.download_button("📥 Download PDF", data=generate_pdf(hasil), ...)
```

Fungsi `generate_csv`, `generate_excel`, `generate_pdf` ditulis sebagai helper di `app.py` (bukan file terpisah — hanya dipakai di satu tempat).

---

## Files yang Berubah

| File | Perubahan |
|---|---|
| `matcher.py` | Tambah `_extract_angka()`, ubah return type `cocokkan_variasi` → tuple |
| `order_bot.py` | Update `proses_item()` handle tuple + status `variasi_tidak_pasti` |
| `app.py` | Tambah metric 🔄, ikon status, 3 tombol download + helper functions |

**Files baru:** tidak ada

---

## Status Baru

| Status | Ikon | Arti |
|---|---|---|
| `berhasil` | ✅ | Masuk keranjang, variasi cocok pasti |
| `variasi_tidak_pasti` | 🔄 | Masuk keranjang, variasi dipilih otomatis tapi ambigu — **wajib crosscheck** |
| `variasi_tidak_cocok` | ⚠️ | Tidak masuk keranjang, tidak ada kandidat |
| `skip_tidak_aktif` | ⏭ | Produk tidak aktif |
| `error` | ❌ | Error teknis |
| `session_expired` | 🔒 | Session Edge expired |
