# V22 Inaproc Order Bot — Matcher Improvement & Export Hasil

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tambah Pass 3 di matcher (angka overlap → flag ambigu tapi tetap masuk keranjang) dan 3 tombol export hasil (CSV / Excel / PDF).

**Architecture:** Tiga file yang berubah secara berurutan — `matcher.py` (logika baru + return type baru), `order_bot.py` (konsumen tuple baru), `app.py` (UI metric + download). Tidak ada file baru.

**Tech Stack:** Python 3.12, Streamlit, openpyxl, PyMuPDF (fitz), pandas — semua sudah terinstall di `C:\Users\MSI\AppData\Local\Programs\Python\Python312\`.

---

## File Map

| File | Perubahan |
|---|---|
| `V22_InaprocOrder/matcher.py` | Tambah `_extract_angka()`, ubah `cocokkan_variasi` return → `tuple[str\|None, list[str]]` |
| `V22_InaprocOrder/order_bot.py` | Update `proses_item()` — handle tuple, tambah status `variasi_tidak_pasti` |
| `V22_InaprocOrder/app.py` | Tambah metric 🔄, update IKON, tambah 3 helper + 3 tombol download |

---

## Task 1: Update matcher.py

**Files:**
- Modify: `V22_InaprocOrder/matcher.py`

### Konteks
`cocokkan_variasi` saat ini return `str | None`. Pass baru (Pass 3) mengekstrak angka dari nama barang dan mencari variasi yang share angka yang sama. Return type berubah jadi **tuple** sehingga `order_bot.py` bisa tahu apakah match itu ambigu.

- Return `(variasi, [])` → match pasti dari Pass 1 atau 2
- Return `(kandidat[0], [semua_kandidat])` → match ambigu dari Pass 3
- Return `(None, [])` → tidak cocok sama sekali

- [ ] **Step 1: Tulis ulang matcher.py**

Ganti seluruh isi `V22_InaprocOrder/matcher.py` dengan:

```python
"""
matcher.py — Cocokkan nama barang dari Excel ke variasi yang tersedia di halaman produk.
"""

import re


def _extract_angka(teks: str) -> set[str]:
    """Extract semua sequence angka dari teks.
    "No. 10-12 isi 20" → {"10", "12", "20"}
    """
    return set(re.findall(r'\d+', teks))


def cocokkan_variasi(nama_barang: str, variasi_tersedia: list[str]) -> tuple[str | None, list[str]]:
    """
    Cari variasi yang paling cocok dengan nama_barang dari Excel.

    Return: tuple (match, kandidat)
      - (variasi, [])           → cocok pasti (Pass 1 atau 2), langsung pakai
      - (kandidat[0], kandidat) → ambigu (Pass 3), masuk keranjang tapi harus dicek
      - (None, [])              → tidak cocok sama sekali, skip

    Contoh:
      "Tinta Printer Hitam", ["Hitam", "Cyan"] → ("Hitam", [])
      "Staples No. 10-12",   ["Max 10", "GW10", "Max 3"] → ("Max 10", ["Max 10", "GW10"])
      "Tinta Warna",         ["Hitam", "Cyan", "Magenta"] → (None, [])
    """
    if not variasi_tersedia:
        return None, []  # produk tanpa variasi → langsung isi qty

    nama_lower = nama_barang.lower()

    # Pass 1: exact substring match
    for variasi in variasi_tersedia:
        if variasi.lower() in nama_lower:
            return variasi, []

    # Pass 2: token-based match (ada irisan kata)
    kata_nama = set(re.findall(r'\w+', nama_lower))
    for variasi in variasi_tersedia:
        kata_variasi = set(re.findall(r'\w+', variasi.lower()))
        if kata_variasi & kata_nama:
            return variasi, []

    # Pass 3: angka overlap — ambigu, tapi tetap pilih kandidat pertama
    angka_nama = _extract_angka(nama_lower)
    if angka_nama:
        kandidat = [
            v for v in variasi_tersedia
            if _extract_angka(v.lower()) & angka_nama
        ]
        if kandidat:
            return kandidat[0], kandidat

    return None, []  # tidak cocok sama sekali


def ada_variasi(variasi_tersedia: list[str]) -> bool:
    """Return True jika produk punya pilihan variasi."""
    return len(variasi_tersedia) > 0
```

- [ ] **Step 2: Verifikasi matcher baru dengan quick test**

Jalankan di terminal:
```
cd "d:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"
"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe" -c "
import sys; sys.path.insert(0, 'V22_InaprocOrder')
from matcher import cocokkan_variasi, _extract_angka

# Test _extract_angka
assert _extract_angka('No. 10-12') == {'10', '12'}, 'FAIL: extract angka'
assert _extract_angka('Max 1213') == {'1213'}, 'FAIL: extract single'
assert _extract_angka('tanpa angka') == set(), 'FAIL: empty'

# Pass 1: exact substring
assert cocokkan_variasi('Tinta Hitam', ['Hitam', 'Cyan']) == ('Hitam', []), 'FAIL: pass1'

# Pass 2: token overlap
assert cocokkan_variasi('Tinta Printer Cyan Epson', ['Hitam', 'Cyan', 'Magenta']) == ('Cyan', []), 'FAIL: pass2'

# Pass 3: angka ambigu
match, kandidat = cocokkan_variasi('Staples No. 10-12 isi 20 kotak', ['Max 10', 'GW10', 'Max 3', 'Max 1213'])
assert match == 'Max 10', f'FAIL: pass3 match={match}'
assert 'Max 10' in kandidat and 'GW10' in kandidat, f'FAIL: kandidat={kandidat}'
assert 'Max 3' not in kandidat, f'FAIL: Max 3 seharusnya tidak masuk, kandidat={kandidat}'

# Tidak cocok
assert cocokkan_variasi('Tinta Warna', ['Hitam', 'Cyan', 'Magenta']) == (None, []), 'FAIL: no match'

# Tanpa variasi
assert cocokkan_variasi('Amplop', []) == (None, []), 'FAIL: empty variasi'

print('Semua test PASS')
"
```

Expected output:
```
Semua test PASS
```

- [ ] **Step 3: Commit**

```bash
cd "d:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"
git add V22_InaprocOrder/matcher.py
git commit -m "feat(v22): matcher pass3 — angka overlap, return tuple (match, kandidat)"
```

---

## Task 2: Update order_bot.py

**Files:**
- Modify: `V22_InaprocOrder/order_bot.py`

### Konteks
`proses_item()` perlu handle return tuple baru dari `cocokkan_variasi`. Jika `kandidat` tidak kosong → item berhasil masuk keranjang tapi status `variasi_tidak_pasti`. Keterangan mencantumkan variasi yang dipilih dan kandidat lainnya.

- [ ] **Step 1: Update import dan proses_item di order_bot.py**

Di [order_bot.py](V22_InaprocOrder/order_bot.py), cari blok ini (sekitar baris 143–153):

```python
            variasi_tersedia = self._ambil_variasi()
            if ada_variasi(variasi_tersedia):
                cocok = cocokkan_variasi(nama, variasi_tersedia)
                if cocok is None:
                    self._log(f"  ⚠️  Variasi tidak cocok. Tersedia: {variasi_tersedia}")
                    return {
                        **base,
                        "status": "variasi_tidak_cocok",
                        "pesan": f"Variasi tidak cocok. Tersedia: {', '.join(variasi_tersedia)}",
                    }
                self._pilih_variasi(cocok)
                self._log(f"  Variasi: {cocok}")
                self._page.wait_for_timeout(800)
```

Ganti dengan:

```python
            variasi_tersedia = self._ambil_variasi()
            status_ambigu = False
            pesan_ambigu = ""

            if ada_variasi(variasi_tersedia):
                cocok, kandidat = cocokkan_variasi(nama, variasi_tersedia)
                if cocok is None:
                    self._log(f"  ⚠️  Variasi tidak cocok. Tersedia: {variasi_tersedia}")
                    return {
                        **base,
                        "status": "variasi_tidak_cocok",
                        "pesan": f"Variasi tidak cocok. Tersedia: {', '.join(variasi_tersedia)}",
                    }
                self._pilih_variasi(cocok)
                self._log(f"  Variasi: {cocok}")
                if kandidat:
                    status_ambigu = True
                    kandidat_lain = [k for k in kandidat if k != cocok]
                    pesan_ambigu = f"Dipilih: {cocok} (ambigu)."
                    if kandidat_lain:
                        pesan_ambigu += f" Kandidat lain: {', '.join(kandidat_lain)}"
                    self._log(f"  🔄 {pesan_ambigu}")
                self._page.wait_for_timeout(800)
```

- [ ] **Step 2: Update return statement berhasil di proses_item**

Cari blok ini (sekitar baris 167–170):

```python
            tambah_btn.click()
            self._page.wait_for_timeout(1500)
            self._log("  ✅ Berhasil")
            return {**base, "status": "berhasil", "pesan": "Berhasil ditambahkan ke keranjang"}
```

Ganti dengan:

```python
            tambah_btn.click()
            self._page.wait_for_timeout(1500)
            if status_ambigu:
                self._log("  🔄 Berhasil (perlu dicek)")
                return {**base, "status": "variasi_tidak_pasti", "pesan": pesan_ambigu}
            self._log("  ✅ Berhasil")
            return {**base, "status": "berhasil", "pesan": "Berhasil ditambahkan ke keranjang"}
```

- [ ] **Step 3: Verifikasi syntax order_bot.py**

```
"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe" -c "
import sys; sys.path.insert(0, 'V22_InaprocOrder')
import order_bot
print('order_bot import OK')
"
```

Expected:
```
order_bot import OK
```

- [ ] **Step 4: Commit**

```bash
git add V22_InaprocOrder/order_bot.py
git commit -m "feat(v22): order_bot handle variasi_tidak_pasti dari matcher pass3"
```

---

## Task 3: Update app.py — Metric + Download

**Files:**
- Modify: `V22_InaprocOrder/app.py`

### Konteks
Tiga perubahan di `app.py`:
1. Update dict `IKON` — tambah `variasi_tidak_pasti`
2. Update section metrics — 5 kolom (tambah 🔄 Perlu Cek)
3. Tambah 3 helper functions + 3 tombol download di Bagian 4

### Helper functions

Tiga fungsi ditulis di bagian atas `app.py` (setelah imports, sebelum `st.set_page_config`):

- `generate_csv(hasil)` → `bytes` via `pandas.DataFrame.to_csv()`
- `generate_excel(hasil)` → `bytes` via `openpyxl.Workbook`
- `generate_pdf(hasil)` → `bytes` via `fitz.Story` (PyMuPDF)

- [ ] **Step 1: Tambah helper functions di app.py**

Di [app.py](V22_InaprocOrder/app.py), cari baris:
```python
st.set_page_config(page_title="V22 Inaproc Order Bot", page_icon="🛒", layout="centered")
```

Sisipkan blok ini **tepat di atasnya**:

```python
import datetime
import io

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
import fitz  # PyMuPDF


_STATUS_LABEL = {
    "berhasil": "Berhasil",
    "variasi_tidak_pasti": "Perlu Cek (ambigu)",
    "variasi_tidak_cocok": "Variasi Tidak Cocok",
    "skip_tidak_aktif": "Skip (tidak aktif)",
    "error": "Error",
    "session_expired": "Session Expired",
}
_STATUS_COLOR_HEX = {
    "berhasil": "D4F4DD",
    "variasi_tidak_pasti": "FFF3CD",
    "variasi_tidak_cocok": "FDE8D8",
    "skip_tidak_aktif": "F0F0F0",
    "error": "FFCCCC",
    "session_expired": "E0E0FF",
}


def _buat_rows(hasil: list[dict]) -> list[dict]:
    return [
        {
            "No": i + 1,
            "Status": _STATUS_LABEL.get(h["status"], h["status"]),
            "Nama Barang": h["nama_barang"].replace("\n", " "),
            "Qty": h["kuantitas"],
            "Keterangan": h["pesan"],
            "Link": h["link"],
        }
        for i, h in enumerate(hasil)
    ]


def generate_csv(hasil: list[dict]) -> bytes:
    df = pd.DataFrame(_buat_rows(hasil))
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    return buf.getvalue()


def generate_excel(hasil: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hasil Order"
    headers = ["No", "Status", "Nama Barang", "Qty", "Keterangan", "Link"]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for i, h in enumerate(hasil):
        row_data = [
            i + 1,
            _STATUS_LABEL.get(h["status"], h["status"]),
            h["nama_barang"].replace("\n", " "),
            h["kuantitas"],
            h["pesan"],
            h["link"],
        ]
        ws.append(row_data)
        color = _STATUS_COLOR_HEX.get(h["status"], "FFFFFF")
        fill = PatternFill("solid", fgColor=color)
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=i + 2, column=col_idx).fill = fill
    col_widths = [6, 22, 45, 6, 55, 55]
    for col_letter, width in zip("ABCDEF", col_widths):
        ws.column_dimensions[col_letter].width = width
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_pdf(hasil: list[dict]) -> bytes:
    STATUS_COLOR_CSS = {
        "berhasil": "#d4f4dd",
        "variasi_tidak_pasti": "#fff3cd",
        "variasi_tidak_cocok": "#fde8d8",
        "skip_tidak_aktif": "#f0f0f0",
        "error": "#ffcccc",
        "session_expired": "#e0e0ff",
    }
    rows_html = ""
    for i, h in enumerate(hasil):
        bg = STATUS_COLOR_CSS.get(h["status"], "#ffffff")
        label = _STATUS_LABEL.get(h["status"], h["status"])
        nama = h["nama_barang"].replace("\n", " ")[:70]
        ket = h["pesan"][:90]
        rows_html += (
            f'<tr style="background-color:{bg}">'
            f'<td style="padding:3px;border:1px solid #ccc;text-align:center">{i+1}</td>'
            f'<td style="padding:3px;border:1px solid #ccc">{label}</td>'
            f'<td style="padding:3px;border:1px solid #ccc">{nama}</td>'
            f'<td style="padding:3px;border:1px solid #ccc;text-align:center">{h["kuantitas"]}</td>'
            f'<td style="padding:3px;border:1px solid #ccc">{ket}</td>'
            f"</tr>\n"
        )
    tgl = datetime.date.today().strftime("%d %B %Y")
    html = f"""<html><body style="font-family:sans-serif;font-size:9px">
<h3 style="margin-bottom:4px">Hasil Order Bot — {tgl}</h3>
<table style="border-collapse:collapse;width:100%">
<tr style="background-color:#4472C4;color:white">
  <th style="padding:4px;border:1px solid #ccc;width:4%">No</th>
  <th style="padding:4px;border:1px solid #ccc;width:18%">Status</th>
  <th style="padding:4px;border:1px solid #ccc;width:38%">Nama Barang</th>
  <th style="padding:4px;border:1px solid #ccc;width:5%">Qty</th>
  <th style="padding:4px;border:1px solid #ccc;width:35%">Keterangan</th>
</tr>
{rows_html}
</table></body></html>"""
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    story = fitz.Story(html)
    mediabox = fitz.paper_rect("a4-l")  # landscape A4
    where = mediabox + (30, 30, -30, -30)
    more = True
    while more:
        device = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(device)
        writer.end_page()
    writer.close()
    return buf.getvalue()
```

- [ ] **Step 2: Update dict IKON di Bagian 4**

Cari baris ini di Bagian 4:
```python
    IKON = {"berhasil":"✅","skip_tidak_aktif":"⏭","variasi_tidak_cocok":"⚠️","error":"❌","session_expired":"🔒"}
```

Ganti dengan:
```python
    IKON = {"berhasil":"✅","variasi_tidak_pasti":"🔄","skip_tidak_aktif":"⏭","variasi_tidak_cocok":"⚠️","error":"❌","session_expired":"🔒"}
```

- [ ] **Step 3: Update metrics — 5 kolom**

Cari blok ini di Bagian 4:
```python
    n_berhasil = sum(1 for h in hasil if h["status"] == "berhasil")
    n_skip     = sum(1 for h in hasil if h["status"] == "skip_tidak_aktif")
    n_variasi  = sum(1 for h in hasil if h["status"] == "variasi_tidak_cocok")
    n_error    = sum(1 for h in hasil if h["status"] == "error")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ Berhasil", n_berhasil)
    c2.metric("⏭ Skip", n_skip)
    c3.metric("⚠️ Variasi ?", n_variasi)
    c4.metric("❌ Error", n_error)
```

Ganti dengan:
```python
    n_berhasil    = sum(1 for h in hasil if h["status"] == "berhasil")
    n_tidak_pasti = sum(1 for h in hasil if h["status"] == "variasi_tidak_pasti")
    n_skip        = sum(1 for h in hasil if h["status"] == "skip_tidak_aktif")
    n_variasi     = sum(1 for h in hasil if h["status"] == "variasi_tidak_cocok")
    n_error       = sum(1 for h in hasil if h["status"] == "error")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("✅ Berhasil", n_berhasil)
    c2.metric("🔄 Perlu Cek", n_tidak_pasti)
    c3.metric("⏭ Skip", n_skip)
    c4.metric("⚠️ Variasi ?", n_variasi)
    c5.metric("❌ Error", n_error)
```

- [ ] **Step 4: Tambah 3 tombol download di Bagian 4**

Cari baris ini di Bagian 4 (sebelum tombol Proses Ulang):
```python
    st.markdown("---")
    if st.button("🔄 Proses Ulang / File Baru"):
```

Sisipkan blok ini **tepat di atasnya**:
```python
    st.markdown("---")
    tgl_str = datetime.date.today().strftime("%Y-%m-%d")
    col_csv, col_xlsx, col_pdf = st.columns(3)
    with col_csv:
        st.download_button(
            "📥 Download CSV",
            data=generate_csv(hasil),
            file_name=f"hasil_order_{tgl_str}.csv",
            mime="text/csv",
        )
    with col_xlsx:
        st.download_button(
            "📥 Download Excel",
            data=generate_excel(hasil),
            file_name=f"hasil_order_{tgl_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_pdf:
        st.download_button(
            "📥 Download PDF",
            data=generate_pdf(hasil),
            file_name=f"hasil_order_{tgl_str}.pdf",
            mime="application/pdf",
        )
```

- [ ] **Step 5: Verifikasi syntax app.py**

```
"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe" -c "
import ast, pathlib
src = pathlib.Path('V22_InaprocOrder/app.py').read_text(encoding='utf-8')
ast.parse(src)
print('app.py syntax OK')
"
```

Expected:
```
app.py syntax OK
```

- [ ] **Step 6: Verifikasi generate functions tanpa browser**

```
"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe" -c "
import sys; sys.path.insert(0, 'V22_InaprocOrder')

# Mock hasil untuk test
hasil_mock = [
    {'status':'berhasil', 'nama_barang':'Kertas A4', 'kuantitas':5, 'pesan':'Berhasil', 'link':'https://katalog.inaproc.id/p/1'},
    {'status':'variasi_tidak_pasti', 'nama_barang':'Staples No. 10-12', 'kuantitas':2, 'pesan':'Dipilih: Max 10 (ambigu). Kandidat lain: GW10', 'link':'https://katalog.inaproc.id/p/2'},
    {'status':'variasi_tidak_cocok', 'nama_barang':'Tinta Warna', 'kuantitas':1, 'pesan':'Variasi tidak cocok', 'link':'https://katalog.inaproc.id/p/3'},
]

# Import helpers langsung dari app.py (tanpa run streamlit)
import importlib.util, types
# Patch streamlit agar tidak crash saat import
import unittest.mock as mock
with mock.patch.dict('sys.modules', {'streamlit': mock.MagicMock()}):
    spec = importlib.util.spec_from_file_location('app', 'V22_InaprocOrder/app.py')
    app = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app)

csv_bytes = app.generate_csv(hasil_mock)
assert len(csv_bytes) > 100, 'CSV kosong'
assert b'Kertas A4' in csv_bytes, 'CSV tidak ada nama barang'
print(f'CSV OK ({len(csv_bytes)} bytes)')

xlsx_bytes = app.generate_excel(hasil_mock)
assert len(xlsx_bytes) > 1000, 'Excel kosong'
print(f'Excel OK ({len(xlsx_bytes)} bytes)')

pdf_bytes = app.generate_pdf(hasil_mock)
assert pdf_bytes[:4] == b'%PDF', f'Bukan PDF valid: {pdf_bytes[:10]}'
print(f'PDF OK ({len(pdf_bytes)} bytes)')

print('Semua generate function OK')
"
```

Expected:
```
CSV OK (... bytes)
Excel OK (... bytes)
PDF OK (... bytes)
Semua generate function OK
```

- [ ] **Step 7: Commit**

```bash
git add V22_InaprocOrder/app.py
git commit -m "feat(v22): tambah metric variasi_tidak_pasti + export CSV/Excel/PDF"
```

---

## Task 4: Smoke Test Manual

- [ ] **Step 1: Jalankan app via bat**

Double-click `V22_InaprocOrder/Jalankan Order Bot.bat` atau:
```
"C:\Users\MSI\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run V22_InaprocOrder/app.py
```

- [ ] **Step 2: Upload Excel pesanan**

Upload file Excel PPK yang ada. Verifikasi tabel item muncul dengan benar.

- [ ] **Step 3: Jalankan proses (opsional — butuh Edge terbuka)**

Jika Edge tersedia: hubungkan dan proses beberapa item. Verifikasi:
- Item dengan variasi pasti → status ✅ `berhasil`
- Item dengan angka overlap → status 🔄 `variasi_tidak_pasti` + keterangan kandidat
- Metrics menampilkan 5 kolom (bukan 4)

- [ ] **Step 4: Test tombol download tanpa proses**

Jika tidak ada Edge, bisa tes tombol download dengan memanipulasi `st.session_state` via browser console, atau cukup percaya hasil Step 6 di Task 3.

- [ ] **Step 5: Final commit + tag versi**

```bash
git add -A
git commit -m "v1.3: matcher pass3 ambigu + export CSV/Excel/PDF"
```

---

## Ringkasan Perubahan Status

| Status | v1.2 | v1.3 |
|---|---|---|
| `berhasil` | ✅ ada | ✅ ada (tidak berubah) |
| `variasi_tidak_pasti` | — | 🔄 **baru** — masuk keranjang, perlu crosscheck |
| `variasi_tidak_cocok` | ⚠️ ada | ⚠️ ada (tidak berubah) |
| `skip_tidak_aktif` | ⏭ ada | ⏭ ada (tidak berubah) |
| `error` | ❌ ada | ❌ ada (tidak berubah) |
