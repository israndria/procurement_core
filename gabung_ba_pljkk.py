"""
gabung_ba_pljkk.py — Gabung BA Utama PLJKK/PLPK + sisipan BA evaluasi/hasil.

Logic:
1. Deteksi file input:
   - BA Utama: BA_PLJKK_*.pdf atau BA_PLPK_*.pdf
   - BA Evaluasi: 7. Berita Acara + Summary Non Tender/5. BA Evaluasi Penawaran PL-*.pdf
   - BA Hasil: 7. Berita Acara + Summary Non Tender/7. BA Hasil Non Tender PL-*.pdf
2. Gunakan pdfplumber untuk mencari:
   - Teks "DAFTAR HADIR PEMBUKTIAN KUALIFIKASI" (occurrence ke-1 & ke-2) -> page index p1 dan p2
   - Teks "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" (occurrence ke-1 & ke-2) -> page index q1 dan q2
3. Gabungkan halaman menggunakan pypdf tanpa menghapus BA utama:
   - Pertahankan blok BA Pembuktian Kualifikasi tervalidasi.
   - Sisipkan BA Evaluasi sebelum daftar hadir pembuktian occurrence ke-2.
   - Duplikasi halaman akhir BA Klarifikasi sebelum sheet 7.2.
   - Sisipkan BA Hasil setelah daftar hadir klarifikasi occurrence ke-1.
4. Output ke "7. Berita Acara + Summary Non Tender/BA_{jenis}_{kode}.pdf"

Usage:
    python gabung_ba_pljkk.py <folder_paket> [PLJKK|PLPK]
"""
import os
import sys
import glob
import ctypes
import shutil
import re
import pdfplumber
from pypdf import PdfReader, PdfWriter


SUBFOLDER = "7. Berita Acara + Summary Non Tender"


def _normalize_jenis(jenis: str) -> str:
    return "PLPK" if str(jenis or "").upper() == "PLPK" else "PLJKK"


def deteksi_file(folder_paket: str, jenis: str = "PLJKK") -> dict:
    """Deteksi file-file input di root folder_paket."""
    jenis = _normalize_jenis(jenis)
    prefix = f"BA_{jenis}_"
    res = {
        'ba_utama': None,
        'ba_pembuktian': None,
        'ba_eval': None,
        'ba_hasil': None,
        'kode': None,
        'err': None
    }
    
    # 1. BA Utama sesuai konteks workflow (PLJKK atau PLPK).
    ba_utama_pattern = os.path.join(folder_paket, f"{prefix}*.pdf")
    ba_utama_files = glob.glob(ba_utama_pattern)
    if not ba_utama_files:
        res['err'] = f"File {prefix}*.pdf tidak ditemukan di root folder paket."
        return res
    
    # Pilih yang terbaru jika ada lebih dari 1
    ba_utama_files_sorted = sorted(ba_utama_files, key=os.path.getmtime, reverse=True)
    res['ba_utama'] = ba_utama_files_sorted[0]
    # Jika ada BA sebelumnya, gunakan blok Pembuktian Kualifikasi tervalidasi
    # dari file itu. Ekspor Word terbaru tetap menjadi sumber halaman lainnya.
    if len(ba_utama_files_sorted) > 1:
        res['ba_pembuktian'] = ba_utama_files_sorted[1]
    
    # Ekstrak kode dari nama file (BA_{jenis}_{kode}.pdf)
    base_name = os.path.basename(res['ba_utama'])
    name_no_ext, _ = os.path.splitext(base_name)
    if name_no_ext.startswith(prefix):
        res['kode'] = name_no_ext[len(prefix):]
    else:
        res['kode'] = "FULL"
        
    # 2. BA Evaluasi berada di subfolder output bersama BA Hasil.
    # Jangan cari di root folder paket; struktur paket PL menyimpan keduanya
    # di "7. Berita Acara + Summary Non Tender".
    ba_eval_pattern = os.path.join(
        folder_paket, SUBFOLDER, "5. BA Evaluasi Penawaran PL-*.pdf"
    )
    ba_eval_files = glob.glob(ba_eval_pattern)
    if ba_eval_files:
        res['ba_eval'] = sorted(ba_eval_files, key=os.path.getmtime, reverse=True)[0]
        
    # 3. BA Hasil di subfolder yang sama.
    ba_hasil_pattern = os.path.join(
        folder_paket, SUBFOLDER, "7. BA Hasil Non Tender PL-*.pdf"
    )
    ba_hasil_files = glob.glob(ba_hasil_pattern)
    if ba_hasil_files:
        res['ba_hasil'] = sorted(ba_hasil_files, key=os.path.getmtime, reverse=True)[0]
        
    return res


def cari_halaman_sisipan(pdf_path: str) -> tuple:
    """
    Mencari indeks halaman penanda.
    - "DAFTAR HADIR PEMBUKTIAN KUALIFIKASI" -> p1 (occurrence ke-1), p2 (occurrence ke-2)
    - "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" -> q1 (occurrence ke-1), q2 (occurrence ke-2)
    Returns: (p1, p2, q1, q2)
    """
    p1 = p2 = q1 = q2 = None
    p_indices = []
    q_indices = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for idx, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if "DAFTAR HADIR PEMBUKTIAN KUALIFIKASI" in text:
                    p_indices.append(idx)
                if "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" in text:
                    q_indices.append(idx)
                    
        if len(p_indices) >= 2:
            p1, p2 = p_indices[0], p_indices[1]
        elif len(p_indices) == 1:
            p1 = p_indices[0]
            
        if len(q_indices) >= 2:
            q1, q2 = q_indices[0], q_indices[1]
        elif len(q_indices) == 1:
            q1 = q_indices[0]
            
    except Exception as e:
        print(f"[WARN] Gagal membaca teks PDF: {e}")
        
    return p1, p2, q1, q2


def cari_halaman_ttd_penyedia(pdf_path: str) -> list[int]:
    """Cari halaman tanda tangan utama PLPK berdasarkan marker isi aktual."""
    indices = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for index, page in enumerate(pdf.pages):
                text = re.sub(r"\s+", " ", page.extract_text() or "").upper()
                has_provider_label = bool(
                    re.search(r"DIREKTUR\s*/\s*PIMPINAN", text)
                )
                has_official_label = "PEJABAT PENGADAAN" in text
                has_main_ba_closing = (
                    "DEMIKIAN BERITA ACARA KLARIFIKASI DAN NEGOSIASI INI" in text
                )
                is_attachment = "LAMPIRAN I BERITA ACARA" in text
                if (
                    has_provider_label
                    and has_official_label
                    and has_main_ba_closing
                    and not is_attachment
                ):
                    indices.append(index)
    except Exception as exc:
        raise RuntimeError(f"Gagal mendeteksi halaman tanda tangan PLPK: {exc}") from exc
    return indices


def _page_text(page) -> str:
    try:
        return re.sub(r"\s+", " ", page.extract_text() or "").strip()
    except Exception:
        return ""


def _pages_equivalent(left, right) -> bool:
    """Bandingkan halaman untuk mencegah duplikasi signature berulang."""
    left_text = _page_text(left)
    right_text = _page_text(right)
    if not left_text or left_text != right_text:
        return False
    try:
        return (
            float(left.mediabox.width) == float(right.mediabox.width)
            and float(left.mediabox.height) == float(right.mediabox.height)
        )
    except Exception:
        return True


def ensure_plpk_provider_signature_copy(pdf_path: str) -> bool:
    """Pastikan halaman signature penyedia PLPK muncul dua kali berurutan.

    Posisi ditentukan dari marker teks aktual, bukan nomor halaman. Operasi
    idempotent; marker tidak ditemukan berarti gagal tertutup dan PDF tidak
    diubah.
    """
    reader = PdfReader(pdf_path)
    signature_pages = cari_halaman_ttd_penyedia(pdf_path)
    if not signature_pages:
        raise RuntimeError(
            "Halaman tanda tangan penyedia PLPK tidak terdeteksi; PDF tidak diubah."
        )
    target = signature_pages[0]
    if target + 1 < len(reader.pages) and _pages_equivalent(
        reader.pages[target], reader.pages[target + 1]
    ):
        return False

    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        writer.add_page(page)
        if index == target:
            writer.add_page(page)

    temporary = f"{pdf_path}.signature-copy.tmp"
    try:
        with open(temporary, "wb") as output:
            writer.write(output)
        os.replace(temporary, pdf_path)
    finally:
        try:
            if os.path.exists(temporary):
                os.remove(temporary)
        except OSError:
            pass
    return True


def gabung(folder_paket: str, jenis: str = "PLJKK") -> dict:
    jenis = _normalize_jenis(jenis)
    files = deteksi_file(folder_paket, jenis)
    if files['err']:
        return {'ok': False, 'output': '', 'pesan': files['err'], 'warning': None}
        
    ba_utama = files['ba_utama']
    ba_pembuktian = files['ba_pembuktian']
    ba_eval = files['ba_eval']
    ba_hasil = files['ba_hasil']
    kode = files['kode']
    
    # Buat output folder
    out_dir = os.path.join(folder_paket, SUBFOLDER)
    os.makedirs(out_dir, exist_ok=True)
    
    output_filename = f"BA_{jenis}_{kode}.pdf"
    output_path = os.path.join(out_dir, output_filename)

    # Tanpa BA evaluasi/hasil dan tanpa BA lama, tidak ada halaman yang perlu
    # dicari atau disisipkan. Salin BA utama langsung ke output final.
    if not ba_eval and not ba_hasil and not ba_pembuktian:
        try:
            shutil.copy2(ba_utama, output_path)
            return {
                'ok': True,
                'output': output_path,
                'pesan': f"BA_{jenis} berhasil disalin: {os.path.basename(output_path)}",
                'warning': None,
            }
        except Exception as e:
            return {'ok': False, 'output': '', 'pesan': f"Gagal menyalin PDF: {e}", 'warning': None}
    
    warning_msgs = []
    if not ba_eval:
        warning_msgs.append("File '7. Berita Acara + Summary Non Tender/5. BA Evaluasi Penawaran PL-*.pdf' tidak ditemukan, skip sisipan evaluasi.")
    if not ba_hasil:
        warning_msgs.append("File '7. Berita Acara + Summary Non Tender/7. BA Hasil Non Tender PL-*.pdf' tidak ditemukan, skip sisipan hasil non tender.")
        
    try:
        rdr_utama = PdfReader(ba_utama)
        n_utama = len(rdr_utama.pages)
        rdr_pembuktian = None
        
        # Cari penanda halaman
        p1, p2, q1, q2 = cari_halaman_sisipan(ba_utama)

        # Blok pembuktian dapat berasal dari BA tervalidasi sebelumnya, hanya
        # jika anchor-nya sama persis dengan BA utama terbaru.
        if ba_pembuktian and p1 is not None and p2 is not None:
            old_p1, old_p2, _, _ = cari_halaman_sisipan(ba_pembuktian)
            if (old_p1, old_p2) == (p1, p2):
                candidate = PdfReader(ba_pembuktian)
                if len(candidate.pages) > p2:
                    rdr_pembuktian = candidate
        
        # Dokumen final mempertahankan seluruh BA utama. Summary hanya disisipkan:
        # - Evaluasi sebelum daftar hadir pembuktian occurrence ke-2;
        # - Hasil setelah daftar hadir klarifikasi occurrence ke-1.
        # Dua halaman tidak boleh menggantikan blok BA utama.
        if ba_eval and (p1 is None or p2 is None):
            warning_msgs.append("Penanda 'DAFTAR HADIR PEMBUKTIAN KUALIFIKASI' tidak lengkap, skip sisipan evaluasi.")
            ba_eval = None
        if ba_hasil and (q1 is None or q2 is None):
            warning_msgs.append("Penanda 'DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI' tidak lengkap, skip sisipan hasil.")
            ba_hasil = None

        # PLPK perlu dua copy halaman tanda tangan penyedia. Deteksi marker
        # aktual agar tidak bergantung pada nomor halaman atau layout paket.
        # PLJKK mempertahankan aturan lama: halaman sebelum sisipan landscape.
        duplicate_after = None
        if jenis == "PLPK":
            signature_pages = cari_halaman_ttd_penyedia(ba_utama)
            if signature_pages:
                duplicate_after = signature_pages[0]
        else:
            for idx, page in enumerate(rdr_utama.pages):
                box = page.mediabox
                if float(box.width) > float(box.height) and idx > 0:
                    duplicate_after = idx - 1
                    break

        insert_before = {}
        insert_after = {}
        if ba_eval:
            insert_before[p2] = ba_eval
        if ba_hasil:
            insert_after[q1] = ba_hasil

        writer = PdfWriter()
        for idx, page in enumerate(rdr_utama.pages):
            if idx in insert_before:
                for summary_page in PdfReader(insert_before[idx]).pages:
                    writer.add_page(summary_page)

            # Halaman di antara dua daftar hadir Pembuktian Kualifikasi adalah
            # blok tersertifikasi; pakai versi BA lama bila tersedia.
            source_page = (
                rdr_pembuktian.pages[idx]
                if rdr_pembuktian is not None and p1 < idx < p2
                else page
            )
            writer.add_page(source_page)

            if idx == duplicate_after and not (
                idx + 1 < n_utama
                and _pages_equivalent(page, rdr_utama.pages[idx + 1])
            ):
                writer.add_page(page)

            if idx in insert_after:
                for summary_page in PdfReader(insert_after[idx]).pages:
                    writer.add_page(summary_page)
                
        # Tulis output dengan retry logic PermissionError
        path = output_path
        for attempt in range(5):
            try:
                with open(path, 'wb') as f:
                    writer.write(f)
                break
            except PermissionError:
                if attempt == 4:
                    raise
                base, ext = os.path.splitext(output_path)
                path = f"{base}_v{attempt + 2}{ext}"
                
        warning_str = "\n".join(warning_msgs) if warning_msgs else None
        return {
            'ok': True,
            'output': path,
            'pesan': f"BA_{jenis} berhasil digabung: {os.path.basename(path)} ({len(writer.pages)} halaman)",
            'warning': warning_str
        }
        
    except Exception as e:
        return {'ok': False, 'output': '', 'pesan': f"Gagal menggabungkan PDF: {e}", 'warning': None}


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: python gabung_ba_pljkk.py <folder_paket> [PLJKK|PLPK]")
        sys.exit(1)
        
    folder_paket = os.path.abspath(args[0])
    if not os.path.isdir(folder_paket):
        print(f"[ERROR] Folder tidak ditemukan: {folder_paket}")
        try:
            ctypes.windll.user32.MessageBoxW(0, f"Folder tidak ditemukan:\n{folder_paket}", "Gabung BA PLJKK - Error", 0x10)
        except Exception:
            pass
        sys.exit(1)
        
    jenis = args[1] if len(args) > 1 else "PLJKK"
    result = gabung(folder_paket, jenis)
    
    if result['ok']:
        print(f"[OK] {result['pesan']}")
        if result['warning']:
            print(f"[WARN] {result['warning']}")
            
        msg = result['pesan']
        if result['warning']:
            msg += f"\n\n⚠️ Peringatan:\n{result['warning']}"
            
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, f"Gabung BA {_normalize_jenis(jenis)} - Selesai", 0x40)
        except Exception:
            pass
            
        if os.path.exists(result['output']):
            os.startfile(result['output'])
    else:
        print(f"[ERROR] {result['pesan']}")
        try:
            ctypes.windll.user32.MessageBoxW(0, result['pesan'], "Gabung BA PLJKK - Error", 0x10)
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
