"""
gabung_ba_pljkk.py — Gabung BA Utama PLJKK + Sisipan BA Evaluasi & BA Hasil Non Tender.

Logic:
1. Deteksi file input:
   - BA Utama: BA_PLJKK_*.pdf
   - BA Evaluasi: 5. BA Evaluasi Penawaran PL-*.pdf
   - BA Hasil: 7. BA Hasil Non Tender PL-*.pdf
2. Gunakan pdfplumber untuk mencari:
   - Teks "DAFTAR HADIR PEMBUKTIAN KUALIFIKASI" (occurrence ke-1 & ke-2) -> page index p1 dan p2
   - Teks "DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI" (occurrence ke-1 & ke-2) -> page index q1 dan q2
3. Gabungkan halaman menggunakan pypdf:
   - BA Utama 0 s/d p1 (inclusive)
   - BA Evaluasi (semua hal)
   - BA Utama p2 s/d q1 (inclusive)
   - BA Hasil (semua hal)
   - BA Utama q2 s/d akhir
4. Output ke "7. Berita Acara + Summary Non Tender/BA_PLJKK_{kode}.pdf"

Usage:
    python gabung_ba_pljkk.py <folder_paket>
"""
import os
import sys
import glob
import ctypes
import pdfplumber
from pypdf import PdfReader, PdfWriter


SUBFOLDER = "7. Berita Acara + Summary Non Tender"


def deteksi_file(folder_paket: str) -> dict:
    """Deteksi file-file input di root folder_paket."""
    res = {
        'ba_utama': None,
        'ba_eval': None,
        'ba_hasil': None,
        'kode': None,
        'err': None
    }
    
    # 1. BA Utama (BA_PLJKK_*.pdf)
    ba_utama_pattern = os.path.join(folder_paket, "BA_PLJKK_*.pdf")
    ba_utama_files = glob.glob(ba_utama_pattern)
    if not ba_utama_files:
        res['err'] = "File BA_PLJKK_*.pdf tidak ditemukan di root folder paket."
        return res
    
    # Pilih yang terbaru jika ada lebih dari 1
    ba_utama_files_sorted = sorted(ba_utama_files, key=os.path.getmtime, reverse=True)
    res['ba_utama'] = ba_utama_files_sorted[0]
    
    # Ekstrak kode dari nama file (BA_PLJKK_{kode}.pdf)
    base_name = os.path.basename(res['ba_utama'])
    name_no_ext, _ = os.path.splitext(base_name)
    if name_no_ext.startswith("BA_PLJKK_"):
        res['kode'] = name_no_ext[len("BA_PLJKK_"):]
    else:
        res['kode'] = "FULL"
        
    # 2. BA Evaluasi (5. BA Evaluasi Penawaran PL-*.pdf)
    ba_eval_pattern = os.path.join(folder_paket, "5. BA Evaluasi Penawaran PL-*.pdf")
    ba_eval_files = glob.glob(ba_eval_pattern)
    if ba_eval_files:
        res['ba_eval'] = sorted(ba_eval_files, key=os.path.getmtime, reverse=True)[0]
        
    # 3. BA Hasil (7. BA Hasil Non Tender PL-*.pdf)
    ba_hasil_pattern = os.path.join(folder_paket, "7. BA Hasil Non Tender PL-*.pdf")
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


def gabung(folder_paket: str) -> dict:
    files = deteksi_file(folder_paket)
    if files['err']:
        return {'ok': False, 'output': '', 'pesan': files['err'], 'warning': None}
        
    ba_utama = files['ba_utama']
    ba_eval = files['ba_eval']
    ba_hasil = files['ba_hasil']
    kode = files['kode']
    
    # Buat output folder
    out_dir = os.path.join(folder_paket, SUBFOLDER)
    os.makedirs(out_dir, exist_ok=True)
    
    output_filename = f"BA_PLJKK_{kode}.pdf"
    output_path = os.path.join(out_dir, output_filename)
    
    warning_msgs = []
    if not ba_eval:
        warning_msgs.append("File '5. BA Evaluasi Penawaran PL-*.pdf' tidak ditemukan, skip sisipan evaluasi.")
    if not ba_hasil:
        warning_msgs.append("File '7. BA Hasil Non Tender PL-*.pdf' tidak ditemukan, skip sisipan hasil non tender.")
        
    try:
        rdr_utama = PdfReader(ba_utama)
        n_utama = len(rdr_utama.pages)
        
        # Cari penanda halaman
        p1, p2, q1, q2 = cari_halaman_sisipan(ba_utama)
        
        # Log/Warning jika penanda tidak lengkap
        if ba_eval and (p1 is None or p2 is None):
            warning_msgs.append("Penanda 'DAFTAR HADIR PEMBUKTIAN KUALIFIKASI' tidak lengkap, skip sisipan evaluasi.")
            ba_eval = None
            
        if ba_hasil and (q1 is None or q2 is None):
            warning_msgs.append("Penanda 'DAFTAR HADIR KLARIFIKASI DAN NEGOSIASI' tidak lengkap, skip sisipan hasil.")
            ba_hasil = None
            
        writer = PdfWriter()
        
        # Kondisi 1: Kedua sisipan aktif
        if ba_eval and ba_hasil:
            # 0 s/d p1 (inclusive)
            for idx in range(0, p1 + 1):
                writer.add_page(rdr_utama.pages[idx])
            # Sisipkan BA Evaluasi
            rdr_eval = PdfReader(ba_eval)
            for page in rdr_eval.pages:
                writer.add_page(page)
            # p2 s/d q1 (inclusive)
            for idx in range(p2, q1 + 1):
                writer.add_page(rdr_utama.pages[idx])
            # Sisipkan BA Hasil
            rdr_hasil = PdfReader(ba_hasil)
            for page in rdr_hasil.pages:
                writer.add_page(page)
            # q2 s/d akhir
            for idx in range(q2, n_utama):
                writer.add_page(rdr_utama.pages[idx])
                
        # Kondisi 2: Hanya BA Evaluasi aktif
        elif ba_eval:
            for idx in range(0, p1 + 1):
                writer.add_page(rdr_utama.pages[idx])
            rdr_eval = PdfReader(ba_eval)
            for page in rdr_eval.pages:
                writer.add_page(page)
            for idx in range(p2, n_utama):
                writer.add_page(rdr_utama.pages[idx])
                
        # Kondisi 3: Hanya BA Hasil aktif
        elif ba_hasil:
            for idx in range(0, q1 + 1):
                writer.add_page(rdr_utama.pages[idx])
            rdr_hasil = PdfReader(ba_hasil)
            for page in rdr_hasil.pages:
                writer.add_page(page)
            for idx in range(q2, n_utama):
                writer.add_page(rdr_utama.pages[idx])
                
        # Kondisi 4: Tanpa sisipan (output BA Utama saja)
        else:
            for page in rdr_utama.pages:
                writer.add_page(page)
                
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
            'pesan': f"BA_PLJKK berhasil digabung: {os.path.basename(path)} ({len(writer.pages)} halaman)",
            'warning': warning_str
        }
        
    except Exception as e:
        return {'ok': False, 'output': '', 'pesan': f"Gagal menggabungkan PDF: {e}", 'warning': None}


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: python gabung_ba_pljkk.py <folder_paket>")
        sys.exit(1)
        
    folder_paket = os.path.abspath(args[0])
    if not os.path.isdir(folder_paket):
        print(f"[ERROR] Folder tidak ditemukan: {folder_paket}")
        try:
            ctypes.windll.user32.MessageBoxW(0, f"Folder tidak ditemukan:\n{folder_paket}", "Gabung BA PLJKK - Error", 0x10)
        except Exception:
            pass
        sys.exit(1)
        
    result = gabung(folder_paket)
    
    if result['ok']:
        print(f"[OK] {result['pesan']}")
        if result['warning']:
            print(f"[WARN] {result['warning']}")
            
        msg = result['pesan']
        if result['warning']:
            msg += f"\n\n⚠️ Peringatan:\n{result['warning']}"
            
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, "Gabung BA PLJKK - Selesai", 0x40)
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
