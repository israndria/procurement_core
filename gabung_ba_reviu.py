"""
gabung_ba_reviu.py — Gabung BA Reviu Lengkap dari 2 file PDF.

Flow urutan halaman:
  Hal 1-2   : hasil_scan.pdf halaman 1-2 (lembar ttd basah + absensi awal)
  Hal 3+    : isi_reviu.pdf semua halaman
  Hal akhir : hasil_scan.pdf halaman 3 (lembar absensi/ttd akhir)

Output: 6. BA Reviu Lengkap/BA_REVIU_FULL_{nama_paket}.pdf

Usage:
    python gabung_ba_reviu.py <folder_paket>
    python gabung_ba_reviu.py <folder_paket> --dry-run
"""
import os
import sys
import re
import glob


SUBFOLDER = "6. BA Reviu Lengkap"
OUTPUT_PREFIX = "BA_REVIU_FULL_"


def _safe_filename(s: str, max_len: int = 80) -> str:
    s = re.sub(r'[<>:"/\\|?*]', '', str(s)).strip()
    return s[:max_len] if s else "Dokumen"


def _nama_paket_dari_folder(folder_paket: str) -> str:
    """Ekstrak nama paket dari nama folder paket."""
    nama = os.path.basename(folder_paket)
    # Buang prefix "N. PLJKK - " atau "N. PLPK - " atau "NN. Pokja NNN"
    nama = re.sub(r'^\d+\.\s*(PLJKK|PLPK)\s*-\s*', '', nama).strip()
    nama = re.sub(r'\s*\(PL\s*-?\s*Ulang\)\s*$', '', nama, flags=re.IGNORECASE).strip()
    return nama


def deteksi_file(subfolder_path: str) -> dict:
    """
    Deteksi file scan dan Isi Reviu di subfolder.
    Return: {
        'isi_reviu': path atau None,
        'scan': path atau None,
        'scan_candidates': [list paths kalau >1],
        'warning': pesan warning atau None
    }
    """
    if not os.path.isdir(subfolder_path):
        return {'isi_reviu': None, 'scan': None, 'scan_candidates': [], 'warning': f"Folder tidak ditemukan: {subfolder_path}"}

    pdfs = [f for f in os.listdir(subfolder_path) if f.lower().endswith('.pdf')]

    isi_reviu_path = None
    scan_candidates = []

    for f in pdfs:
        full = os.path.join(subfolder_path, f)
        if f.startswith("Isi_Reviu_DPP_") or f.startswith("Isi_Reviu_"):
            isi_reviu_path = full
        else:
            scan_candidates.append(full)

    warning = None
    scan_path = None

    if len(scan_candidates) == 0:
        warning = "File scan tidak ditemukan di folder."
    elif len(scan_candidates) == 1:
        scan_path = scan_candidates[0]
    else:
        # Lebih dari 1 — ambil yang terbaru
        scan_candidates_sorted = sorted(scan_candidates, key=os.path.getmtime, reverse=True)
        scan_path = scan_candidates_sorted[0]
        names = [os.path.basename(p) for p in scan_candidates_sorted]
        warning = f"Ada {len(scan_candidates)} file scan: {', '.join(names)}. Dipakai: {os.path.basename(scan_path)} (terbaru)."

    return {
        'isi_reviu': isi_reviu_path,
        'scan': scan_path,
        'scan_candidates': scan_candidates,
        'warning': warning
    }


def gabung(folder_paket: str, dry_run: bool = False) -> dict:
    """
    Gabung BA Reviu Lengkap.
    Return: {'ok': bool, 'output': path, 'pesan': str, 'warning': str}
    """
    subfolder_path = os.path.join(folder_paket, SUBFOLDER)
    files = deteksi_file(subfolder_path)

    if not files['isi_reviu']:
        return {'ok': False, 'output': '', 'pesan': "File Isi_Reviu_DPP_*.pdf tidak ditemukan di folder '6. BA Reviu Lengkap'.", 'warning': files['warning']}

    if not files['scan']:
        return {'ok': False, 'output': '', 'pesan': files['warning'] or "File scan tidak ditemukan.", 'warning': files['warning']}

    nama_paket = _nama_paket_dari_folder(folder_paket)
    output_name = f"{OUTPUT_PREFIX}{_safe_filename(nama_paket)}.pdf"
    output_path = os.path.join(subfolder_path, output_name)

    if dry_run:
        return {
            'ok': True,
            'output': output_path,
            'pesan': f"[DRY RUN] Akan gabung:\n  Scan: {os.path.basename(files['scan'])}\n  Isi Reviu: {os.path.basename(files['isi_reviu'])}\n  Output: {output_name}",
            'warning': files['warning']
        }

    try:
        from pypdf import PdfReader, PdfWriter

        rdr_scan = PdfReader(files['scan'])
        rdr_reviu = PdfReader(files['isi_reviu'])

        n_scan = len(rdr_scan.pages)
        n_reviu = len(rdr_reviu.pages)

        writer = PdfWriter()

        # Hal 1-2: scan halaman 1 dan 2 (index 0 dan 1)
        for i in range(min(2, n_scan)):
            writer.add_page(rdr_scan.pages[i])

        # Hal 3+: semua halaman Isi Reviu
        for i in range(n_reviu):
            writer.add_page(rdr_reviu.pages[i])

        # Hal akhir: scan halaman 3 (index 2), kalau ada
        if n_scan >= 3:
            writer.add_page(rdr_scan.pages[2])

        # Tulis output (fallback suffix kalau locked)
        path = output_path
        for attempt in range(5):
            try:
                with open(path, 'wb') as f:
                    writer.write(f)
                break
            except PermissionError:
                base, ext = os.path.splitext(output_path)
                path = f"{base}_v{attempt + 2}{ext}"

        return {
            'ok': True,
            'output': path,
            'pesan': f"BA_REVIU_FULL berhasil: {os.path.basename(path)} ({len(writer.pages)} halaman)",
            'warning': files['warning']
        }

    except Exception as e:
        return {'ok': False, 'output': '', 'pesan': str(e), 'warning': files['warning']}


def main():
    args = sys.argv[1:]
    dry_run = '--dry-run' in args
    args = [a for a in args if not a.startswith('--')]

    if not args:
        print("Usage: python gabung_ba_reviu.py <folder_paket> [--dry-run]")
        sys.exit(1)

    folder_paket = os.path.abspath(args[0])
    if not os.path.isdir(folder_paket):
        print(f"[ERROR] Folder tidak ditemukan: {folder_paket}")
        sys.exit(1)

    result = gabung(folder_paket, dry_run=dry_run)

    if result['warning']:
        print(f"[WARN] {result['warning']}")
    if result['ok']:
        print(f"[OK] {result['pesan']}")
        try:
            import ctypes
            msg = result['pesan']
            if result.get('warning'):
                msg += f'\n\n⚠️ {result["warning"]}'
            ctypes.windll.user32.MessageBoxW(0, msg, 'Gabung BA Reviu - Selesai', 0x40)
        except Exception:
            pass
        # Buka file hasil
        if not dry_run and os.path.exists(result['output']):
            os.startfile(result['output'])
    else:
        print(f"[ERROR] {result['pesan']}")
        # Tampilkan popup error
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, result['pesan'], "Gabung BA Reviu - Error", 0x10)
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
