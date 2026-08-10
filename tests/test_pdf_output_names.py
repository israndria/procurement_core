from pathlib import Path

from gabung_ba_reviu import gabung
from pypdf import PdfReader, PdfWriter
from word_merge import _pdf_output_suffix


def test_tender_pdf_uses_pokja_marker_over_package_name():
    assert _pdf_output_suffix("POKJA_045", "Paket Jalan Sangat Panjang") == "045"
    assert _pdf_output_suffix("", "Paket Jalan") == "Paket Jalan"


def test_gabung_ba_reviu_uses_pokja_marker(tmp_path):
    subfolder = tmp_path / "6. BA Reviu Lengkap"
    subfolder.mkdir()
    (subfolder / "scan.pdf").touch()
    (subfolder / "Isi_Reviu_DPP_045.pdf").touch()

    result = gabung(str(tmp_path), dry_run=True, suffix="POKJA_045")

    assert result["ok"] is True
    assert Path(result["output"]).name == "BA_REVIU_FULL_045.pdf"


def _write_blank_pages(path, widths):
    writer = PdfWriter()
    for width in widths:
        writer.add_blank_page(width=width, height=100)
    with path.open("wb") as output:
        writer.write(output)


def test_gabung_ba_reviu_keeps_scan_pages_three_and_four_in_order(tmp_path):
    subfolder = tmp_path / "6. BA Reviu Lengkap"
    subfolder.mkdir()
    _write_blank_pages(subfolder / "BA Acara Reviu.pdf", [100, 101, 102, 103])
    _write_blank_pages(subfolder / "Isi Reviu.pdf", [200, 201])

    result = gabung(str(tmp_path), suffix="POKJA_041")

    assert result["ok"] is True
    reader = PdfReader(result["output"])
    assert [float(page.mediabox.width) for page in reader.pages] == [
        100.0, 101.0, 200.0, 201.0, 102.0, 103.0
    ]
