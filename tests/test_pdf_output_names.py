from pathlib import Path

from gabung_ba_reviu import gabung
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
