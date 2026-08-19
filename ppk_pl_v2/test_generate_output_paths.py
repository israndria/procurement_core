import os

from ppk_pl_v2 import generate_dokumen_ppk as engine


def _excel_data(stage="UPLOAD AWAL"):
    return {
        "Nama Paket (Lengkap)": "Paket Contoh",
        "Nomor Urut Paket": "13",
        "Tahap Dokumen": stage,
    }


def test_new_output_is_directly_under_package_root(tmp_path):
    output = engine.document_output_dir(_excel_data(), str(tmp_path), mode="generate")

    assert output == os.path.join(
        str(tmp_path), "13. Paket Contoh", "01. Upload Awal"
    )
    assert engine.LEGACY_DOCUMENT_OUTPUT_ROOT not in output


def test_new_contract_output_is_directly_under_package_root(tmp_path):
    output = engine.document_output_dir(
        _excel_data("BERKONTRAK"), str(tmp_path), mode="generate"
    )

    assert output == os.path.join(
        str(tmp_path), "13. Paket Contoh", "02. Berkontrak"
    )


def test_pdf_prefers_new_stage_folder(tmp_path):
    package = tmp_path / "13. Paket Contoh"
    new_stage = package / "01. Upload Awal"
    new_stage.mkdir(parents=True)

    output = engine.document_output_dir(_excel_data(), str(tmp_path), mode="pdf")

    assert output == os.path.normpath(str(new_stage))


def test_pdf_reads_nested_legacy_stage_folder(tmp_path):
    package = tmp_path / "13. Paket Contoh"
    legacy_stage = package / engine.LEGACY_DOCUMENT_OUTPUT_ROOT / "01. Upload Awal"
    legacy_stage.mkdir(parents=True)

    output = engine.document_output_dir(_excel_data(), str(tmp_path), mode="pdf")

    assert output == os.path.normpath(str(legacy_stage))


def test_pdf_falls_back_to_legacy_package_root(tmp_path):
    package = tmp_path / "13. Paket Contoh"
    package.mkdir(parents=True)

    output = engine.document_output_dir(_excel_data(), str(tmp_path), mode="pdf")

    assert output == os.path.normpath(str(package))
