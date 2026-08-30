import pytest

from word_merge import _validate_merge_source_data, _validate_merged_document_text


def _source():
    return {
        "_source_sheet": "list_reviu",
        "Uraian_Pekerjaan_Resiko_K3": "Pekerjaan Struktur - Beton fc 15 Mpa",
        "Resiko_TertinggiFatal_K3": "Tertimpa Material",
        "_dokpil_equipment": [{
            "jenis": "Dump Truck", "kapasitas": "3,5 Ton", "jumlah": "1 Unit",
        }],
        "_dokpil_personnel": [{
            "jabatan": "Petugas K3", "sertifikat": "SKK Petugas K3",
            "pengalaman": "0 Tahun",
        }],
    }


def test_reviu_source_rejects_missing_active_value():
    data = _source()
    data["_dokpil_personnel"][0]["pengalaman"] = "0"
    with pytest.raises(ValueError, match="Pengalaman Kerja Personil 1"):
        _validate_merge_source_data(data)


def test_reviu_source_allows_dash_for_not_applicable_equipment_capacity():
    data = _source()
    data["_dokpil_equipment"][0]["kapasitas"] = "-"
    _validate_merge_source_data(data)


def test_reviu_readback_accepts_source_values_and_rejects_marker():
    data = _source()
    text = " ".join([
        "Pekerjaan Struktur - Beton fc 15 Mpa", "Tertimpa Material",
        "Dump Truck", "3,5 Ton", "1 Unit", "Petugas K3", "SKK Petugas K3",
        "0 Tahun",
    ])
    _validate_merged_document_text(text, data)

    with pytest.raises(ValueError, match="Marker"):
        _validate_merged_document_text(text + " [[KAPASITAS_ALAT]]", data)


def test_reviu_readback_rejects_blank_render():
    with pytest.raises(ValueError, match="Semantic read-back"):
        _validate_merged_document_text("0 0", _source())
