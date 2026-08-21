import json
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from pl_snapshot_revision import (
    FORMULA_ADDRESSES,
    IMMUTABLE_ADDRESSES,
    CELL_METADATA,
    SnapshotError,
    WHITELIST_ADDRESSES,
    compare_snapshots,
    parse_snapshot,
    promote_proposal,
    seed_proposal,
    validate_snapshot,
)


def _write_snapshot(path: Path, *, code: str = "11000000000", overrides=None, omit=()):
    overrides = overrides or {}
    root = ET.Element(
        "snapshot",
        schema="pokja-pl-master-data",
        version="1",
        saved_at="2026-08-21 00:00:00",
        workbook="test.xlsm",
        kode_paket=code,
    )
    cells = ET.SubElement(root, "cells")
    for address in sorted(WHITELIST_ADDRESSES):
        if address in omit:
            continue
        default = ("formula", f"={address}") if address in FORMULA_ADDRESSES else ("text", f"old-{address}")
        cell_type, text = overrides.get(address, default)
        node = ET.SubElement(cells, "cell", address=address, type=cell_type)
        node.text = text
    path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))


def test_whitelist_contains_date_inputs_but_not_static_labels():
    assert {"H8", "H9", "H10", "H11", "I8", "I9", "I10"} <= WHITELIST_ADDRESSES
    assert not {"F8", "F9", "F10", "F11"} & WHITELIST_ADDRESSES
    assert FORMULA_ADDRESSES <= WHITELIST_ADDRESSES
    assert {"C11", "C12", "C20", "C22", "C24", "C26"} <= FORMULA_ADDRESSES
    assert all(not CELL_METADATA[address]["editable"] for address in FORMULA_ADDRESSES)
    assert IMMUTABLE_ADDRESSES <= WHITELIST_ADDRESSES


def test_compare_reports_semantic_cell_change(tmp_path):
    baseline = tmp_path / "baseline.xml"
    candidate = tmp_path / "candidate.xml"
    _write_snapshot(baseline, overrides={"C15": ("text", "180 Hari")})
    _write_snapshot(candidate, overrides={"C15": ("text", "90 Hari")})

    changes = compare_snapshots(baseline, candidate)

    assert changes == [
        {
            "address": "C15",
            "key": "jangka_waktu",
            "label": "Jangka Waktu",
            "section": "Kontrak",
            "editable": True,
            "status": "changed",
            "old_type": "text",
            "old_value": "180 Hari",
            "new_type": "text",
            "new_value": "90 Hari",
        }
    ]


def test_promote_keeps_baseline_creates_backup_and_audit(tmp_path):
    baseline = tmp_path / "input_data_baseline.xml"
    current = tmp_path / "input_data_snapshot.xml"
    proposal = tmp_path / "input_data_proposal.xml"
    audit = tmp_path / "input_data_audit.jsonl"
    _write_snapshot(baseline, overrides={"C15": ("text", "180 Hari")})
    _write_snapshot(current, overrides={"C15": ("text", "180 Hari")})
    _write_snapshot(proposal, overrides={"C15": ("text", "90 Hari")})
    baseline_before = baseline.read_bytes()

    result = promote_proposal(proposal, current, baseline_path=baseline, audit_path=audit)

    assert result["applied"] is True
    assert parse_snapshot(current).cells["C15"].text == "90 Hari"
    assert baseline.read_bytes() == baseline_before
    assert list(tmp_path.glob("input_data_snapshot.bak-*.xml"))
    record = json.loads(audit.read_text(encoding="utf-8"))
    assert record["changed_cells"] == ["C15"]


def test_promote_rejects_incomplete_or_read_only_proposal(tmp_path):
    current = tmp_path / "current.xml"
    incomplete = tmp_path / "incomplete.xml"
    readonly = tmp_path / "readonly.xml"
    _write_snapshot(current)
    _write_snapshot(incomplete, omit={"C14"})
    _write_snapshot(readonly, overrides={"H10": ("formula", "=1")})

    with pytest.raises(SnapshotError, match="tidak lengkap"):
        promote_proposal(incomplete, current)
    with pytest.raises(SnapshotError, match="read-only"):
        promote_proposal(readonly, current)


def test_seed_proposal_records_source_hash_and_rejects_stale_current(tmp_path):
    current = tmp_path / "input_data_snapshot.xml"
    proposal = tmp_path / "input_data_proposal.xml"
    _write_snapshot(current)

    result = seed_proposal(current, proposal)
    assert result["source_sha256"] == parse_snapshot(current).sha256

    _write_snapshot(current, overrides={"C15": ("text", "90 Hari")})
    with pytest.raises(SnapshotError, match="current lama"):
        promote_proposal(proposal, current)


def test_promote_rejects_formula_on_input_field(tmp_path):
    current = tmp_path / "current.xml"
    proposal = tmp_path / "proposal.xml"
    _write_snapshot(current)
    _write_snapshot(proposal, overrides={"C15": ("formula", "=1")})

    with pytest.raises(SnapshotError, match="menambah formula"):
        promote_proposal(proposal, current)


def test_snapshot_rejects_dtd_entity(tmp_path):
    path = tmp_path / "malicious.xml"
    path.write_text(
        '<!DOCTYPE snapshot [<!ENTITY xxe "blocked">]>'
        '<snapshot schema="pokja-pl-master-data" version="1" kode_paket="1">'
        "<cells></cells></snapshot>",
        encoding="utf-8",
    )

    with pytest.raises(SnapshotError, match="DTD/entity"):
        validate_snapshot(path)


def test_snapshot_rejects_empty_or_mismatched_package_code(tmp_path):
    empty = tmp_path / "empty.xml"
    mismatched = tmp_path / "mismatched.xml"
    _write_snapshot(empty, code="")
    _write_snapshot(mismatched, code="22000000000")

    with pytest.raises(SnapshotError, match="kode_paket"):
        validate_snapshot(empty)
    with pytest.raises(SnapshotError, match="tidak cocok"):
        validate_snapshot(mismatched, expected_kode_paket="11000000000")


def test_vba_save_keeps_first_baseline_and_load_rejects_empty_code():
    source_path = Path(__file__).parents[1] / "ModDraftPaketPL.bas"
    source = source_path.read_text(encoding="utf-8")

    assert 'SNAPSHOT_BASELINE_FILE_PL As String = "input_data_baseline.xml"' in source
    assert "If Not fso.FileExists(baselinePath) Then" in source
    assert '"Load dibatalkan: kode paket snapshot/workbook kosong."' in source
    assert "NextSnapshotBackupPathPL(snapshotPath)" in source
    assert "pulihkan current lama" in source
    assert '"Field read-only berubah: " & address' in source
