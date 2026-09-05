import json
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from pl_snapshot_revision import (
    FAMILY_PLPK,
    FAMILY_PLJKK,
    FORMULA_ADDRESSES,
    IMMUTABLE_ADDRESSES,
    LAYOUT_VERSION_PLPK,
    PLPK_FORMULA_ADDRESSES,
    PLPK_WHITELIST_ADDRESSES,
    CELL_METADATA,
    SnapshotError,
    XML_DATA_SUBFOLDER,
    WHITELIST_ADDRESSES,
    compare_snapshots,
    migrate_legacy_snapshot_files,
    parse_snapshot,
    promote_proposal,
    resolve_snapshot_path,
    seed_proposal,
    snapshot_paths,
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


def _write_profile_snapshot(path: Path, family: str, *, code: str = "11000000000", overrides=None, omit=()):
    overrides = overrides or {}
    whitelist = PLPK_WHITELIST_ADDRESSES if family == FAMILY_PLPK else WHITELIST_ADDRESSES
    formulas = PLPK_FORMULA_ADDRESSES if family == FAMILY_PLPK else FORMULA_ADDRESSES
    layout = LAYOUT_VERSION_PLPK if family == FAMILY_PLPK else "PLJKK-MASTER-DATA-v1"
    root = ET.Element(
        "snapshot",
        schema="pokja-pl-master-data",
        version="2",
        family=family,
        layout_version=layout,
        saved_at="2026-08-21 00:00:00",
        workbook="test.xlsm",
        kode_paket=code,
    )
    cells = ET.SubElement(root, "cells")
    for address in sorted(whitelist):
        if address in omit:
            continue
        default = ("formula", f"={address}") if address in formulas else ("text", f"old-{address}")
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


def test_snapshot_resolver_prefers_canonical_directory(tmp_path):
    legacy = tmp_path / "input_data_snapshot.xml"
    canonical = tmp_path / XML_DATA_SUBFOLDER / "input_data_snapshot.xml"
    canonical.parent.mkdir()
    _write_snapshot(legacy, overrides={"C15": ("text", "legacy")})
    _write_snapshot(canonical, overrides={"C15": ("text", "canonical")})

    resolved = resolve_snapshot_path(legacy)

    assert resolved == canonical
    assert snapshot_paths(tmp_path).data_dir == canonical.parent
    assert parse_snapshot(legacy).cells["C15"].text == "canonical"


def test_seed_routes_write_to_provisioned_canonical_directory(tmp_path):
    legacy_current = tmp_path / "input_data_snapshot.xml"
    proposal_request = tmp_path / "input_data_proposal.xml"
    (tmp_path / XML_DATA_SUBFOLDER).mkdir()
    _write_snapshot(legacy_current)

    result = seed_proposal(legacy_current, proposal_request)

    assert Path(result["proposal"]) == tmp_path / XML_DATA_SUBFOLDER / proposal_request.name
    assert Path(result["proposal"]).is_file()
    assert not proposal_request.exists()


def test_promote_routes_legacy_current_to_canonical_directory(tmp_path):
    legacy_current = tmp_path / "input_data_snapshot.xml"
    proposal_request = tmp_path / "input_data_proposal.xml"
    canonical_dir = tmp_path / XML_DATA_SUBFOLDER
    canonical_dir.mkdir()
    _write_snapshot(legacy_current, overrides={"C15": ("text", "180 Hari")})
    _write_snapshot(proposal_request, overrides={"C15": ("text", "90 Hari")})

    result = promote_proposal(proposal_request, legacy_current)

    canonical_current = canonical_dir / legacy_current.name
    assert result["applied"] is True
    assert canonical_current.is_file()
    assert parse_snapshot(canonical_current).cells["C15"].text == "90 Hari"
    assert legacy_current.read_bytes() != canonical_current.read_bytes()
    assert list(canonical_dir.glob("input_data_snapshot.bak-*.xml"))


def test_legacy_migration_is_opt_in_copy_and_never_overwrites_root(tmp_path):
    legacy = tmp_path / "input_data_snapshot.xml"
    _write_snapshot(legacy)

    preview = migrate_legacy_snapshot_files(tmp_path)
    assert preview["dry_run"] is True
    assert preview["planned"]
    assert not (tmp_path / XML_DATA_SUBFOLDER).exists()

    result = migrate_legacy_snapshot_files(tmp_path, apply=True)

    canonical = tmp_path / XML_DATA_SUBFOLDER / legacy.name
    assert result["migrated"]
    assert canonical.read_bytes() == legacy.read_bytes()
    assert legacy.is_file()


def test_plpk_profile_contains_unique_fields_and_excludes_reserved_rows():
    assert {"C28", "C39", "C45", "C51", "C56", "C63", "C64", "C66", "C75", "C77", "C82", "C89"} <= PLPK_WHITELIST_ADDRESSES
    assert not {f"C{row}" for row in range(57, 63)} & PLPK_WHITELIST_ADDRESSES
    assert {"H15", "H18", "I17", "I18"} <= PLPK_WHITELIST_ADDRESSES
    assert {"H18", "I17", "I18"} <= PLPK_FORMULA_ADDRESSES


def test_profile_collision_is_rejected_and_v1_stays_legacy(tmp_path):
    plpk = tmp_path / "plpk.xml"
    pljkk = tmp_path / "pljkk.xml"
    _write_profile_snapshot(plpk, FAMILY_PLPK)
    _write_profile_snapshot(pljkk, FAMILY_PLJKK)

    assert parse_snapshot(plpk).family == FAMILY_PLPK
    assert parse_snapshot(pljkk).family == FAMILY_PLJKK
    with pytest.raises(SnapshotError, match="Family/layout"):
        compare_snapshots(plpk, pljkk)


def test_plpk_wrong_family_proposal_is_rejected(tmp_path):
    current = tmp_path / "current.xml"
    proposal = tmp_path / "proposal.xml"
    _write_profile_snapshot(current, FAMILY_PLPK)
    _write_profile_snapshot(proposal, FAMILY_PLJKK)
    with pytest.raises(SnapshotError, match="Family/layout"):
        promote_proposal(proposal, current)


def test_vba_save_keeps_first_baseline_and_load_rejects_empty_code():
    source_path = Path(__file__).parents[1] / "ModDraftPaketPL.bas"
    source = source_path.read_text(encoding="utf-8")

    assert 'SNAPSHOT_BASELINE_FILE_PL As String = "input_data_baseline.xml"' in source
    assert 'SNAPSHOT_DATA_FOLDER_PL As String = "11. XML Data"' in source
    assert 'SnapshotFilePathPL(SNAPSHOT_FILE_PL, True)' in source
    assert 'SnapshotFilePathPL(SNAPSHOT_FILE_PL, False)' in source
    assert 'Backward compatibility: snapshot root lama tetap dapat dibaca.' in source
    assert "If Not fso.FileExists(baselinePath) Then" in source
    assert '"Load dibatalkan: kode paket snapshot kosong."' in source
    assert "NextSnapshotBackupPathPL(snapshotPath)" in source
    assert "pulihkan current lama" in source
    assert '"Field read-only berubah: " & address' in source


def test_load_data_pl_keeps_strict_guard_and_explicit_template_migration():
    source = (Path(__file__).parents[1] / "ModDraftPaketPL.bas").read_text(encoding="utf-8")
    assert "Public Sub LoadDataPL(Optional ByVal allowTemplateMigration As Boolean = False)" in source
    assert "snapshotCode <> currentCode And Not allowTemplateMigration" in source
    assert "kode paket workbook kosong. Gunakan mode migrasi template terverifikasi." in source
