"""Protocol aman untuk membandingkan dan menerapkan revisi snapshot PL.

Excel tetap menjadi boundary terakhir. Modul ini hanya membaca/menulis XML
lokal dan tidak membuka workbook atau memakai Excel COM.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


SNAPSHOT_SCHEMA = "pokja-pl-master-data"
SNAPSHOT_VERSION = "2"
LEGACY_SNAPSHOT_VERSION = "1"
FAMILY_PLJKK = "PLJKK"
FAMILY_PLPK = "PLPK"
LAYOUT_VERSION_PLJKK = "PLJKK-MASTER-DATA-v1"
LAYOUT_VERSION_PLPK = "PLPK-MASTER-DATA-v1"
FAMILY_LAYOUTS = {
    FAMILY_PLJKK: LAYOUT_VERSION_PLJKK,
    FAMILY_PLPK: LAYOUT_VERSION_PLPK,
}
MAX_XML_BYTES = 2_000_000
XML_DATA_SUBFOLDER = "11. XML Data"
SNAPSHOT_FILE_NAME = "input_data_snapshot.xml"
PROPOSAL_FILE_NAME = "input_data_proposal.xml"
BASELINE_FILE_NAME = "input_data_baseline.xml"
AUDIT_FILE_NAME = "input_data_audit.jsonl"


@dataclass(frozen=True)
class SnapshotPaths:
    """Canonical local artefact paths for one PL package."""

    package_dir: Path
    data_dir: Path
    snapshot: Path
    proposal: Path
    baseline: Path
    audit: Path

    def ensure_dir(self) -> Path:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        return self.data_dir


def snapshot_paths(package_dir: str | os.PathLike[str], *, create: bool = False) -> SnapshotPaths:
    """Return canonical snapshot paths below ``11. XML Data``.

    ``create=False`` is deliberately side-effect free so read-only audit and
    resolver calls cannot create package state. Writers may request
    ``create=True`` immediately before their atomic write.
    """
    package = Path(package_dir)
    data_dir = package / XML_DATA_SUBFOLDER
    paths = SnapshotPaths(
        package_dir=package,
        data_dir=data_dir,
        snapshot=data_dir / SNAPSHOT_FILE_NAME,
        proposal=data_dir / PROPOSAL_FILE_NAME,
        baseline=data_dir / BASELINE_FILE_NAME,
        audit=data_dir / AUDIT_FILE_NAME,
    )
    if create:
        paths.ensure_dir()
    return paths


def _is_snapshot_artifact_name(name: str) -> bool:
    return name in {
        SNAPSHOT_FILE_NAME,
        PROPOSAL_FILE_NAME,
        BASELINE_FILE_NAME,
        AUDIT_FILE_NAME,
    } or bool(re.fullmatch(r"input_data_snapshot\.bak-[^/\\]+\.xml", name))


def resolve_snapshot_path(
    path: str | os.PathLike[str],
    *,
    for_write: bool = False,
) -> Path:
    """Resolve a snapshot artefact with canonical-first, legacy-safe rules.

    Explicit paths inside ``11. XML Data`` are always honoured. For a legacy
    root path, an existing canonical file wins for reads; writes use the
    canonical directory when it already exists. A package with no provisioned
    directory remains backward-compatible and keeps using its legacy path
    until an explicit migration/provisioning step is performed.
    """
    requested = Path(path)
    if not _is_snapshot_artifact_name(requested.name):
        return requested
    if requested.parent.name.casefold() == XML_DATA_SUBFOLDER.casefold():
        return requested
    candidate = requested.parent / XML_DATA_SUBFOLDER / requested.name
    # New/provisioned packages write canonically. Legacy packages without the
    # folder retain their old path until an explicit migration/provisioning
    # step, preserving backward compatibility for existing tests/workflows.
    if for_write and candidate.parent.is_dir():
        return candidate
    if candidate.is_file():
        return candidate
    return requested


def migrate_legacy_snapshot_files(
    package_dir: str | os.PathLike[str],
    *,
    apply: bool = False,
) -> dict[str, object]:
    """Plan or safely copy root snapshot artefacts into ``11. XML Data``.

    This operation never deletes or overwrites legacy root files. Without
    ``apply`` it is a dry-run; with ``apply=True`` only non-conflicting files
    are copied with metadata preserved. Conflicts are returned for review.
    """
    package = Path(package_dir)
    if not package.is_dir():
        raise SnapshotError(f"Folder paket tidak ditemukan: {package}")
    paths = snapshot_paths(package)
    planned: list[dict[str, str]] = []
    conflicts: list[dict[str, str]] = []
    for source in sorted(package.iterdir(), key=lambda item: item.name.casefold()):
        if not source.is_file() or not _is_snapshot_artifact_name(source.name):
            continue
        target = paths.data_dir / source.name
        row = {"source": str(source), "target": str(target)}
        if target.exists():
            conflicts.append(row)
        else:
            planned.append(row)

    migrated: list[dict[str, str]] = []
    if apply and planned:
        paths.ensure_dir()
        for row in planned:
            source = Path(row["source"])
            target = Path(row["target"])
            shutil.copy2(source, target)
            migrated.append(row)
    return {
        "ok": not conflicts,
        "package_dir": str(package),
        "data_dir": str(paths.data_dir),
        "dry_run": not apply,
        "planned": planned,
        "migrated": migrated,
        "conflicts": conflicts,
    }


def _addresses(*ranges: tuple[str, int, int]) -> frozenset[str]:
    return frozenset(
        f"{column}{row}"
        for column, start, end in ranges
        for row in range(start, end + 1)
    )


PLJKK_WHITELIST_ADDRESSES = _addresses(
    ("C", 3, 26), ("C", 29, 30), ("C", 32, 43),
    ("C", 51, 54), ("C", 56, 63), ("F", 2, 2),
    ("H", 8, 11), ("I", 8, 10),
)
PLPK_WHITELIST_ADDRESSES = _addresses(
    ("C", 3, 28), ("C", 30, 31), ("C", 33, 38),
    ("C", 39, 56), ("C", 63, 64), ("C", 66, 75),
    ("C", 77, 80), ("C", 82, 89), ("F", 2, 2),
    ("H", 8, 11), ("I", 8, 10), ("H", 15, 18), ("I", 17, 18),
)

# Kompatibilitas API lama = profile PLJKK/legacy.
WHITELIST_ADDRESSES = PLJKK_WHITELIST_ADDRESSES

PLJKK_FORMULA_ADDRESSES = frozenset(
    {
        "C11",
        "C12",
        "C20",
        "C22",
        "C24",
        "C26",
        "H10",
        "H11",
        "I8",
        "I9",
        "I10",
    }
)
PLPK_FORMULA_ADDRESSES = PLJKK_FORMULA_ADDRESSES | frozenset({"C27", "H18", "I17", "I18"})
FORMULA_ADDRESSES = PLJKK_FORMULA_ADDRESSES
IMMUTABLE_ADDRESSES = frozenset({"C3", "F2"})
PLJKK_READ_ONLY_ADDRESSES = PLJKK_FORMULA_ADDRESSES | IMMUTABLE_ADDRESSES
PLPK_READ_ONLY_ADDRESSES = PLPK_FORMULA_ADDRESSES | IMMUTABLE_ADDRESSES
READ_ONLY_ADDRESSES = PLJKK_READ_ONLY_ADDRESSES


def _metadata(key: str, label: str, section: str, editable: bool = True) -> dict[str, object]:
    return {
        "key": key,
        "label": label,
        "section": section,
        "editable": editable,
    }


CELL_METADATA: dict[str, dict[str, object]] = {
    "C3": _metadata("kode_paket", "Kode Paket", "Identitas", False),
    "C4": _metadata("kode_rup", "Kode RUP", "Identitas"),
    "C5": _metadata("nama_pekerjaan", "Nama Pekerjaan", "Identitas"),
    "C6": _metadata("nama_skpd", "Nama SKPD", "Identitas"),
    "C7": _metadata("sub_kegiatan", "Sub Kegiatan", "Identitas"),
    "C8": _metadata("nama_ppk", "Nama PPK", "PPK"),
    "C9": _metadata("nip_ppk", "NIP PPK", "PPK"),
    "C10": _metadata("no_sk_ppk", "Nomor SK PPK", "PPK"),
    "C11": _metadata("nama_pp", "Nama Pejabat Pengadaan", "PP"),
    "C12": _metadata("nip_pp", "NIP Pejabat Pengadaan", "PP"),
    "C13": _metadata("pagu", "Pagu", "Anggaran"),
    "C14": _metadata("hps", "HPS", "Anggaran"),
    "C15": _metadata("jangka_waktu", "Jangka Waktu", "Kontrak"),
    "C16": _metadata("sumber_dana", "Sumber Dana", "Anggaran"),
    "C17": _metadata("lokasi", "Lokasi", "Pekerjaan"),
    "C18": _metadata("jenis_kontrak", "Jenis Kontrak", "Kontrak"),
    "C19": _metadata("uraian_singkat", "Uraian Singkat", "Pekerjaan"),
    "C20": _metadata("nomor_dokpil", "Nomor Dokpil", "Dokpil"),
    "C21": _metadata("tanggal_dokpil", "Tanggal Dokpil", "Dokpil"),
    "C22": _metadata("nomor_undangan", "Nomor Undangan", "Administrasi"),
    "C23": _metadata("tahun_anggaran", "Tahun Anggaran", "Anggaran"),
    "C24": _metadata("tanggal_undangan", "Tanggal Undangan", "Administrasi"),
    "C25": _metadata("kode_rekening", "Kode Rekening", "Anggaran"),
    "C26": _metadata("nomor_ba_reviu", "Nomor BA Reviu", "Administrasi"),
    "C29": _metadata("sbu_baru", "SBU Baru", "Kualifikasi"),
    "C30": _metadata("sbu_lama", "SBU Lama", "Kualifikasi"),
    "C32": _metadata("personil_1_jabatan", "Personil 1 - Jabatan", "Personil/K3"),
    "C33": _metadata("personil_1_nama", "Personil 1 - Nama", "Personil/K3"),
    "C34": _metadata("personil_1_pengalaman", "Personil 1 - Pengalaman", "Personil/K3"),
    "C35": _metadata("personil_1_sertifikat", "Personil 1 - Sertifikat", "Personil/K3"),
    "C36": _metadata("personil_2_jabatan", "Personil 2 - Jabatan", "Personil/K3"),
    "C37": _metadata("personil_2_nama", "Personil 2 - Nama", "Personil/K3"),
    "C38": _metadata("personil_2_pengalaman", "Personil 2 - Pengalaman", "Personil/K3"),
    "C39": _metadata("personil_2_sertifikat", "Personil 2 - Sertifikat", "Personil/K3"),
    "C40": _metadata("personil_3_jabatan", "Personil 3 - Jabatan", "Personil/K3"),
    "C41": _metadata("personil_3_nama", "Personil 3 - Nama", "Personil/K3"),
    "C42": _metadata("personil_3_pengalaman", "Personil 3 - Pengalaman", "Personil/K3"),
    "C43": _metadata("personil_3_sertifikat", "Personil 3 - Sertifikat", "Personil/K3"),
    "C51": _metadata("nama_peserta", "Nama Peserta", "Peserta"),
    "C52": _metadata("npwp_peserta", "NPWP Peserta", "Peserta"),
    "C53": _metadata("peserta_tambahan_1", "Data Peserta Tambahan 1", "Peserta"),
    "C54": _metadata("peserta_tambahan_2", "Data Peserta Tambahan 2", "Peserta"),
    "C56": _metadata("metadata_56", "Metadata 56", "Metadata"),
    "C57": _metadata("metadata_57", "Metadata 57", "Metadata"),
    "C58": _metadata("telepon_pp", "Telepon PP", "Metadata"),
    "C59": _metadata("alamat_pp", "Alamat PP", "Metadata"),
    "C60": _metadata("masa_berlaku", "Masa Berlaku", "Metadata"),
    "C61": _metadata("nomor_nota_dinas", "Nomor Nota Dinas", "Metadata"),
    "C62": _metadata("tanggal_nota_dinas", "Tanggal Nota Dinas", "Metadata"),
    "C63": _metadata("nomor_rekomendasi", "Nomor Rekomendasi", "Metadata"),
    "F2": _metadata("kode_unik", "Kode Unik", "Administrasi", False),
}
for _address in FORMULA_ADDRESSES:
    CELL_METADATA.setdefault(_address, _metadata(_address.lower(), _address, "Turunan", False))[
        "editable"
    ] = False
for _address in ("H8", "H9"):
    CELL_METADATA[_address] = _metadata(_address.lower(), _address, "Tanggal Sumber")

# Metadata lama dipertahankan untuk caller existing, tetapi semantic labels
# profile-specific mencegah collision C51:C63 antara PLJKK dan PLPK.
PLJKK_CELL_METADATA = dict(CELL_METADATA)
PLJKK_CELL_METADATA.update(
    {
        "C33": _metadata("personil_1_pengalaman", "Personil 1 - Pengalaman", "Personil/K3"),
        "C34": _metadata("personil_1_sertifikat", "Personil 1 - Sertifikat", "Personil/K3"),
        "C36": _metadata("personil_2_pengalaman", "Personil 2 - Pengalaman", "Personil/K3"),
        "C37": _metadata("personil_2_sertifikat", "Personil 2 - Sertifikat", "Personil/K3"),
        "C39": _metadata("personil_3_pengalaman", "Personil 3 - Pengalaman", "Personil/K3"),
        "C40": _metadata("personil_3_sertifikat", "Personil 3 - Sertifikat", "Personil/K3"),
        "C53": _metadata("nilai_penawaran", "Nilai Penawaran", "Peserta"),
        "C54": _metadata("nilai_negosiasi", "Nilai Negosiasi", "Peserta"),
        "C56": _metadata("nama_pp", "Nama Pejabat Pengadaan", "Metadata"),
        "C57": _metadata("nip_pp", "NIP Pejabat Pengadaan", "Metadata"),
    }
)
PLPK_CELL_METADATA = dict(CELL_METADATA)
PLPK_CELL_METADATA.update(
    {
        "C28": _metadata("cara_pembayaran", "Cara Pembayaran", "Kontrak"),
        "C30": _metadata("sbu_baru", "SBU Baru", "Kualifikasi"),
        "C31": _metadata("sbu_lama", "SBU Lama", "Kualifikasi"),
        "C33": _metadata("personil_1_jabatan", "Personil 1 - Jabatan", "Personil/K3"),
        "C34": _metadata("personil_1_pengalaman", "Personil 1 - Pengalaman", "Personil/K3"),
        "C35": _metadata("personil_1_sertifikat", "Personil 1 - Sertifikat", "Personil/K3"),
        "C36": _metadata("personil_2_jabatan", "Personil 2 - Jabatan", "Personil/K3"),
        "C37": _metadata("personil_2_pengalaman", "Personil 2 - Pengalaman", "Personil/K3"),
        "C38": _metadata("personil_2_sertifikat", "Personil 2 - Sertifikat", "Personil/K3"),
        "C63": _metadata("uraian_risiko_k3", "Uraian Pekerjaan Risiko K3", "K3"),
        "C64": _metadata("risiko_tertinggi_k3", "Risiko Tertinggi/Fatal K3", "K3"),
        "C77": _metadata("nama_peserta", "Nama Peserta", "Peserta"),
        "C78": _metadata("npwp_peserta", "NPWP Peserta", "Peserta"),
        "C79": _metadata("nilai_penawaran", "Nilai Penawaran", "Peserta"),
        "C80": _metadata("nilai_negosiasi", "Nilai Negosiasi", "Peserta"),
        "C82": _metadata("nama_pp", "Nama Pejabat Pengadaan", "Metadata"),
        "C83": _metadata("nip_pp", "NIP Pejabat Pengadaan", "Metadata"),
        "C84": _metadata("telepon_pp", "Telepon PP", "Metadata"),
        "C85": _metadata("alamat_pp", "Alamat PP", "Metadata"),
        "C86": _metadata("masa_berlaku", "Masa Berlaku", "Metadata"),
        "C87": _metadata("nomor_nota_dinas", "Nomor Nota Dinas", "Metadata"),
        "C88": _metadata("tanggal_nota_dinas", "Tanggal Nota Dinas", "Metadata"),
        "C89": _metadata("nomor_rekomendasi", "Nomor Rekomendasi", "Metadata"),
    }
)
for _row in range(39, 45):
    PLPK_CELL_METADATA[f"C{_row}"] = _metadata(f"alat_{_row - 38}", f"Alat {_row - 38}", "Peralatan")
for _row in range(45, 51):
    PLPK_CELL_METADATA[f"C{_row}"] = _metadata(f"kapasitas_{_row - 44}", f"Kapasitas {_row - 44}", "Peralatan")
for _row in range(51, 57):
    PLPK_CELL_METADATA[f"C{_row}"] = _metadata(f"jumlah_alat_{_row - 50}", f"Jumlah Alat {_row - 50}", "Peralatan")
for _row in range(66, 76):
    PLPK_CELL_METADATA[f"C{_row}"] = _metadata(
        f"uraian_pekerjaan_{_row - 65}", f"Uraian Pekerjaan {_row - 65}", "Pekerjaan"
    )

CELL_METADATA = PLJKK_CELL_METADATA


def _profile_config(family: str) -> tuple[frozenset[str], frozenset[str], dict[str, dict[str, object]]]:
    family = family.strip().upper()
    if family == FAMILY_PLJKK:
        return PLJKK_WHITELIST_ADDRESSES, PLJKK_FORMULA_ADDRESSES, PLJKK_CELL_METADATA
    if family == FAMILY_PLPK:
        return PLPK_WHITELIST_ADDRESSES, PLPK_FORMULA_ADDRESSES, PLPK_CELL_METADATA
    raise SnapshotError(f"Family snapshot tidak didukung: {family!r}")


def _read_only_addresses(family: str, cells: dict[str, "SnapshotCell"] | None = None) -> frozenset[str]:
    _, formulas, _ = _profile_config(family)
    result = set(formulas) | set(IMMUTABLE_ADDRESSES)
    # C21 boleh manual pada PLPK; jika workbook menyimpannya sebagai formula,
    # formula tersebut tetap read-only dan sumber tanggal tidak boleh ditimpa AI.
    if family.strip().upper() == FAMILY_PLPK and cells and cells.get("C21") is not None:
        if cells["C21"].cell_type == "formula":
            result.add("C21")
    return frozenset(result)


def _profile_from_root(version: str, root: ET.Element) -> tuple[str, str]:
    if version == LEGACY_SNAPSHOT_VERSION:
        return FAMILY_PLJKK, "legacy-v1"
    if version != SNAPSHOT_VERSION:
        raise SnapshotError(f"Versi snapshot tidak didukung: {version!r}")
    family = (root.get("family") or "").strip().upper()
    layout_version = (root.get("layout_version") or "").strip()
    if family not in FAMILY_LAYOUTS:
        raise SnapshotError(f"Family snapshot tidak didukung: {family!r}")
    if layout_version != FAMILY_LAYOUTS[family]:
        raise SnapshotError(f"Layout snapshot tidak cocok untuk {family}: {layout_version!r}")
    return family, layout_version


class SnapshotError(ValueError):
    """Snapshot tidak aman untuk dibaca atau dipromosikan."""


@dataclass(frozen=True)
class SnapshotCell:
    address: str
    cell_type: str
    text: str


@dataclass(frozen=True)
class Snapshot:
    path: Path
    schema: str
    version: str
    kode_paket: str
    family: str
    layout_version: str
    attributes: dict[str, str]
    cells: dict[str, SnapshotCell]

    @property
    def sha256(self) -> str:
        return sha256_file(self.path)


def sha256_file(path: str | os.PathLike[str]) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalize_address(value: str) -> str:
    address = value.strip().upper().replace("$", "")
    if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", address):
        raise SnapshotError(f"Alamat cell tidak valid: {value!r}")
    return address


def _cell_text(node: ET.Element) -> str:
    return "".join(node.itertext())


def parse_snapshot(path: str | os.PathLike[str]) -> Snapshot:
    source = resolve_snapshot_path(path)
    if not source.is_file():
        raise SnapshotError(f"File snapshot tidak ditemukan: {source}")
    if source.stat().st_size > MAX_XML_BYTES:
        raise SnapshotError(f"Snapshot terlalu besar: {source.stat().st_size} byte")
    raw = source.read_bytes()
    if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
        raise SnapshotError("DTD/entity XML tidak diizinkan")
    try:
        root = ET.fromstring(raw)
    except (ET.ParseError, OSError) as exc:
        raise SnapshotError(f"XML snapshot tidak valid: {source}: {exc}") from exc
    if root.tag != "snapshot":
        raise SnapshotError("Root XML harus <snapshot>")
    schema = (root.get("schema") or "").strip()
    version = (root.get("version") or "").strip()
    if schema != SNAPSHOT_SCHEMA:
        raise SnapshotError(f"Schema tidak cocok: {schema!r}")
    family, layout_version = _profile_from_root(version, root)
    kode_paket = (root.get("kode_paket") or "").strip()
    if not kode_paket:
        raise SnapshotError("kode_paket wajib ada dan tidak boleh kosong")

    cells: dict[str, SnapshotCell] = {}
    cell_parent = root.find("cells")
    if cell_parent is None:
        raise SnapshotError("Elemen <cells> tidak ditemukan")
    for node in cell_parent.findall("cell"):
        address = _normalize_address(node.get("address") or "")
        whitelist, _, _ = _profile_config(family)
        if address not in whitelist:
            raise SnapshotError(f"Alamat cell di luar whitelist {family}: {address}")
        if address in cells:
            raise SnapshotError(f"Alamat cell duplikat: {address}")
        cell_type = (node.get("type") or "text").strip().lower()
        if cell_type not in {"empty", "formula", "number", "boolean", "error", "text"}:
            raise SnapshotError(f"Tipe cell tidak didukung ({address}): {cell_type!r}")
        cells[address] = SnapshotCell(address, cell_type, _cell_text(node))

    return Snapshot(
        path=source,
        schema=schema,
        version=version,
        kode_paket=kode_paket,
        family=family,
        layout_version=layout_version,
        attributes={key: value for key, value in root.attrib.items()},
        cells=cells,
    )


def validate_snapshot(
    path: str | os.PathLike[str],
    *,
    expected_kode_paket: str | None = None,
    expected_family: str | None = None,
    require_complete: bool = False,
) -> dict[str, object]:
    snapshot = parse_snapshot(path)
    if expected_kode_paket is not None and snapshot.kode_paket != expected_kode_paket.strip():
        raise SnapshotError(
            f"Kode paket tidak cocok: snapshot={snapshot.kode_paket}, "
            f"expected={expected_kode_paket}"
        )
    if expected_family and snapshot.family != expected_family.strip().upper():
        raise SnapshotError(
            f"Family snapshot tidak cocok: snapshot={snapshot.family}, expected={expected_family}"
        )
    whitelist, formulas, _ = _profile_config(snapshot.family)
    missing = sorted(whitelist - snapshot.cells.keys())
    if require_complete and missing:
        raise SnapshotError(
            "Snapshot proposal tidak lengkap; gunakan type=empty untuk clear: "
            + ", ".join(missing)
        )
    if require_complete:
        wrong_formula = sorted(
            address
            for address in formulas
            if snapshot.cells[address].cell_type != "formula"
        )
        if wrong_formula:
            raise SnapshotError(
                "Formula wajib tidak bertipe formula: " + ", ".join(wrong_formula)
            )
    return {
        "ok": True,
        "path": str(snapshot.path),
        "kode_paket": snapshot.kode_paket,
        "family": snapshot.family,
        "layout_version": snapshot.layout_version,
        "cell_count": len(snapshot.cells),
        "missing_count": len(missing),
        "sha256": snapshot.sha256,
    }


def _canonical(cell: SnapshotCell | None) -> tuple[str, str] | None:
    if cell is None:
        return None
    text = " ".join(cell.text.split())
    if cell.cell_type == "number":
        try:
            return cell.cell_type, format(Decimal(text.replace(",", ".")), "f")
        except InvalidOperation:
            pass
    return cell.cell_type, text


def _metadata_for(address: str, family: str = FAMILY_PLJKK) -> dict[str, object]:
    _, _, metadata = _profile_config(family)
    return metadata.get(
        address,
        _metadata(address.lower(), address, "Unknown", False),
    )


def compare_snapshots(
    baseline_path: str | os.PathLike[str],
    candidate_path: str | os.PathLike[str],
) -> list[dict[str, object]]:
    baseline = parse_snapshot(baseline_path)
    candidate = parse_snapshot(candidate_path)
    if baseline.kode_paket != candidate.kode_paket:
        raise SnapshotError(
            f"Kode paket berbeda: baseline={baseline.kode_paket}, "
            f"candidate={candidate.kode_paket}"
        )
    if baseline.family != candidate.family or baseline.layout_version != candidate.layout_version:
        raise SnapshotError(
            f"Family/layout berbeda: baseline={baseline.family}/{baseline.layout_version}, "
            f"candidate={candidate.family}/{candidate.layout_version}"
        )
    changes: list[dict[str, object]] = []
    for address in sorted(set(baseline.cells) | set(candidate.cells)):
        old = baseline.cells.get(address)
        new = candidate.cells.get(address)
        if _canonical(old) == _canonical(new):
            continue
        meta = _metadata_for(address, baseline.family)
        changes.append(
            {
                "address": address,
                "key": meta["key"],
                "label": meta["label"],
                "section": meta["section"],
                "editable": meta["editable"],
                "status": "changed" if old and new else ("removed" if old else "added"),
                "old_type": old.cell_type if old else None,
                "old_value": old.text if old else None,
                "new_type": new.cell_type if new else None,
                "new_value": new.text if new else None,
            }
        )
    return changes


def render_compare_markdown(
    baseline_path: str | os.PathLike[str],
    candidate_path: str | os.PathLike[str],
    changes: Iterable[dict[str, object]] | None = None,
) -> str:
    baseline = parse_snapshot(baseline_path)
    candidate = parse_snapshot(candidate_path)
    rows = list(changes if changes is not None else compare_snapshots(baseline_path, candidate_path))
    lines = [
        "# Perbandingan Revisi Snapshot PL",
        "",
        f"- Kode paket: `{baseline.kode_paket}`",
        f"- Family/layout: `{baseline.family}` / `{baseline.layout_version}`",
        f"- Baseline: `{baseline.path.name}` (`{baseline.sha256[:12]}`)",
        f"- Kandidat: `{candidate.path.name}` (`{candidate.sha256[:12]}`)",
        f"- Perubahan: **{len(rows)}**",
        "",
    ]
    if not rows:
        lines.append("Tidak ada perbedaan nilai.")
        return "\n".join(lines) + "\n"
    lines.extend(
        [
            "| Cell | Field | Nilai awal | Nilai revisi | Status |",
            "|---|---|---|---|---|",
        ]
    )
    for row in rows:
        old = str(row["old_value"] or "").replace("|", "\\|").replace("\n", " ")
        new = str(row["new_value"] or "").replace("|", "\\|").replace("\n", " ")
        lines.append(
            f"| `{row['address']}` | {row['label']} | {old} | {new} | {row['status']} |"
        )
    lines.extend(
        [
            "",
            "> File kandidat belum dianggap diterapkan. Promosikan hanya setelah sumber revisi diverifikasi.",
        ]
    )
    return "\n".join(lines) + "\n"


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _backup_path(target: Path) -> Path:
    base = target.with_name(f"{target.stem}.bak-{_utc_stamp()}{target.suffix}")
    if not base.exists():
        return base
    index = 2
    while True:
        candidate = target.with_name(
            f"{target.stem}.bak-{_utc_stamp()}-{index}{target.suffix}"
        )
        if not candidate.exists():
            return candidate
        index += 1


def _atomic_copy(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=target.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        shutil.copyfile(source, temp_path)
        with temp_path.open("r+b") as handle:
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, target)
    finally:
        temp_path.unlink(missing_ok=True)


def _atomic_write_bytes(data: bytes, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=target.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_bytes(data)
        with temp_path.open("r+b") as handle:
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, target)
    finally:
        temp_path.unlink(missing_ok=True)


def _same_path(first: Path, second: Path) -> bool:
    return first.resolve() == second.resolve()


def promote_proposal(
    proposal_path: str | os.PathLike[str],
    current_path: str | os.PathLike[str],
    *,
    baseline_path: str | os.PathLike[str] | None = None,
    audit_path: str | os.PathLike[str] | None = None,
    expected_kode_paket: str | None = None,
) -> dict[str, object]:
    proposal = parse_snapshot(resolve_snapshot_path(proposal_path))
    current = parse_snapshot(resolve_snapshot_path(current_path))
    if _same_path(proposal.path, current.path):
        raise SnapshotError("Proposal dan current tidak boleh file yang sama")
    if baseline_path is not None:
        baseline_target = resolve_snapshot_path(baseline_path)
        if _same_path(baseline_target, current.path) or _same_path(baseline_target, proposal.path):
            raise SnapshotError("Baseline immutable tidak boleh menjadi target promosi")
    if proposal.kode_paket != current.kode_paket:
        raise SnapshotError("Kode paket proposal berbeda dari current")
    if proposal.family != current.family or proposal.layout_version != current.layout_version:
        raise SnapshotError("Family/layout proposal berbeda dari current")
    if expected_kode_paket and proposal.kode_paket != expected_kode_paket.strip():
        raise SnapshotError("Kode paket proposal tidak sesuai expected")
    if baseline_path is not None:
        baseline = parse_snapshot(resolve_snapshot_path(baseline_path))
        if baseline.kode_paket != current.kode_paket:
            raise SnapshotError("Kode paket baseline berbeda dari current")
        if baseline.family != current.family or baseline.layout_version != current.layout_version:
            raise SnapshotError("Family/layout baseline berbeda dari current")
        validate_snapshot(
            baseline.path,
            expected_kode_paket=current.kode_paket,
            expected_family=current.family,
            require_complete=True,
        )
    validate_snapshot(proposal.path, expected_family=current.family, require_complete=True)
    validate_snapshot(
        current.path,
        expected_kode_paket=current.kode_paket,
        expected_family=current.family,
        require_complete=True,
    )
    whitelist, formulas, _ = _profile_config(current.family)
    missing_current = whitelist - current.cells.keys()
    if missing_current:
        raise SnapshotError("Current snapshot tidak lengkap: " + ", ".join(sorted(missing_current)))

    source_sha256 = proposal.attributes.get("source_sha256", "").strip()
    if source_sha256 and source_sha256 != current.sha256:
        raise SnapshotError(
            "Proposal dibuat dari current lama; buat proposal baru dari snapshot terbaru"
        )

    read_only = _read_only_addresses(current.family, current.cells)
    for address in read_only:
        if _canonical(current.cells.get(address)) != _canonical(proposal.cells.get(address)):
            meta = _metadata_for(address, current.family)
            raise SnapshotError(f"Field read-only berubah: {address} ({meta['label']})")
    formula_added = sorted(
        address
        for address, cell in proposal.cells.items()
        if cell.cell_type == "formula"
        and address not in formulas
        and not (address == "C21" and current.cells.get("C21", None) is not None
                 and current.cells["C21"].cell_type == "formula")
    )
    if formula_added:
        raise SnapshotError(
            "Proposal tidak boleh menambah formula pada field input: "
            + ", ".join(formula_added)
        )

    changes = compare_snapshots(current.path, proposal.path)
    if not changes:
        return {"ok": True, "changed": 0, "applied": False, "message": "Tidak ada perubahan."}

    # If caller used legacy root notation for a provisioned package, promote
    # into the canonical directory instead of recreating a root artefact.
    target = resolve_snapshot_path(current_path, for_write=True)
    backup = _backup_path(target)
    before_sha = current.sha256
    target_existed = target.is_file()
    # Jika current masih legacy di root sementara folder canonical sudah
    # diprovisioning, target belum ada. Backup harus mengambil sumber current
    # yang benar, bukan memaksa membaca target canonical yang kosong.
    shutil.copy2(target if target_existed else current.path, backup)
    after_sha = sha256_file(proposal.path)
    try:
        _atomic_copy(proposal.path, target)
        written = parse_snapshot(target)
        if (
            written.sha256 != after_sha
            or written.kode_paket != current.kode_paket
            or written.family != current.family
        ):
            raise SnapshotError("Verifikasi current sesudah promosi gagal")
        if audit_path is not None:
            audit = resolve_snapshot_path(audit_path, for_write=True)
            audit.parent.mkdir(parents=True, exist_ok=True)
            record = {
                "timestamp_utc": datetime.now(timezone.utc).isoformat(),
                "actor": "ai-proposal",
                "kode_paket": current.kode_paket,
                "baseline": str(baseline_path) if baseline_path else None,
                "proposal": str(proposal.path),
                "target": str(target),
                "before_sha256": before_sha,
                "after_sha256": after_sha,
                "source_sha256": source_sha256 or None,
                "changed_cells": [row["address"] for row in changes],
                "status": "applied",
            }
            with audit.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        if target_existed:
            _atomic_copy(backup, target)
        else:
            target.unlink(missing_ok=True)
        raise
    return {
        "ok": True,
        "changed": len(changes),
        "applied": True,
        "backup": str(backup),
        "before_sha256": before_sha,
        "after_sha256": after_sha,
        "changed_cells": [row["address"] for row in changes],
    }


def seed_proposal(current_path: str | os.PathLike[str], proposal_path: str | os.PathLike[str]) -> dict[str, object]:
    current = parse_snapshot(resolve_snapshot_path(current_path))
    proposal = resolve_snapshot_path(proposal_path, for_write=True)
    if proposal.resolve() == current.path.resolve():
        raise SnapshotError("Proposal dan current tidak boleh file yang sama")
    validate_snapshot(current.path, require_complete=True)
    root = ET.fromstring(current.path.read_bytes())
    root.set("source_sha256", current.sha256)
    root.set("proposal_created_at", datetime.now(timezone.utc).isoformat())
    _atomic_write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True), proposal)
    return {
        "ok": True,
        "proposal": str(proposal),
        "kode_paket": current.kode_paket,
        "source_sha256": current.sha256,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Validasi dan bandingkan revisi XML snapshot PL")
    sub = parser.add_subparsers(dest="command", required=True)

    validate = sub.add_parser("validate")
    validate.add_argument("snapshot")
    validate.add_argument("--expected-kode-paket")
    validate.add_argument("--expected-family")
    validate.add_argument("--complete", action="store_true")

    compare = sub.add_parser("compare")
    compare.add_argument("baseline")
    compare.add_argument("candidate")
    compare.add_argument("--output", type=Path)

    seed = sub.add_parser("seed-proposal")
    seed.add_argument("current")
    seed.add_argument("proposal")

    migrate = sub.add_parser(
        "migrate-root",
        help="copy artefak snapshot legacy dari root ke 11. XML Data tanpa menghapus root",
    )
    migrate.add_argument("package_dir")
    migrate.add_argument("--apply", action="store_true")

    promote = sub.add_parser("promote")
    promote.add_argument("proposal")
    promote.add_argument("current")
    promote.add_argument("--baseline")
    promote.add_argument("--audit")
    promote.add_argument("--expected-kode-paket")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "validate":
            result = validate_snapshot(
                args.snapshot,
                expected_kode_paket=args.expected_kode_paket,
                expected_family=args.expected_family,
                require_complete=args.complete,
            )
        elif args.command == "compare":
            result = {"changes": compare_snapshots(args.baseline, args.candidate)}
            report = render_compare_markdown(args.baseline, args.candidate, result["changes"])
            if args.output:
                args.output.write_text(report, encoding="utf-8")
                result["report"] = str(args.output)
            else:
                print(report, end="")
        elif args.command == "seed-proposal":
            result = seed_proposal(args.current, args.proposal)
        elif args.command == "migrate-root":
            result = migrate_legacy_snapshot_files(args.package_dir, apply=args.apply)
        else:
            result = promote_proposal(
                args.proposal,
                args.current,
                baseline_path=args.baseline,
                audit_path=args.audit,
                expected_kode_paket=args.expected_kode_paket,
            )
        if args.command != "compare" or args.output:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except SnapshotError as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
