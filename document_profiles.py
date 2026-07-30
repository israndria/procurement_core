"""Shared document profile helpers for PL Word/PDF generation.

The body template remains independent from the agency header. Profiles are
stored on the document Drive and are injected into a temporary/merged copy.
"""
from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path


def normalize_agency(value: object) -> str:
    """Return canonical agency key: ``pupr`` or ``perdagangan``."""
    text = str(value or "").strip().lower()
    if any(token in text for token in ("perdagangan", "disdag", "dagang")):
        return "perdagangan"
    return "pupr"


def detect_agency(data: dict | None = None, default: str = "pupr") -> str:
    """Detect agency from merged Excel/Supabase values with env override."""
    forced = os.environ.get("POKJA_HEADER_PROFILE", "").strip()
    if forced:
        return normalize_agency(forced)
    data = data or {}
    values = []
    for key in (
        "SKPDOPD", "OPD/SKPD", "OPD", "Nama_Dinas", "nama_dinas",
        "satker", "dinas", "nama_satker", "Satuan_Kerja",
    ):
        if data.get(key):
            values.append(str(data[key]))
    if values:
        return normalize_agency(" ".join(values))
    return normalize_agency(default)


def profile_root(pokja_root: str | os.PathLike[str]) -> Path:
    return Path(pokja_root) / "Paket Experiment - Pengadaan Langsung" / "_Profiles"


def header_profile_path(pokja_root: str | os.PathLike[str], agency: str) -> Path:
    path = profile_root(pokja_root) / f"Header {'Perdagangan' if normalize_agency(agency) == 'perdagangan' else 'PUPR'}.docx"
    if not path.is_file():
        raise FileNotFoundError(f"Header profile tidak ditemukan: {path}")
    return path


def is_official_header_document(path: str | os.PathLike[str]) -> bool:
    """Only official BA/Reviu outputs receive a dynamic agency header."""
    name = Path(path).name.lower()
    return (
        "ba reviu dpp" in name
        or "5. ba pljkk" in name
        or "5. ba plpk" in name
        or "berita acara utama plpk" in name
        or "ba dengan timpang plpk" in name
        or "full dokumen ba plpk" in name
    )


def inject_header_profile(template_path: str | os.PathLike[str], profile_path: str | os.PathLike[str], output_path: str | os.PathLike[str]) -> Path:
    """Copy a DOCX and replace only its header parts with a profile header."""
    template_path = Path(template_path)
    profile_path = Path(profile_path)
    output_path = Path(output_path)
    with zipfile.ZipFile(template_path, "r") as body_zip, zipfile.ZipFile(profile_path, "r") as header_zip:
        body_infos = {item.filename: item for item in body_zip.infolist()}
        files = {name: body_zip.read(name) for name in body_infos}
        for name in ("word/header1.xml", "word/_rels/header1.xml.rels"):
            if name in header_zip.namelist():
                files[name] = header_zip.read(name)
        for name in header_zip.namelist():
            if name.startswith("word/media/") and name not in files:
                files[name] = header_zip.read(name)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for name, data in files.items():
                out_zip.writestr(body_infos.get(name, name), data)
    return output_path


def apply_header_to_copy(template_copy: str | os.PathLike[str], pokja_root: str | os.PathLike[str], data: dict | None = None) -> Path:
    """Replace header in an already-created temporary copy, atomically."""
    target = Path(template_copy)
    if not is_official_header_document(target):
        return target
    profile = header_profile_path(pokja_root, detect_agency(data))
    temp = target.with_suffix(target.suffix + ".header.tmp")
    inject_header_profile(target, profile, temp)
    os.replace(temp, target)
    return target
