"""Shared document profile helpers for PL Word/PDF generation.

The body template remains independent from the agency header. Profiles are
stored on the document Drive and are injected into a temporary/merged copy.
"""
from __future__ import annotations

import os
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
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
        "SKPDOPD", "SKPD", "OPD/SKPD", "OPD", "Nama_Dinas", "nama_dinas",
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
    label = "Perdagangan" if normalize_agency(agency) == "perdagangan" else "PUPR"
    # Prototype header dinamis adalah sumber desain resmi. _Profiles tetap
    # fallback portable agar paket lama/PC lain tidak langsung rusak bila
    # folder prototype belum tersinkron.
    candidates = (
        Path(pokja_root) / "Paket Experiment - Pengadaan Langsung"
        / "V2 - Prototype Header Dinamis" / f"Header {label} - V2.docx",
        profile_root(pokja_root) / f"Header {label}.docx",
    )
    for path in candidates:
        if path.is_file():
            return path
    raise FileNotFoundError(
        "Header profile tidak ditemukan. Dicari pada: "
        + "; ".join(str(path) for path in candidates)
    )


def is_official_header_document(path: str | os.PathLike[str]) -> bool:
    """Only official BA/Reviu outputs receive a dynamic agency header."""
    name = Path(path).name.lower()
    return (
        "ba reviu dpp" in name
        or "ba reviu pl" in name
        or "5. ba pljkk" in name
        or "5. ba plpk" in name
        or "berita acara utama plpk" in name
        or "ba dengan timpang plpk" in name
        or "full dokumen ba plpk" in name
    )


def strip_static_headers(
    template_copy: str | os.PathLike[str],
    output_path: str | os.PathLike[str] | None = None,
) -> Path:
    """Kosongkan header statis donor pada salinan dokumen paket.

    Donor V2 masih membawa header PUPR sebagai placeholder historis. Header
    resmi dipasang saat export melalui :func:`apply_header_to_copy`, sehingga
    salinan DOCX paket harus netral juga ketika dibuka langsung oleh operator.
    Struktur ZIP, section, relationship, field, dan media tetap dipertahankan.
    """
    target = Path(template_copy)
    result = Path(output_path) if output_path else target
    changed = False
    with zipfile.ZipFile(target, "r") as source_zip:
        items = source_zip.infolist()
        files = {item.filename: source_zip.read(item.filename) for item in items}
        for name in tuple(files):
            if not (name.startswith("word/header") and name.endswith(".xml")):
                continue
            root = ET.fromstring(files[name])
            for child in list(root):
                root.remove(child)
            files[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changed = True
    if not changed:
        return result
    if result != target:
        result.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(result, "w", zipfile.ZIP_DEFLATED) as output_zip:
            for item in items:
                output_zip.writestr(item, files[item.filename])
        return result
    temp_fd, temp_name = tempfile.mkstemp(prefix="h_", suffix=".tmp", dir=target.parent)
    os.close(temp_fd)
    temp = Path(temp_name)
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as output_zip:
            for item in items:
                output_zip.writestr(item, files[item.filename])
        os.replace(temp, target)
    finally:
        if temp.exists():
            temp.unlink()
    return target


def _rewrite_header_relationships(rels: bytes, media_map: dict[str, str]) -> bytes:
    """Rewrite relative image targets after profile media files are renamed."""
    if not any(source != target for source, target in media_map.items()):
        return rels

    root = ET.fromstring(rels)
    for relationship in root:
        target = relationship.get("Target", "")
        if not target or target.startswith(("/", "#")):
            target_path = target.lstrip("/")
        else:
            target_path = posixpath.normpath(posixpath.join("word", target))
        replacement = media_map.get(target_path)
        if replacement:
            relationship.set("Target", posixpath.relpath(replacement, "word"))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _next_relationship_id(root: ET.Element) -> str:
    """Return a document relationship id not used by the target document."""
    used = {node.get("Id", "") for node in root}
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def _ensure_header_reference(document_xml: bytes, relationship_id: str) -> bytes:
    """Attach the supplied default header to every section in the document.

    The V2 PL templates are intentionally headerless. Replacing a header part
    alone therefore has no visible effect: Word has no ``headerReference`` to
    render. This helper adds/updates only the default reference and preserves
    all other section properties.
    """
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("w", ns_w)
    ET.register_namespace("r", ns_r)
    root = ET.fromstring(document_xml)
    sect_prs = root.findall(f".//{{{ns_w}}}sectPr")
    for sect_pr in sect_prs:
        reference = None
        for candidate in sect_pr.findall(f"{{{ns_w}}}headerReference"):
            if candidate.get(f"{{{ns_w}}}type") == "default":
                reference = candidate
                break
        if reference is None:
            reference = ET.Element(f"{{{ns_w}}}headerReference")
            reference.set(f"{{{ns_w}}}type", "default")
            # headerReference belongs before footerReference/pgSz in sectPr.
            insert_at = len(sect_pr)
            for i, child in enumerate(list(sect_pr)):
                if child.tag in {
                    f"{{{ns_w}}}footerReference",
                    f"{{{ns_w}}}pgSz",
                    f"{{{ns_w}}}pgMar",
                }:
                    insert_at = i
                    break
            sect_pr.insert(insert_at, reference)
        reference.set(f"{{{ns_r}}}id", relationship_id)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_content_type(files: dict[str, bytes], extension: str, content_type: str) -> None:
    """Add a media default only when the target package does not have one."""
    content_types_name = "[Content_Types].xml"
    if content_types_name not in files:
        return
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    root = ET.fromstring(files[content_types_name])
    if any(node.get("Extension", "").lower() == extension.lower() for node in root):
        return
    ET.SubElement(root, f"{{{ns}}}Default", Extension=extension, ContentType=content_type)
    files[content_types_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_part_content_type(
    files: dict[str, bytes], part_name: str, content_type: str
) -> None:
    """Register an added OPC part in ``[Content_Types].xml``.

    Headerless V2 templates do not have a ``header`` override. Adding the XML
    part and relationship without this override makes Word treat the package
    as incomplete even though the relationship itself is valid.
    """
    content_types_name = "[Content_Types].xml"
    if content_types_name not in files:
        return
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    root = ET.fromstring(files[content_types_name])
    normalized = "/" + part_name.lstrip("/")
    for node in root:
        if node.get("PartName", "").lstrip("/") == normalized.lstrip("/"):
            return
    ET.SubElement(
        root,
        f"{{{ns}}}Override",
        PartName=normalized,
        ContentType=content_type,
    )
    files[content_types_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def inject_header_profile(template_path: str | os.PathLike[str], profile_path: str | os.PathLike[str], output_path: str | os.PathLike[str]) -> Path:
    """Copy a DOCX and replace its existing header sections with one profile.

    Some BA/Reviu templates have two sections. The second section may point to
    ``header2.xml``; replacing only ``header1.xml`` makes the official header
    disappear halfway through the generated PDF. Reuse the profile header and
    its image relationships for every header part already present in target.
    """
    template_path = Path(template_path)
    profile_path = Path(profile_path)
    output_path = Path(output_path)
    with zipfile.ZipFile(template_path, "r") as body_zip, zipfile.ZipFile(profile_path, "r") as header_zip:
        body_infos = {item.filename: item for item in body_zip.infolist()}
        files = {name: body_zip.read(name) for name in body_infos}
        profile_names = set(header_zip.namelist())
        profile_header = header_zip.read("word/header1.xml")
        profile_header_rels = header_zip.read("word/_rels/header1.xml.rels") if "word/_rels/header1.xml.rels" in profile_names else None
        target_media = {
            name for name in body_infos if name.startswith("word/media/")
        }
        media_map = {}
        used_media = set(target_media)
        for name in sorted(profile_names):
            if not name.startswith("word/media/"):
                continue
            candidate = name
            if candidate in used_media:
                path = Path(name)
                stem = path.stem
                suffix = path.suffix
                index = 1
                candidate = f"word/media/header_profile_{stem}{suffix}"
                while candidate in used_media:
                    index += 1
                    candidate = f"word/media/header_profile_{stem}_{index}{suffix}"
            media_map[name] = candidate
            used_media.add(candidate)
        for name in ("word/header1.xml", "word/header2.xml"):
            if name in body_infos:
                files[name] = profile_header
        for name in ("word/_rels/header1.xml.rels", "word/_rels/header2.xml.rels"):
            needs_header2_rels = name == "word/_rels/header2.xml.rels" and "word/header2.xml" in body_infos
            if (name in body_infos or needs_header2_rels) and profile_header_rels is not None:
                files[name] = _rewrite_header_relationships(profile_header_rels, media_map)
        for name in profile_names:
            if name.startswith("word/media/"):
                destination = media_map[name]
                if destination not in files:
                    files[destination] = header_zip.read(name)

        # Headerless templates need a real part + document relationship +
        # section reference. Existing headers keep their original part names;
        # otherwise create the conventional header1.xml pair.
        header_parts = sorted(
            name for name in body_infos
            if name.startswith("word/header") and name.endswith(".xml")
        )
        header_name = header_parts[0] if header_parts else "word/header1.xml"
        header_rels_name = "word/_rels/" + Path(header_name).name + ".rels"
        files[header_name] = profile_header
        if profile_header_rels is not None:
            files[header_rels_name] = _rewrite_header_relationships(
                profile_header_rels, media_map
            )

        rels_name = "word/_rels/document.xml.rels"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        rels_root = ET.fromstring(files[rels_name])
        header_rel = None
        header_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        for relationship in rels_root:
            if relationship.get("Type") == header_type:
                header_rel = relationship
                break
        if header_rel is None:
            relationship_id = _next_relationship_id(rels_root)
            header_rel = ET.SubElement(
                rels_root,
                f"{{{rels_ns}}}Relationship",
                Id=relationship_id,
                Type=header_type,
                Target=Path(header_name).name,
            )
        else:
            relationship_id = header_rel.get("Id")
            header_rel.set("Target", Path(header_name).name)
        files[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
        files["word/document.xml"] = _ensure_header_reference(
            files["word/document.xml"], relationship_id
        )
        _ensure_part_content_type(
            files,
            header_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        )

        for media_name in media_map.values():
            suffix = Path(media_name).suffix.lower().lstrip(".")
            if suffix == "png":
                _ensure_content_type(files, "png", "image/png")
            elif suffix in {"jpg", "jpeg"}:
                _ensure_content_type(files, suffix, f"image/{suffix}")
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
