#!/usr/bin/env python3
"""Patch the copied Konstruksi KAK/Uraian Singkat DOCX templates.

The patch is intentionally XML-level: it changes only the donor text/rows and
keeps the existing package parts, styles, images, relationships, and sections.
"""

from __future__ import annotations

import argparse
import copy
import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def text_of(node) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


def set_text(node, value: str) -> None:
    """Set visible text while retaining the first run's formatting."""
    # Target cells are being rewritten as a single logical block. Remove any
    # donor nested table (the copied KAK has a legacy personil table here) so
    # the generator can insert exactly one Excel-driven table later.
    for nested_table in node.xpath("./w:tbl", namespaces=NS):
        node.remove(nested_table)
    paragraphs = node.xpath("./w:p", namespaces=NS)
    if not paragraphs:
        paragraph = etree.SubElement(node, qn("p"))
        run = etree.SubElement(paragraph, qn("r"))
        text_node = etree.SubElement(run, qn("t"))
        text_node.set(f"{{{XML_NS}}}space", "preserve")
        text_node.text = value
        return

    first_paragraph = paragraphs[0]
    for extra in paragraphs[1:]:
        node.remove(extra)

    text_nodes = first_paragraph.xpath(".//w:t", namespaces=NS)
    if not text_nodes:
        run = etree.SubElement(first_paragraph, qn("r"))
        text_node = etree.SubElement(run, qn("t"))
        text_nodes = [text_node]
    text_nodes[0].text = value
    if value[:1].isspace() or value[-1:].isspace():
        text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
    for text_node in text_nodes[1:]:
        text_node.text = ""


def set_paragraph_text(paragraph, value: str) -> None:
    text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
    if not text_nodes:
        run = etree.SubElement(paragraph, qn("r"))
        text_nodes = [etree.SubElement(run, qn("t"))]
    text_nodes[0].text = value
    if value[:1].isspace() or value[-1:].isspace():
        text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
    for text_node in text_nodes[1:]:
        text_node.text = ""


def paragraphs(root):
    return root.xpath(".//w:p", namespaces=NS)


def replace_paragraph_start(root, prefix: str, value: str) -> bool:
    for paragraph in paragraphs(root):
        if text_of(paragraph).strip().startswith(prefix):
            set_paragraph_text(paragraph, value)
            return True
    return False


def find_main_table(root):
    tables = root.xpath(".//w:tbl", namespaces=NS)
    if not tables:
        raise ValueError("DOCX tidak memiliki tabel utama")
    return tables[0]


def rows(table):
    return table.xpath("./w:tr", namespaces=NS)


def row_label(row) -> str:
    cells = row.xpath("./w:tc", namespaces=NS)
    return text_of(cells[0]).strip() if cells else ""


def find_row(table, label: str):
    for row in rows(table):
        if row_label(row) == label:
            return row
    raise ValueError(f"Baris KAK tidak ditemukan: {label}")


def set_row(table, label: str, value: str, new_label: str | None = None) -> None:
    try:
        row = find_row(table, label)
    except ValueError:
        if not new_label:
            raise
        row = find_row(table, new_label)
    cells = row.xpath("./w:tc", namespaces=NS)
    if len(cells) < 2:
        raise ValueError(f"Baris KAK bukan dua kolom: {label}")
    set_text(cells[0], new_label or label)
    set_text(cells[1], value)


def insert_row_after(table, anchor_label: str, label: str, value: str) -> None:
    existing = next((row for row in rows(table) if row_label(row) == label), None)
    if existing is not None:
        cells = existing.xpath("./w:tc", namespaces=NS)
        set_text(cells[0], label)
        set_text(cells[1], value)
        return
    anchor = find_row(table, anchor_label)
    new_row = copy.deepcopy(anchor)
    cells = new_row.xpath("./w:tc", namespaces=NS)
    set_text(cells[0], label)
    set_text(cells[1], value)
    table.insert(table.index(anchor) + 1, new_row)


def insert_content_row_after_heading(table, heading_label: str, label: str, value: str) -> None:
    """Insert a normal two-column row after a one-cell section heading."""
    existing = next((row for row in rows(table) if row_label(row) == label), None)
    if existing is not None:
        cells = existing.xpath("./w:tc", namespaces=NS)
        set_text(cells[0], label)
        set_text(cells[1], value)
        return
    heading = find_row(table, heading_label)
    candidates = [row for row in rows(table) if len(row.xpath("./w:tc", namespaces=NS)) >= 2]
    if not candidates:
        raise ValueError("Tidak ada baris dua kolom untuk dijadikan template")
    new_row = copy.deepcopy(candidates[0])
    cells = new_row.xpath("./w:tc", namespaces=NS)
    set_text(cells[0], label)
    set_text(cells[1], value)
    table.insert(table.index(heading) + 1, new_row)


def patch_kak(root) -> None:
    table = find_main_table(root)

    replacements = {
        "PEMERINTAH KABUPATEN TAPIN": "PEMERINTAH «KABUPATEN_KOTA»",
        "«NAMA_SKPD_SINGKAT» KABUPATEN TAPIN": "«NAMA_SKPD_SINGKAT» «KABUPATEN_KOTA»",
        "Rantau, «TANGGAL_KAK_HPS»": "«KOTA_DOKUMEN», «TANGGAL_KAK_HPS»",
    }
    for paragraph in paragraphs(root):
        text = text_of(paragraph)
        for old, new in replacements.items():
            if old in text:
                set_paragraph_text(paragraph, text.replace(old, new))
                break

    set_row(
        table,
        "Latar Belakang",
        "Pekerjaan «NAMA_PAKET_LENGKAP» diperlukan untuk memenuhi kebutuhan layanan/infrastruktur masyarakat melalui pelaksanaan pekerjaan fisik yang fungsional, aman, bermutu, tepat waktu, dan sesuai kondisi lokasi serta dokumen teknis paket.",
    )
    set_row(
        table,
        "Maksud dan Tujuan",
        "Maksud: menjadi acuan pelaksanaan pekerjaan konstruksi. Tujuan: menghasilkan pekerjaan fisik sesuai lingkup, volume, mutu, waktu, keselamatan konstruksi, dan ketentuan kontrak sehingga dapat diperiksa dan diserahterimakan.",
    )
    set_row(
        table,
        "Sasaran",
        "Terlaksananya «NAMA_PAKET_LENGKAP» dan tersedianya hasil pekerjaan fisik yang memenuhi spesifikasi teknis, persyaratan mutu, keselamatan, dan fungsi yang ditetapkan.",
    )
    set_row(table, "Lokasi Kegiatan", "«LOKASI_PEKERJAAN», «KABUPATEN_KOTA».")
    set_row(
        table,
        "Sumber Pendanaan",
        "Pagu anggaran pekerjaan sebesar Rp. «PAGU_ANGKA_FORMAT» ( «PAGU_TERBILANG» ), «SUMBER_DANA_DETAIL».",
    )
    set_row(
        table,
        "Nama dan Organisasi PPK",
        "Nama PPK: «NAMA_PPK»; NIP: «NIP_PPK»; Jabatan: «JABATAN_PPK»; Satuan kerja: «NAMA_SKPD»." ,
    )
    set_row(
        table,
        "Data Dasar",
        "Data dasar meliputi kondisi lokasi, RAB/HPS, gambar, spesifikasi teknis, dan dokumen pendukung lain yang disediakan PPK sesuai kebutuhan paket.",
    )
    set_row(
        table,
        "Standar Teknis",
        "Pekerjaan wajib mengikuti gambar, spesifikasi teknis, RAB, SNI/standar teknis yang relevan, serta peraturan perundang-undangan yang berlaku.",
    )
    insert_row_after(
        table,
        "Standar Teknis",
        "Pengendalian Mutu dan Pengujian",
        "Penyedia melaksanakan pemeriksaan bahan, metode kerja, pengujian yang dipersyaratkan, dokumentasi hasil uji, dan perbaikan atas pekerjaan yang tidak memenuhi persyaratan.",
    )
    insert_row_after(
        table,
        "Pengendalian Mutu dan Pengujian",
        "Keselamatan Konstruksi, K3, dan Lingkungan",
        "Penyedia wajib menerapkan SMKK/K3, mengendalikan risiko pekerjaan, menjaga keselamatan pekerja dan masyarakat, mengatur lalu lintas/akses lokasi bila diperlukan, serta mengelola dampak lingkungan sesuai ketentuan.",
    )
    set_row(
        table,
        "Studi-Studi Terdahulu",
        "Dokumen teknis paket yang tersedia dan relevan menjadi acuan; apabila data belum tersedia, penyedia wajib menyampaikan kebutuhan klarifikasi kepada PPK sebelum pelaksanaan bagian pekerjaan terkait.",
        "Dokumen Teknis Paket",
    )
    set_row(
        table,
        "Referensi Hukum",
        "Pelaksanaan mengikuti peraturan pengadaan barang/jasa, jasa konstruksi, keselamatan konstruksi, bangunan/pekerjaan terkait, standar teknis, dan ketentuan lain yang berlaku serta relevan dengan paket.",
    )
    set_row(
        table,
        "Lingkup Kegiatan",
        "Lingkup pekerjaan meliputi: «LINGKUP_PEKERJAAN». Rincian volume, satuan, gambar, dan spesifikasi teknis mengikuti dokumen teknis paket.",
    )
    set_row(
        table,
        "Keluaran",
        "Keluaran berupa hasil pekerjaan fisik sesuai lingkup, hasil pengujian/pemeriksaan, dokumentasi progres dan hasil akhir, as-built drawing bila dipersyaratkan, serta dokumen pemeriksaan dan serah terima.",
    )
    set_row(
        table,
        "Peralatan, Material, Personel dan Fasilitas dari PPK",
        "PPK menyediakan data, dokumen teknis, akses/koordinasi lokasi, dan fasilitas koordinasi yang secara wajar berada dalam kewenangan PPK. Material, personel, dan peralatan pelaksanaan disediakan penyedia kecuali ditentukan lain dalam dokumen kontrak.",
    )
    set_row(
        table,
        "Peralatan dan Material dari Penyedia Jasa Konsultansi",
        "Penyedia menyediakan personel, material, peralatan, perlengkapan kerja, dan fasilitas pendukung yang diperlukan. Daftar minimal personel dan peralatan mengikuti data paket berikut.",
        "Sumber Daya dari Penyedia",
    )
    set_row(
        table,
        "Lingkup Kewenangan Penyedia Jasa",
        "Penyedia bertanggung jawab melaksanakan pekerjaan sesuai dokumen kontrak, mengatur sumber daya, berkoordinasi dengan PPK, menjaga mutu dan keselamatan, melaporkan progres, serta memperbaiki pekerjaan yang tidak sesuai.",
        "Tanggung Jawab Penyedia",
    )
    set_row(
        table,
        "Jangka Waktu Penyelesaian Kegiatan",
        "Jangka waktu pelaksanaan adalah «JANGKA_WAKTU_HARI» («JANGKA_WAKTU_TERBILANG») hari kalender sesuai ketentuan kontrak/SPMK dan jadwal yang disetujui.",
        "Jangka Waktu Pelaksanaan",
    )
    insert_row_after(
        table,
        "Jangka Waktu Pelaksanaan",
        "Masa Pemeliharaan",
        "Masa pemeliharaan pekerjaan adalah «MASA_PEMELIHARAAN_HARI» hari kalender sejak serah terima pertama, atau sesuai ketentuan kontrak. Penyedia wajib memperbaiki kerusakan/cacat yang menjadi tanggung jawabnya.",
    )
    set_row(
        table,
        "Kebutuhan Personel Minimal",
        "«TABEL_PERSONEL_PK»",
    )
    set_row(
        table,
        "Jadwal Tahapan Pelaksanaan Kegiatan",
        "Penyedia menyusun jadwal pelaksanaan, urutan pekerjaan, dan milestone berdasarkan jangka waktu; jadwal menjadi alat pengendalian progres dan dikoordinasikan dengan PPK.",
    )
    insert_content_row_after_heading(
        table,
        "Laporan*)",
        "Pelaporan dan Dokumentasi",
        "Penyedia menyampaikan laporan progres, dokumentasi lapangan, hasil pengujian, kendala dan tindak lanjut, serta dokumen akhir yang dipersyaratkan.",
    )
    set_row(
        table,
        "Produksi Dalam Negeri",
        "Penggunaan material, peralatan, dan produk dalam negeri diutamakan sesuai ketentuan dan ketersediaan yang memenuhi persyaratan teknis.",
        "Produk Dalam Negeri",
    )
    set_row(
        table,
        "Persyaratan Kerja Sama",
        "Subkontrak/kerja sama hanya dilakukan sesuai ketentuan dokumen pemilihan dan kontrak. Penyedia tetap bertanggung jawab penuh atas mutu, waktu, keselamatan, dan seluruh hasil pekerjaan.",
        "Subkontrak/Kerja Sama",
    )
    set_row(
        table,
        "Alih Pengetahuan",
        "Pengukuran prestasi dilakukan berdasarkan pekerjaan yang benar-benar selesai dan diterima. Pembayaran mengikuti jenis kontrak, sistem pembayaran, dokumen kontrak, dan hasil pemeriksaan PPK.",
        "Pengukuran dan Pembayaran",
    )

    # The marker is consumed by the generator and replaced with an Excel-driven
    # nested table. It deliberately lives in the existing two-column KAK table.
    equipment_row = find_row(table, "Sumber Daya dari Penyedia")
    set_text(equipment_row.xpath("./w:tc", namespaces=NS)[1], "Penyedia menyediakan personel, material, peralatan, perlengkapan kerja, dan fasilitas pendukung yang diperlukan.\n«TABEL_PERALATAN_PK»")


def patch_uraian_singkat(root) -> None:
    replacements = {
        "Rantau, «BULAN_TAHUN_KAK_HPS»": "«KOTA_DOKUMEN», «BULAN_TAHUN_KAK_HPS»",
    }
    for paragraph in paragraphs(root):
        text = text_of(paragraph)
        for old, new in replacements.items():
            if old in text:
                set_paragraph_text(paragraph, text.replace(old, new))

    replace_paragraph_start(
        root,
        "Pekerjaan «NAMA_PAKET_LENGKAP»",
        "Pekerjaan «NAMA_PAKET_LENGKAP» merupakan pekerjaan konstruksi untuk menghasilkan hasil fisik yang fungsional, aman, bermutu, dan sesuai dengan kebutuhan lokasi serta dokumen teknis paket.",
    )
    replace_paragraph_start(
        root,
        "Lingkup pekerjaan diawali",
        "Lingkup pekerjaan meliputi «LINGKUP_PEKERJAAN». Rincian volume, satuan, gambar, dan spesifikasi teknis mengikuti dokumen teknis paket serta hasil klarifikasi yang disetujui PPK.",
    )
    replace_paragraph_start(
        root,
        "Keluaran pekerjaan",
        "Keluaran pekerjaan berupa hasil fisik, hasil pemeriksaan/pengujian, dokumentasi, dan dokumen serah terima yang dipersyaratkan. Jangka waktu pelaksanaan adalah «JANGKA_WAKTU_HARI» («JANGKA_WAKTU_TERBILANG») hari kalender. Masa pemeliharaan pekerjaan adalah «MASA_PEMELIHARAAN_HARI» hari kalender sesuai ketentuan kontrak.",
    )

    table = root.xpath(".//w:tbl", namespaces=NS)[0]
    desired = [
        ("PROGRAM", "«PROGRAM»"),
        ("KEGIATAN", "«KEGIATAN»"),
        ("SUB KEGIATAN", "«SUB_KEGIATAN»"),
        ("PEKERJAAN", "«NAMA_PAKET_LENGKAP»"),
        ("RUP", "«KODE_RUP»"),
        ("LOKASI", "«LOKASI_PEKERJAAN», «KABUPATEN_KOTA»"),
        ("SUMBER DANA", "«SUMBER_DANA_DETAIL»"),
        ("TAHUN ANGGARAN", "«TAHUN_ANGGARAN»"),
        ("PAGU", "Rp. «PAGU_ANGKA_FORMAT» («PAGU_TERBILANG»)"),
        ("JANGKA WAKTU", "«JANGKA_WAKTU_HARI» («JANGKA_WAKTU_TERBILANG») hari kalender"),
    ]
    existing = rows(table)
    while len(existing) < len(desired):
        new_row = copy.deepcopy(existing[-1])
        table.append(new_row)
        existing = rows(table)
    for row, (label, value) in zip(existing, desired):
        cells = row.xpath("./w:tc", namespaces=NS)
        if len(cells) >= 3:
            set_text(cells[0], label)
            set_text(cells[1], ":")
            set_text(cells[2], value)
    for row in existing[len(desired):]:
        table.remove(row)


def patch_header_uraian(root) -> None:
    header_paragraphs = paragraphs(root)
    if len(header_paragraphs) >= 1:
        set_paragraph_text(header_paragraphs[0], "PEMERINTAH «KABUPATEN_KOTA»")
    if len(header_paragraphs) >= 2:
        set_paragraph_text(header_paragraphs[1], "«NAMA_SKPD_SINGKAT»")
    if len(header_paragraphs) >= 3:
        set_paragraph_text(header_paragraphs[2], "«ALAMAT_SKPD», «KABUPATEN_KOTA»")


def patch_docx(path: Path, patcher, auxiliary_patchers=None) -> tuple[int, int, int]:
    temp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as source:
        payloads = {"word/document.xml": source.read("word/document.xml")}
        root = etree.fromstring(payloads["word/document.xml"])
        before = len(source.infolist())
        patcher(root)
        payloads["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        for part_name, part_patcher in (auxiliary_patchers or {}).items():
            if part_name not in source.namelist():
                continue
            part_root = etree.fromstring(source.read(part_name))
            part_patcher(part_root)
            payloads[part_name] = etree.tostring(
                part_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                target.writestr(item, payloads.get(item.filename, source.read(item.filename)))
    os.replace(temp, path)
    with zipfile.ZipFile(path, "r") as check:
        if check.testzip() is not None:
            raise ValueError(f"DOCX hasil patch rusak: {path}")
        after = len(check.infolist())
    return before, after, len(payloads["word/document.xml"])


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, help="Folder V2 - Template PPK PL")
    args = parser.parse_args()
    root = Path(args.root).resolve()
    targets = [
        (root / "Konstruksi" / "1. KAK.docx", patch_kak, {}),
        (
            root / "Konstruksi" / "2. U_Singkat.docx",
            patch_uraian_singkat,
            {"word/header1.xml": patch_header_uraian},
        ),
    ]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for path, patcher, auxiliary_patchers in targets:
        if not path.is_file():
            raise FileNotFoundError(path)
        backup = path.with_name(f"{path.stem}.before-konstruksi-{stamp}{path.suffix}")
        shutil.copy2(path, backup)
        before, after, size = patch_docx(path, patcher, auxiliary_patchers)
        print(f"PATCHED|{path}|backup={backup.name}|entries={before}->{after}|xml_bytes={size}")


if __name__ == "__main__":
    main()
