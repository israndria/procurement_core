"""Repair Revaluasi workbook XML without reserializing the XLSM package."""
from __future__ import annotations

import html
import os
import re
import shutil
import sys
import tempfile
import zipfile


def replace_cell(xml: str, ref: str, formula: str, cached: str, style: str, cell_type: str = "str") -> str:
    value = html.escape(str(cached), quote=False)
    f = html.escape(formula, quote=False)
    type_attr = f' t="{cell_type}"' if cell_type else ""
    replacement = f'<c r="{ref}" s="{style}"{type_attr}><f>{f}</f><v>{value}</v></c>'
    pattern = rf'<c\s+r="{re.escape(ref)}"[^>]*>.*?</c>'
    updated, count = re.subn(pattern, replacement, xml, count=1, flags=re.S)
    if count != 1:
        raise ValueError(f"Cell {ref} tidak ditemukan")
    return updated


def repair(source: str, target: str) -> None:
    with zipfile.ZipFile(source, "r") as zin:
        payloads = {name: zin.read(name) for name in zin.namelist()}
        infos = zin.infolist()

    s12 = payloads["xl/worksheets/sheet12.xml"].decode("utf-8")
    s38 = payloads["xl/worksheets/sheet38.xml"].decode("utf-8")

    # satu_data: semua angka/harga Revaluasi mengambil hasil final klarifikasi/nego.
    s12 = replace_cell(s12, "AE2", "=CZ2", "98,90%", "503")
    s12 = replace_cell(s12, "AH2", 'TEXT(BM2,"Rp. #.##0,00")', "Rp. 989.000.000,00", "503")
    s12 = replace_cell(s12, "AK2", 'TEXT(BQ2,"Rp. #.##0,00")', "Rp. 989.000.000,00", "503")
    s12 = replace_cell(
        s12, "EY2",
        'IF(AND(\'7.2 Dengan Nego\'!T3>=80%,\'7.2 Dengan Nego\'!T3<=100%),"WAJAR","TIDAK WAJAR")',
        "WAJAR", "499",
    )
    s12 = replace_cell(s12, "EZ2", '=IF(EY2="WAJAR","MEMENUHI","TIDAK MEMENUHI")', "MEMENUHI", "499")

    # Evaluasi Harga: wajar hanya pada rentang 80% s.d. 100% HPS.
    # C11/C12 harus memakai harga final hasil klarifikasi/nego, bukan nilai
    # awal dari 0. Input BA yang membuat rasio Revaluasi bisa >100% palsu.
    s38 = replace_cell(s38, "C7", "'7. BA Klarifikasi HS'!H51", "989000000", "682", cell_type="")
    s38 = replace_cell(s38, "D7", "'7. BA Klarifikasi HS'!G57", "989000000", "682", cell_type="")
    for row in (7, 8, 9):
        s38 = replace_cell(
            s38, f"F{row}",
            'IF(E{0}="-","-",IF(OR(E{0}<80%,E{0}>100%),"TIDAK WAJAR","WAJAR"))'.format(row),
            "WAJAR", "684",
        )
        s38 = replace_cell(
            s38, f"G{row}",
            f'IF(F{row}="WAJAR","MEMENUHI","TIDAK MEMENUHI")',
            "MEMENUHI", "682",
        )

    payloads["xl/worksheets/sheet12.xml"] = s12.encode("utf-8")
    payloads["xl/worksheets/sheet38.xml"] = s38.encode("utf-8")
    os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsm", dir=os.path.dirname(os.path.abspath(target)))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in infos:
                zout.writestr(info, payloads[info.filename])
        os.replace(tmp, target)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: repair_revaluasi_template.py SOURCE.xlsm TARGET.xlsm")
    repair(sys.argv[1], sys.argv[2])
