"""
relink_pl.py — Relink Word template PL ke Excel (tanpa copy, hanya update mailMerge).
Dipanggil dari VBA ModDraftPaketPL.RelinkPL.

Usage:
    python relink_pl.py "D:/path/to/Excel.xlsm"

Logic:
  1. Detect folder dari path Excel
  2. Scan .docx/.docm di folder (skip *Merged*, ~*, *.bak*, *Header*)
  3. Detect template per file (BA PLJKK -> satu_data, Reviu -> list_reviu, Dokpil -> list_dokpil)
  4. Update word/settings.xml + word/_rels/settings.xml.rels via ZIP manipulation
"""
import os
import sys
import re
import zipfile
import shutil
import tempfile


WORD_SHEET_MAP = [
    ("Isi Reviu PLJKK",     "list_reviu"),
    ("BA PLJKK",            "satu_data"),
    ("Reviu DPP PLJKK",     "list_reviu"),
    ("Dokpil Full PLJKK",   "list_dokpil"),
    ("Undangan Full PLJK",  "satu_data"),
    ("BA PLPK",             "satu_data"),
    ("Reviu PLPK",          "list_reviu"),
    ("Dokpil Full PLPK",    "list_dokpil"),
]


def detect_sheet(filename: str) -> str:
    for kw, sheet in WORD_SHEET_MAP:
        if kw.lower() in filename.lower():
            return sheet
    return ""


def excel_to_file_uri(path: str) -> str:
    from urllib.parse import quote
    parts = path.replace("\\", "/").split("/")
    encoded = "/".join(quote(p) for p in parts)
    return "file:///" + encoded


def relink_word(word_path: str, excel_path: str, sheet_name: str) -> bool:
    excel_escaped = excel_path.replace("&", "&amp;")
    mail_merge = (
        '<w:mailMerge>'
        '<w:mainDocumentType w:val="formLetters"/>'
        '<w:linkToQuery/>'
        '<w:dataType w:val="native"/>'
        '<w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;'
        'User ID=Admin;'
        f'Data Source={excel_escaped};'
        'Mode=Read;'
        'Extended Properties=&quot;HDR=YES;IMEX=1&quot;;"/>'
        f'<w:query w:val="SELECT * FROM `{sheet_name}$`"/>'
        '</w:mailMerge>'
    )

    excel_uri = excel_to_file_uri(excel_path)
    new_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_uri}" TargetMode="External"/>'
        f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_uri}" TargetMode="External"/>'
        '</Relationships>'
    )

    try:
        with zipfile.ZipFile(word_path, 'r') as zf:
            files = {n: zf.read(n) for n in zf.namelist()}

        if "word/settings.xml" not in files:
            return False

        settings = files["word/settings.xml"].decode("utf-8")
        settings = re.sub(r'<w:mailMerge>.*?</w:mailMerge>', '', settings, flags=re.DOTALL)
        settings = settings.replace('</w:settings>', mail_merge + '</w:settings>')
        files["word/settings.xml"] = settings.encode("utf-8")
        files["word/_rels/settings.xml.rels"] = new_rels.encode("utf-8")

        tmp = word_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n, d in files.items():
                zout.writestr(n, d)
        os.replace(tmp, word_path)
        return True
    except Exception as e:
        print(f"[ERROR] {os.path.basename(word_path)}: {e}")
        return False


def main():
    if len(sys.argv) < 2:
        print("Usage: relink_pl.py <excel_path>")
        sys.exit(1)

    excel_path = os.path.abspath(sys.argv[1])
    folder = os.path.dirname(excel_path)

    if not os.path.isfile(excel_path):
        print(f"[ERROR] Excel tidak ada: {excel_path}")
        sys.exit(1)

    files = [
        f for f in os.listdir(folder)
        if f.lower().endswith((".docx", ".docm"))
        and "(merged)" not in f.lower()
        and not f.startswith("~")
        and ".bak" not in f.lower()
        and "header" not in f.lower()
        and "undangan" not in f.lower()
    ]

    ok, skip = 0, 0
    for f in files:
        sheet = detect_sheet(f)
        if not sheet:
            print(f"[SKIP] {f}")
            skip += 1
            continue
        word_path = os.path.join(folder, f)
        if relink_word(word_path, excel_path, sheet):
            print(f"[OK] {f} -> {sheet}")
            ok += 1
        else:
            print(f"[GAGAL] {f}")

    print(f"\nSelesai: {ok} ter-relink, {skip} di-skip.")


if __name__ == "__main__":
    main()
