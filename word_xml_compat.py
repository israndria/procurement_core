"""Open XML compatibility helpers for Word documents.

Some lxml round-trips collapse Word's namespace declarations to ``ns0`` while
leaving the original ``mc:Ignorable`` prefix list. Word treats those documents
as corrupted even though the ZIP and XML are otherwise well-formed.
"""

from __future__ import annotations

import os
import tempfile
import zipfile

from lxml import etree


MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Namespace map emitted by Word for the document part. Keeping the complete
# map matters because mc:Ignorable can reference prefixes that have no element
# at the moment but are still required by Word's compatibility parser.
WORD_DOCUMENT_NSMAP = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "mc": MC_NS,
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o": "urn:schemas-microsoft-com:office:office",
    "oel": "http://schemas.microsoft.com/office/2019/extlst",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": MAIN_NS,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}


def normalize_word_document_xml(document_xml: bytes) -> bytes:
    """Restore Word namespace declarations after an lxml round-trip.

    Return the original bytes when the root already has the canonical ``w``
    prefix and all prefixes listed in ``mc:Ignorable`` are declared.
    """

    root = etree.fromstring(document_xml)
    ignorable = root.get(f"{{{MC_NS}}}Ignorable", "").split()
    if root.nsmap.get("w") == MAIN_NS and all(prefix in root.nsmap for prefix in ignorable):
        return document_xml

    nsmap = dict(WORD_DOCUMENT_NSMAP)
    # Preserve non-generated extensions used by a future Word version.
    for prefix, uri in root.nsmap.items():
        if prefix and not prefix.startswith("ns") and prefix not in nsmap:
            nsmap[prefix] = uri

    repaired = etree.Element(root.tag, nsmap=nsmap)
    for key, value in root.attrib.items():
        repaired.set(key, value)
    repaired.text = root.text
    repaired.tail = root.tail
    for child in root:
        repaired.append(child)

    return etree.tostring(
        repaired,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def normalize_word_document_xml_in_zip(docx_path: str | os.PathLike[str]) -> bool:
    """Normalize ``word/document.xml`` in a DOCX/DOCM atomically.

    Return ``True`` only when the package changed. The original ZIP is never
    opened for writing, so a failed rewrite leaves it untouched.
    """

    docx_path = os.fspath(docx_path)
    temp_fd, temp_path = tempfile.mkstemp(
        prefix=".word-xml-",
        suffix=".tmp",
        dir=os.path.dirname(os.path.abspath(docx_path)),
    )
    os.close(temp_fd)
    changed = False
    try:
        with zipfile.ZipFile(docx_path, "r") as source:
            if "word/document.xml" not in source.namelist():
                return False
            original = source.read("word/document.xml")
            repaired = normalize_word_document_xml(original)
            changed = repaired != original
            if not changed:
                return False
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    payload = repaired if item.filename == "word/document.xml" else source.read(item.filename)
                    target.writestr(item, payload)
        os.replace(temp_path, docx_path)
        return True
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
