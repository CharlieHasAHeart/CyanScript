#!/usr/bin/env python3
import zipfile
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def paragraph_text(p):
    return "".join(t.text or "" for t in p.findall(".//w:t", namespaces=NS))


def fix_main_content_placeholder(docx_in: str, docx_out: str) -> None:
    with zipfile.ZipFile(docx_in, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    part = "word/document.xml"
    root = etree.fromstring(files[part])

    fixed = 0
    for p in root.findall(".//w:p", namespaces=NS):
        txt = paragraph_text(p)
        if "main_content" in txt:
            pPr = p.find("./w:pPr", namespaces=NS)
            pPr_copy = etree.fromstring(etree.tostring(pPr)) if pPr is not None else None

            for child in list(p):
                p.remove(child)

            if pPr_copy is not None:
                p.append(pPr_copy)

            r = etree.SubElement(p, f"{{{W_NS}}}r")
            t = etree.SubElement(r, f"{{{W_NS}}}t")
            t.text = "{{main_content}}"

            fixed += 1
            break

    files[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    with zipfile.ZipFile(docx_out, "w") as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"[OK] forced rebuilt main_content paragraph: {fixed}")
    print(f"[OK] wrote: {docx_out}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python fix_main_content_placeholder.py <input.docx> <output.docx>")
        raise SystemExit(2)
    fix_main_content_placeholder(sys.argv[1], sys.argv[2])
