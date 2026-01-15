#!/usr/bin/env python3
import zipfile
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def get_run_text(r):
    ts = r.xpath(".//w:t", namespaces=NS)
    return "".join(t.text or "" for t in ts)


def set_run_text(r, text):
    ts = r.xpath(".//w:t", namespaces=NS)
    if not ts:
        t = etree.Element(f"{{{W_NS}}}t")
        t.text = text
        r.append(t)
        return
    ts[0].text = text
    for t in ts[1:]:
        t.getparent().remove(t)


def merge_placeholder_runs_in_paragraph(p, placeholder: str) -> int:
    runs = p.xpath(".//w:r", namespaces=NS)
    changed = 0
    i = 0
    while i < len(runs):
        acc = ""
        end = None
        for j in range(i, min(i + 12, len(runs))):
            acc += get_run_text(runs[j])
            if placeholder in acc:
                if not acc.startswith(placeholder) or acc != placeholder:
                    break
                end = j
                set_run_text(runs[i], placeholder)
                for k in range(i + 1, end + 1):
                    set_run_text(runs[k], "")
                changed += 1
                i = end
                break
        i += 1
    return changed


def fix_cover_title_placeholder(docx_in: str, docx_out: str) -> None:
    placeholder = "{{software_name}}"
    with zipfile.ZipFile(docx_in, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    part = "word/document.xml"
    root = etree.fromstring(files[part])

    changed = 0
    for p in root.xpath(".//w:p", namespaces=NS):
        changed += merge_placeholder_runs_in_paragraph(p, placeholder)

    files[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    with zipfile.ZipFile(docx_out, "w") as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"[OK] merged cover title placeholder runs: {changed}")
    print(f"[OK] wrote: {docx_out}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python fix_cover_title_placeholder.py <input.docx> <output.docx>")
        raise SystemExit(2)
    fix_cover_title_placeholder(sys.argv[1], sys.argv[2])
