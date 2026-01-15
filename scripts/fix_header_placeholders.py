#!/usr/bin/env python3
import re
import zipfile
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

# 支持 {{ }} / {% %} / {# #}
PH_RE = re.compile(r"({{.*?}}|{%.+?%}|{#.+?#})", re.DOTALL)

SIMPLE_ALLOWED = {
    f"{{{W_NS}}}rPr",
    f"{{{W_NS}}}t",
    f"{{{W_NS}}}tab",
    f"{{{W_NS}}}br",
    f"{{{W_NS}}}cr",
}


def is_simple_run(r: etree._Element) -> bool:
    """只允许 rPr + 纯文本/换行/tab。出现字段、图片、对象就不动。"""
    for child in r:
        if child.tag not in SIMPLE_ALLOWED:
            return False
    return True


def run_text_and_map(p: etree._Element):
    """返回 paragraph 的可见文本，以及每个字符属于哪个 run index。"""
    runs = p.findall("./w:r", namespaces=NS)
    full = []
    char_to_run = []

    for i, r in enumerate(runs):
        for t in r.findall("./w:t", namespaces=NS):
            s = t.text or ""
            for ch in s:
                full.append(ch)
                char_to_run.append(i)

        if r.find("./w:tab", namespaces=NS) is not None:
            full.append("\t")
            char_to_run.append(i)
        if r.find("./w:br", namespaces=NS) is not None:
            full.append("\n")
            char_to_run.append(i)
        if r.find("./w:cr", namespaces=NS) is not None:
            full.append("\n")
            char_to_run.append(i)

    return "".join(full), char_to_run, runs


def clear_text_children(r: etree._Element):
    """清掉 run 里的 t/tab/br/cr，保留 rPr。"""
    for child in list(r):
        if child.tag in {f"{{{W_NS}}}t", f"{{{W_NS}}}tab", f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
            r.remove(child)


def set_run_text(r: etree._Element, text: str):
    clear_text_children(r)
    t = etree.SubElement(r, f"{{{W_NS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def merge_placeholders_in_paragraph(p: etree._Element) -> int:
    """合并一个段落里跨 run 的占位符。返回合并次数。"""
    merges = 0

    while True:
        full_text, char_to_run, runs = run_text_and_map(p)
        if not full_text:
            break

        m = PH_RE.search(full_text)
        if not m:
            break

        s, e = m.span()
        if s >= len(char_to_run) or e - 1 >= len(char_to_run):
            break

        rs = char_to_run[s]
        re_ = char_to_run[e - 1]
        if rs == re_:
            full_text = full_text[:s] + (" " * (e - s)) + full_text[e:]
            m2 = PH_RE.search(full_text)
            if not m2:
                break
            s, e = m2.span()
            rs = char_to_run[s]
            re_ = char_to_run[e - 1]
            if rs == re_:
                break

        if any((j >= len(runs)) or (not is_simple_run(runs[j])) for j in range(rs, re_ + 1)):
            break

        merged_text = "".join(
            "".join((t.text or "") for t in runs[j].findall("./w:t", namespaces=NS))
            + ("\t" if runs[j].find("./w:tab", namespaces=NS) is not None else "")
            + ("\n" if runs[j].find("./w:br", namespaces=NS) is not None else "")
            + ("\n" if runs[j].find("./w:cr", namespaces=NS) is not None else "")
            for j in range(rs, re_ + 1)
        )

        set_run_text(runs[rs], merged_text)

        for j in range(re_, rs, -1):
            p.remove(runs[j])

        merges += 1

    return merges


def fix_header_placeholders(input_docx: str, output_docx: str) -> None:
    with zipfile.ZipFile(input_docx, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    target_parts = [
        n for n in files.keys()
        if n.startswith("word/header") or n.startswith("word/footer")
    ]

    total_merges = 0
    for part in target_parts:
        root = etree.fromstring(files[part])
        for p in root.findall(".//w:p", namespaces=NS):
            total_merges += merge_placeholders_in_paragraph(p)
        files[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    with zipfile.ZipFile(output_docx, "w") as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"[OK] merged placeholder runs in header/footer: {total_merges}")
    print(f"[OK] wrote: {output_docx}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python fix_header_placeholders.py <input.docx> <output.docx>")
        raise SystemExit(2)
    fix_header_placeholders(sys.argv[1], sys.argv[2])
