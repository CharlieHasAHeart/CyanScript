#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
template_selfcheck.py

A docx "lint" script for docxtpl templates / generated outputs.

What it checks (high-signal issues):
1) RUN_SPLIT: Jinja placeholders ({{}}, {%%}, {##}) split across multiple runs.
   This is a common root-cause for docxtpl producing OOXML that renders on one machine
   but becomes blank or corrupted on another Word/WPS version.
2) BODY_STYLE_LOCATION: body-style paragraphs accidentally placed in header/footer/table/textbox.
3) EXTERNAL_RELS: relationships with TargetMode="External" (may trigger "external references" prompt).
4) FIELD_EXTERNAL: field instructions like INCLUDETEXT / INCLUDEPICTURE / LINK / DDEAUTO.
5) EMBEDDED_OBJECT: OLE/embeddings/ActiveX parts that often cause compatibility prompts.

Usage:
  python template_selfcheck.py file.docx
  python template_selfcheck.py file.docx --mode template
  python template_selfcheck.py file.docx --mode output

Exit codes:
  0: ok
  2: issues found
  1: fatal error
"""
import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Iterable

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

PLACEHOLDER_RE = re.compile(r"({{.*?}}|{%.*?%}|{#.*?#})", flags=re.DOTALL)

EXTERNAL_FIELD_KEYWORDS = ("INCLUDETEXT", "INCLUDEPICTURE", "LINK", "DDEAUTO", "DDE")

RUN_SPECIAL_TEXT = {
    "tab": "\t",
    "br": "\n",
    "cr": "\n",
    "noBreakHyphen": "\u2011",
    "softHyphen": "\u00ad",
}

DEFAULT_BODY_STYLE_NAMES = {"正文", "Normal"}


def read_xml(zf: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        data = zf.read(name)
    except KeyError:
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def iter_xml_parts(zf: zipfile.ZipFile) -> Iterable[Tuple[str, bytes]]:
    """
    Iterate all XML parts under word/ (including document, headers, footers, footnotes, etc.).
    """
    for name in zf.namelist():
        if name.startswith("word/") and name.endswith(".xml"):
            yield name, zf.read(name)


def build_parent_map(root: ET.Element) -> Dict[ET.Element, ET.Element]:
    parent_map: Dict[ET.Element, ET.Element] = {}
    for parent in root.iter():
        for child in parent:
            parent_map[child] = parent
    return parent_map


def has_ancestor(node: ET.Element, parent_map: Dict[ET.Element, ET.Element], tag_suffixes: List[str]) -> bool:
    cur = node
    while cur in parent_map:
        cur = parent_map[cur]
        for suf in tag_suffixes:
            if cur.tag.endswith(suf):
                return True
    return False


def paragraph_plain_text(p: ET.Element) -> str:
    """
    Visible text approximation for printing diagnostics.
    """
    parts: List[str] = []
    for r in p.findall(".//w:r", NS):
        for t in r.findall(".//w:t", NS):
            if t.text:
                parts.append(t.text)
        for child in list(r):
            local = child.tag.split("}")[-1]
            if local in RUN_SPECIAL_TEXT:
                parts.append(RUN_SPECIAL_TEXT[local])
    return "".join(parts).strip()


def run_text_streams(p: ET.Element) -> Tuple[List[str], List[int]]:
    """
    Return (run_texts, char_to_run_index).
    Includes w:t and common special elements (tab, br, etc.) so placeholder splitting isn't missed.
    """
    run_texts: List[str] = []
    char_to_run: List[int] = []

    runs = p.findall(".//w:r", NS)
    for run_idx, r in enumerate(runs):
        buf: List[str] = []

        for child in list(r):
            local = child.tag.split("}")[-1]
            if local == "t":
                if child.text:
                    buf.append(child.text)
            elif local in RUN_SPECIAL_TEXT:
                buf.append(RUN_SPECIAL_TEXT[local])
            elif local == "instrText":
                if child.text:
                    buf.append(child.text)

        s = "".join(buf)
        run_texts.append(s)
        char_to_run.extend([run_idx] * len(s))

    return run_texts, char_to_run


def load_style_id_to_name(zf: zipfile.ZipFile) -> Dict[str, str]:
    """
    Parse word/styles.xml to map styleId -> display name.
    If not present, return {}.
    """
    root = read_xml(zf, "word/styles.xml")
    if root is None:
        return {}

    m: Dict[str, str] = {}
    for st in root.findall(".//w:style", NS):
        style_id = st.get(f"{{{NS['w']}}}styleId")
        if not style_id:
            continue
        name_el = st.find("./w:name", NS)
        if name_el is None:
            continue
        name_val = name_el.get(f"{{{NS['w']}}}val")
        if name_val:
            m[style_id] = name_val
    return m


def resolve_paragraph_style_name(p: ET.Element, style_id_to_name: Dict[str, str]) -> Optional[str]:
    p_style = p.find("./w:pPr/w:pStyle", NS)
    if p_style is None:
        return None
    style_id = p_style.get(f"{{{NS['w']}}}val")
    if not style_id:
        return None
    return style_id_to_name.get(style_id, style_id)


def check_run_split_placeholders(
    p: ET.Element,
    part_name: str,
    p_index: int,
) -> List[dict]:
    run_texts, char_to_run = run_text_streams(p)
    if not run_texts:
        return []
    full_text = "".join(run_texts)
    if ("{{" not in full_text and "{%" not in full_text and "{#" not in full_text) or (
        "}}" not in full_text and "%}" not in full_text and "#}" not in full_text
    ):
        return []

    issues: List[dict] = []
    for m in PLACEHOLDER_RE.finditer(full_text):
        start, end = m.span()
        if start >= len(char_to_run) or end - 1 >= len(char_to_run):
            continue
        run_start = char_to_run[start]
        run_end = char_to_run[end - 1]
        if run_start != run_end:
            issues.append(
                {
                    "type": "RUN_SPLIT",
                    "part": part_name,
                    "p_index": p_index,
                    "placeholder": m.group(0).replace("\n", "\\n"),
                    "text": paragraph_plain_text(p) or "[空段落]",
                    "run_start": run_start + 1,
                    "run_end": run_end + 1,
                }
            )
    return issues


def check_body_style_location(
    p: ET.Element,
    part_name: str,
    p_index: int,
    parent_map: Dict[ET.Element, ET.Element],
    style_id_to_name: Dict[str, str],
    body_style_names: set,
) -> Optional[dict]:
    style_name = resolve_paragraph_style_name(p, style_id_to_name)
    if style_name is None or style_name not in body_style_names:
        return None
    if not has_ancestor(p, parent_map, ["hdr", "ftr", "tbl", "txbxContent"]):
        return None

    return {
        "type": "BODY_STYLE_LOCATION",
        "part": part_name,
        "p_index": p_index,
        "style": style_name,
        "location": "header/footer/table/textbox",
        "text": paragraph_plain_text(p) or "[空段落]",
    }


def check_external_rels(zf: zipfile.ZipFile) -> List[dict]:
    issues: List[dict] = []
    for name in zf.namelist():
        if not name.startswith("word/_rels/") or not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(zf.read(name))
        except ET.ParseError:
            continue
        for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            target_mode = rel.get("TargetMode")
            target = rel.get("Target")
            if target_mode == "External":
                issues.append({"type": "EXTERNAL_RELS", "part": name, "target": target})
    return issues


def check_external_fields(root: ET.Element, part_name: str) -> List[dict]:
    issues: List[dict] = []
    for instr in root.findall(".//w:instrText", NS):
        text = instr.text or ""
        upper = text.upper()
        if any(k in upper for k in EXTERNAL_FIELD_KEYWORDS):
            issues.append({"type": "FIELD_EXTERNAL", "part": part_name, "text": text.strip()})
    return issues


def check_embedded_objects(part_name: str) -> Optional[dict]:
    if part_name.startswith("word/embeddings/") or part_name.startswith("word/activeX/"):
        return {"type": "EMBEDDED_OBJECT", "part": part_name}
    return None


def run_checks(docx: Path, mode: str, body_style_names: set, max_issues: int) -> int:
    issues: List[dict] = []

    with zipfile.ZipFile(docx, "r") as zf:
        style_id_to_name = load_style_id_to_name(zf)
        if mode in ("template", "all"):
            for part_name, data in iter_xml_parts(zf):
                root = read_xml(zf, part_name)
                if root is None:
                    continue
                parent_map = build_parent_map(root)
                for idx, p in enumerate(root.findall(".//w:p", NS)):
                    issues.extend(check_run_split_placeholders(p, part_name, idx))
                    body_issue = check_body_style_location(
                        p, part_name, idx, parent_map, style_id_to_name, body_style_names
                    )
                    if body_issue:
                        issues.append(body_issue)
                issues.extend(check_external_fields(root, part_name))

        if mode in ("output", "all"):
            issues.extend(check_external_rels(zf))

        for name in zf.namelist():
            embedded = check_embedded_objects(name)
            if embedded:
                issues.append(embedded)

    if not issues:
        print("[OK] no issues found.")
        return 0

    print(f"[WARN] issues found: {min(len(issues), max_issues)}/{len(issues)}")
    for item in issues[:max_issues]:
        t = item["type"]
        if t == "RUN_SPLIT":
            print(
                f"RUN_SPLIT {item['part']}#p{item['p_index']} runs {item['run_start']}-{item['run_end']}: "
                f"{item['placeholder']} | {item['text']}"
            )
        elif t == "BODY_STYLE_LOCATION":
            print(
                f"BODY_STYLE_LOCATION {item['part']}#p{item['p_index']} style={item['style']} in {item['location']}: "
                f"{item['text']}"
            )
        elif t == "EXTERNAL_RELS":
            print(f"EXTERNAL_RELS {item['part']}: {item['target']}")
        elif t == "FIELD_EXTERNAL":
            print(f"FIELD_EXTERNAL {item['part']}: {item['text']}")
        elif t == "EMBEDDED_OBJECT":
            print(f"EMBEDDED_OBJECT {item['part']}")

    return 2


def main() -> int:
    parser = argparse.ArgumentParser(description="Docx template/output self-check.")
    parser.add_argument("docx", help="Path to .docx (template or generated output)")
    parser.add_argument(
        "--mode",
        choices=["template", "output", "all"],
        default="template",
        help="template: focus on run-split placeholders; output: focus on external links/embeddings; all: everything",
    )
    parser.add_argument(
        "--body-styles",
        default=",".join(sorted(DEFAULT_BODY_STYLE_NAMES)),
        help="Comma-separated paragraph style display names to treat as body styles (default: 正文,Normal)",
    )
    parser.add_argument("--max", type=int, default=200, help="Max issues to print (default 200)")
    args = parser.parse_args()

    body_style_names = {s.strip() for s in args.body_styles.split(",") if s.strip()}

    docx = Path(args.docx)
    if not docx.exists():
        print(f"[ERROR] file not found: {docx}")
        return 1

    code = run_checks(docx, args.mode, body_style_names, args.max)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
