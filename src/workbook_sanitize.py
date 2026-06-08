from __future__ import annotations

import re
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile


PIVOT_PART_RE = re.compile(r"(^|/)(pivotCache|pivotTables?)/|pivotCache|pivotTable", re.IGNORECASE)
PIVOT_REL_RE = re.compile(r"pivot(CacheDefinition|Table)", re.IGNORECASE)


def sanitize_workbook_for_openpyxl(source: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / source.name
    with ZipFile(source, "r") as zin, ZipFile(target, "w", compression=ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            if PIVOT_PART_RE.search(name):
                continue
            data = zin.read(item)
            if name.endswith(".rels"):
                data = _remove_pivot_relationships(data)
            elif name.endswith(".xml"):
                data = _remove_pivot_xml_nodes(data)
            zout.writestr(item, data)
    return target


def _remove_pivot_relationships(data: bytes) -> bytes:
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return data
    removed = False
    for child in list(root):
        rel_type = child.attrib.get("Type", "")
        target = child.attrib.get("Target", "")
        if PIVOT_REL_RE.search(rel_type) or PIVOT_REL_RE.search(target):
            root.remove(child)
            removed = True
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True) if removed else data


def _remove_pivot_xml_nodes(data: bytes) -> bytes:
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return data
    removed = _remove_matching_nodes(root)
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True) if removed else data


def _remove_matching_nodes(parent) -> bool:
    removed = False
    for child in list(parent):
        tag_name = child.tag.rsplit("}", 1)[-1]
        if tag_name.startswith("pivot"):
            parent.remove(child)
            removed = True
        else:
            removed = _remove_matching_nodes(child) or removed
    return removed
