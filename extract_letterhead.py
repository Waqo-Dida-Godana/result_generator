#!/usr/bin/env python3
"""
Extract letterhead assets from a DOCX template without external DOCX parsers.
Run: python extract_letterhead.py
"""

import json
import os
import posixpath
import xml.etree.ElementTree as ET
import zipfile


WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _read_xml_from_zip(docx_zip, member_name):
    with docx_zip.open(member_name) as handle:
        return ET.parse(handle).getroot()


def _get_targets(docx_zip, relationship_suffix):
    rels_root = _read_xml_from_zip(docx_zip, "word/_rels/document.xml.rels")
    targets = []
    for rel in rels_root.findall("rel:Relationship", WORD_NS):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith(relationship_suffix):
            target = rel.attrib.get("Target", "")
            if not target:
                continue
            targets.append(posixpath.normpath(posixpath.join("word", target)))
    return targets


def _extract_text(docx_zip, xml_members):
    lines = []
    for xml_member in xml_members:
        xml_root = _read_xml_from_zip(docx_zip, xml_member)
        for paragraph in xml_root.findall(".//w:p", WORD_NS):
            text = "".join(node.text or "" for node in paragraph.findall(".//w:t", WORD_NS)).strip()
            if text:
                lines.append(text)
    return lines


def _extract_first_image(docx_zip, xml_members):
    for xml_member in xml_members:
        rels_member = posixpath.join(
            posixpath.dirname(xml_member),
            "_rels",
            f"{posixpath.basename(xml_member)}.rels",
        )
        if rels_member not in docx_zip.namelist():
            continue

        rels_root = _read_xml_from_zip(docx_zip, rels_member)
        for rel in rels_root.findall("rel:Relationship", WORD_NS):
            rel_type = rel.attrib.get("Type", "")
            if not rel_type.endswith("/image"):
                continue

            target = rel.attrib.get("Target", "")
            if not target:
                continue

            image_member = posixpath.normpath(posixpath.join(posixpath.dirname(xml_member), target))
            with docx_zip.open(image_member) as handle:
                return handle.read()
    return None


def _save_image(image_bytes, output_path):
    if image_bytes is None:
        return False, None

    with open(output_path, "wb") as handle:
        handle.write(image_bytes)
    return True, len(image_bytes)


def extract_letterhead(docx_path="assets/letterhead.docx", output_dir="assets"):
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        header_members = _get_targets(docx_zip, "/header")
        footer_members = _get_targets(docx_zip, "/footer")
        header_text = _extract_text(docx_zip, header_members)
        footer_text = _extract_text(docx_zip, footer_members)
        header_image = _extract_first_image(docx_zip, header_members)
        footer_image = _extract_first_image(docx_zip, footer_members)

    header_png_path = os.path.join(output_dir, "letterhead.png")
    footer_png_path = os.path.join(output_dir, "letterhead_footer.png")
    json_path = os.path.join(output_dir, "letterhead.json")

    header_image_found, header_image_bytes = _save_image(header_image, header_png_path)
    footer_image_found, footer_image_bytes = _save_image(footer_image, footer_png_path)

    metadata = {
        "docx_path": os.path.abspath(docx_path),
        "header_lines": header_text[:10],
        "footer_lines": footer_text[:10],
        "image_found": header_image_found,
        "image_bytes": header_image_bytes,
        "footer_image_found": footer_image_found,
        "footer_image_bytes": footer_image_bytes,
        "header_count": len(header_members),
        "footer_count": len(footer_members),
    }
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(metadata, handle, indent=2)

    if header_image_found:
        print(f"Header PNG saved: {header_png_path} ({header_image_bytes} bytes)")
    else:
        print("No header image found in DOCX")
    if footer_image_found:
        print(f"Footer PNG saved: {footer_png_path} ({footer_image_bytes} bytes)")
    else:
        print("No footer image found in DOCX")
    print(f"Metadata saved: {json_path}")
    if header_text:
        print("Header text preview:", " | ".join(header_text[:3]))
    if footer_text:
        print("Footer text preview:", " | ".join(footer_text[:3]))


if __name__ == "__main__":
    extract_letterhead()
