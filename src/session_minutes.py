#!/usr/bin/env python3
"""Safely preview, pilot, and apply Arabic session-minute OOXML insertions."""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
import os
import posixpath
import re
import shutil
import sys
import tempfile
import unicodedata
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree


SCHEMA_VERSION = 6
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W_NS, "w14": W14_NS, "wp": WP_NS}

DOCUMENT_XML = "word/document.xml"
DOCUMENT_RELS = "word/_rels/document.xml.rels"
STYLES_XML = "word/styles.xml"
CONTENT_TYPES = "[Content_Types].xml"
GLOSSARY_DOCUMENT = "word/glossary/document.xml"
GLOSSARY_STYLES = "word/glossary/styles.xml"
GLOSSARY_PARTS = (
    GLOSSARY_DOCUMENT,
    "word/glossary/settings.xml",
    "word/glossary/_rels/document.xml.rels",
    GLOSSARY_STYLES,
    "word/glossary/webSettings.xml",
    "word/glossary/fontTable.xml",
)
GLOSSARY_CONTENT_TYPES = {
    "/word/glossary/document.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
    ),
    "/word/glossary/styles.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    ),
    "/word/glossary/settings.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    ),
    "/word/glossary/webSettings.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
    ),
    "/word/glossary/fontTable.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    ),
}
GLOSSARY_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"
)

SESSION_PREFIX = "وبإحالة الدعوى"
SESSION_ALTERNATE = "واستكمالاً"
NOTICE_PREFIX = "صدر هذا القرار"
HEADING_BASIS = "الأسانيد"
HEADING_ORDER = "منطوق القرار"
SIGNATURE_PREFIXES = (
    "أمين سر اللجنة",
    "عضو اللجنة",
    "رئيس اللجنة",
    "الأستاذ/",
    "الدكتور/",
)
CONTROL_TYPES = (
    "comboBox",
    "dropDownList",
    "date",
    "checkBox",
    "picture",
    "richText",
    "text",
    "docPartObj",
    "group",
    "repeatingSection",
    "repeatingSectionItem",
)
PILOT_COUNT = 6
MAX_DOCX_ARCHIVE_BYTES = 256 * 1024 * 1024
MAX_ZIP_ENTRIES = 4096
MAX_ZIP_PART_BYTES = 128 * 1024 * 1024
MAX_ZIP_TOTAL_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_XML_BYTES = 64 * 1024 * 1024
ZIP_ENCRYPTED_FLAG = 0x1
MAX_SDT_ID = 0x7FFFFFFE
MAX_PARAGRAPH_ID = 0xFFFFFFFE
POLICY = {
    "template_mapping": "document_order",
    "session_rule": "insert_unless_current_functional_block_exists",
    "existing_sessions": "preserve_all",
    "closing_rule": "keep_current_or_collapse_all_to_current",
    "adjournment_rule": "process_fully",
    "glossary_rule": "merge_or_create",
    "source_rule": "never_modify",
}

ARABIC_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"
)
TIME_RE = re.compile(r"(?<!\d)([0-9]{1,2})\s*:\s*([0-9]{2})(?!\d)")


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


W_VAL = qn(W_NS, "val")
W_ID = qn(W_NS, "id")
W_STYLE_ID = qn(W_NS, "styleId")
W_SDT_PR = qn(W_NS, "sdtPr")
W_GUID = qn(W_NS, "guid")
W14_PARA_ID = qn(W14_NS, "paraId")
W14_TEXT_ID = qn(W14_NS, "textId")


class WorkflowError(RuntimeError):
    """An expected, user-actionable workflow failure."""


@dataclass(frozen=True)
class OpeningTemplate:
    order: int
    minutes: int
    display_time: str
    text: str
    normalized_text: str
    xml_hash: str
    control_signature: tuple[Any, ...]
    control_counts: dict[str, int]
    node: etree._Element


@dataclass(frozen=True)
class TemplateData:
    path: Path
    sha256: str
    openings: tuple[OpeningTemplate, ...]
    closing_text: str
    closing_normalized_text: str
    closing_xml_hash: str
    closing_control_signature: tuple[Any, ...]
    closing_control_counts: dict[str, int]
    closing_node: etree._Element
    closing_table_node: etree._Element
    closing_table_xml_hash: str
    closing_table_visual_hash: str
    closing_styles: dict[str, etree._Element]
    closing_relationships: dict[str, tuple[str, str | None, str | None, bytes | None]]
    closing_content_types: dict[str, str]
    glossary_docparts: dict[str, etree._Element]
    glossary_styles: dict[str, etree._Element]
    glossary_parts: dict[str, bytes]


@dataclass
class DocxParts:
    document_root: etree._Element
    glossary_root: etree._Element | None
    glossary_styles_root: etree._Element | None
    protected: bool
    names: set[str]


class SdtIdAllocator:
    def __init__(self, document_root: etree._Element):
        values: set[int] = set()
        for element in document_root.xpath(".//w:sdtPr/w:id", namespaces=NS):
            raw = element.get(W_VAL)
            try:
                value = int(raw)
            except (TypeError, ValueError):
                continue
            if 0 < value <= MAX_SDT_ID:
                values.add(value)
        self._used = values
        highest = max(values, default=1000)
        self._next = highest + 1 if highest < MAX_SDT_ID else 1

    def _allocate(self) -> int:
        for _ in range(len(self._used) + 1):
            value = self._next
            self._next = 1 if value >= MAX_SDT_ID else value + 1
            if value not in self._used:
                self._used.add(value)
                return value
        raise WorkflowError("no safe Word content-control identifiers remain")

    def apply(self, node: etree._Element) -> None:
        for element in node.xpath(".//w:sdtPr/w:id", namespaces=NS):
            element.set(W_VAL, str(self._allocate()))


class ParagraphIdAllocator:
    def __init__(self, document_root: etree._Element):
        values: set[int] = set()
        for attribute in (W14_PARA_ID, W14_TEXT_ID):
            for element in document_root.xpath(f".//*[@w14:{etree.QName(attribute).localname}]", namespaces=NS):
                raw = element.get(attribute)
                try:
                    value = int(raw, 16)
                except (TypeError, ValueError):
                    continue
                if 0 < value <= MAX_PARAGRAPH_ID:
                    values.add(value)
        self._used = values
        highest = max(values, default=0x10000000)
        self._next = highest + 1 if highest < MAX_PARAGRAPH_ID else 1

    def _allocate(self) -> str:
        for _ in range(len(self._used) + 1):
            value = self._next
            self._next = 1 if value >= MAX_PARAGRAPH_ID else value + 1
            if value not in self._used:
                self._used.add(value)
                return f"{value:08X}"
        raise WorkflowError("no safe Word paragraph identifiers remain")

    def apply(self, node: etree._Element) -> None:
        paragraphs: list[etree._Element] = []
        if etree.QName(node).localname == "p":
            paragraphs.append(node)
        paragraphs.extend(node.xpath(".//w:p", namespaces=NS))
        for paragraph in paragraphs:
            paragraph.set(W14_PARA_ID, self._allocate())
            paragraph.set(W14_TEXT_ID, self._allocate())


def parse_xml(data: bytes, *, part_name: str = "XML part") -> etree._Element:
    if len(data) > MAX_XML_BYTES:
        raise WorkflowError(
            f"{part_name} exceeds the safe XML size limit ({MAX_XML_BYTES} bytes)"
        )
    if b"<!DOCTYPE" in data.upper():
        raise WorkflowError(f"{part_name} contains a forbidden document type declaration")
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        dtd_validation=False,
        remove_blank_text=False,
        huge_tree=False,
        recover=False,
    )
    try:
        return etree.fromstring(data, parser=parser)
    except etree.XMLSyntaxError as exc:
        raise WorkflowError(f"invalid {part_name}: {exc}") from exc


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(
        root, encoding="UTF-8", xml_declaration=True, standalone=True
    )


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except PermissionError as exc:
        raise WorkflowError(
            f"تعذر قراءة الملف لأنه مفتوح حصريًا. احفظه وأغلقه ثم أعد المحاولة: {path}"
        ) from exc
    return digest.hexdigest()


def normalize_text(value: str) -> str:
    value = value.replace("ـ", "")
    value = "".join(
        character
        for character in unicodedata.normalize("NFD", value)
        if unicodedata.category(character) != "Mn"
    )
    return "".join(character.lower() for character in value if character.isalnum())


SESSION_NORM = normalize_text(SESSION_PREFIX)
SESSION_ALT_NORM = normalize_text(SESSION_ALTERNATE)
NOTICE_NORM = normalize_text(NOTICE_PREFIX)
HEADING_BASIS_NORM = normalize_text(HEADING_BASIS)
HEADING_ORDER_NORM = normalize_text(HEADING_ORDER)
SIGNATURE_NORMS = tuple(normalize_text(item) for item in SIGNATURE_PREFIXES)


def paragraph_text(paragraph: etree._Element) -> str:
    text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
    return re.sub(r"\s+", " ", text).strip()


def body_paragraphs(
    root: etree._Element, *, nonempty: bool = False
) -> list[etree._Element]:
    paragraphs = list(root.xpath(".//w:body//w:p", namespaces=NS))
    if nonempty:
        return [paragraph for paragraph in paragraphs if paragraph_text(paragraph)]
    return paragraphs


def parse_time(value: str) -> tuple[int, str]:
    translated = value.translate(ARABIC_DIGITS)
    matches = list(TIME_RE.finditer(translated))
    if len(matches) != 1:
        raise WorkflowError(
            f"Expected exactly one time in template paragraph, found {len(matches)}"
        )
    hour = int(matches[0].group(1))
    minute = int(matches[0].group(2))
    if hour < 1 or hour > 12 or minute > 59:
        raise WorkflowError(f"Invalid 12-hour time: {matches[0].group(0)}")
    normalized = normalize_text(value)
    if normalize_text("مساء") in normalized:
        hour24 = hour if hour == 12 else hour + 12
        suffix = "مساءً"
    elif normalize_text("صباح") in normalized:
        hour24 = 0 if hour == 12 else hour
        suffix = "صباحًا"
    else:
        hour24 = hour
        suffix = ""
    display = matches[0].group(0) + (f" {suffix}" if suffix else "")
    return hour24 * 60 + minute, display


def remove_volatile_identity(node: etree._Element) -> None:
    for element in node.iter():
        for attribute in list(element.attrib):
            local = etree.QName(attribute).localname
            if local in {"paraId", "textId"} or local.startswith("rsid"):
                del element.attrib[attribute]


def semantic_xml_hash(node: etree._Element) -> str:
    clone = copy.deepcopy(node)
    remove_volatile_identity(clone)

    def signature(element: etree._Element) -> list[Any]:
        element_name = etree.QName(element.tag)
        attributes: list[tuple[str, str, str]] = []
        for raw_name, raw_value in element.attrib.items():
            attribute_name = etree.QName(raw_name)
            value = raw_value
            parent = element.getparent()
            if (
                element.tag == W_ID
                and parent is not None
                and parent.tag == W_SDT_PR
                and attribute_name.localname == "val"
            ):
                value = "0"
            if (
                etree.QName(element).localname == "docPart"
                and parent is not None
                and etree.QName(parent).localname == "placeholder"
                and attribute_name.localname == "val"
            ):
                value = "PLACEHOLDER"
            attributes.append(
                (attribute_name.namespace or "", attribute_name.localname, value)
            )
        attributes.sort()
        return [
            element_name.namespace or "",
            element_name.localname,
            attributes,
            element.text or "",
            [signature(child) for child in element],
        ]

    payload = json.dumps(
        signature(clone),
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    return sha256_bytes(payload)


def closing_semantic_xml_hash(node: etree._Element) -> str:
    """Hash closing content while ignoring pagination-only keep settings."""

    clone = copy.deepcopy(node)
    for property_name in ("keepNext", "keepLines"):
        for element in clone.findall(f".//w:pPr/w:{property_name}", NS):
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
    return semantic_xml_hash(clone)


def signature_table_layout_hash(node: etree._Element) -> str:
    """Hash signature-table content and layout while ignoring remapped IDs."""

    clone = copy.deepcopy(node)
    for element in clone.xpath(".//wp:docPr", namespaces=NS):
        element.set("id", "0")
    for element in clone.iter():
        for attribute in list(element.attrib):
            if etree.QName(attribute).namespace == R_NS:
                element.set(attribute, "REL")
    return semantic_xml_hash(clone)


STYLE_REFERENCE_NAMES = {"tblStyle", "pStyle", "rStyle"}
STYLE_DEPENDENCY_NAMES = {"basedOn", "next", "link"}


def style_reference_ids(node: etree._Element) -> list[str]:
    """Return main-document style IDs referenced by copied table content."""

    values: list[str] = []
    for element in node.iter():
        if etree.QName(element).localname not in STYLE_REFERENCE_NAMES:
            continue
        value = element.get(W_VAL)
        if value and value not in values:
            values.append(value)
    return values


def style_map(root: etree._Element) -> dict[str, etree._Element]:
    return {
        style_id: element
        for element in root.xpath("./w:style", namespaces=NS)
        if (style_id := element.get(W_STYLE_ID))
    }


def collect_style_closure(
    root: etree._Element,
    initial_style_ids: Sequence[str],
    *,
    context: str,
) -> dict[str, etree._Element]:
    """Resolve referenced styles plus basedOn/next/link dependencies."""

    available = style_map(root)
    collected: dict[str, etree._Element] = {}
    pending = list(initial_style_ids)
    while pending:
        style_id = pending.pop(0)
        if not style_id or style_id in collected:
            continue
        style = available.get(style_id)
        if style is None:
            raise WorkflowError(f"{context} is missing Word style: {style_id}")
        collected[style_id] = style
        for child in style:
            if etree.QName(child).localname not in STYLE_DEPENDENCY_NAMES:
                continue
            dependency = child.get(W_VAL)
            if dependency and dependency not in collected:
                pending.append(dependency)
    return collected


def style_definition_visual_hash(style: etree._Element) -> str:
    """Hash a style definition independently of conflict-safe remapped IDs."""

    clone = copy.deepcopy(style)
    clone.set(W_STYLE_ID, "STYLE")
    clone.attrib.pop(qn(W_NS, "default"), None)
    for child in clone:
        local_name = etree.QName(child).localname
        if local_name in STYLE_DEPENDENCY_NAMES:
            child.set(W_VAL, local_name.upper())
    return semantic_xml_hash(clone)


def signature_style_hashes(
    styles_root: etree._Element, table: etree._Element, *, context: str
) -> Counter[str]:
    closure = collect_style_closure(
        styles_root, style_reference_ids(table), context=context
    )
    return Counter(style_definition_visual_hash(style) for style in closure.values())


def stabilize_signature_table_geometry(table: etree._Element) -> None:
    """Keep the template's physical width when target page geometry differs.

    The template table stores both an absolute grid and a 104.6% preferred
    width. Copying the percentage into an A4 document with narrower margins
    makes the floating table wider than its original Letter-page footprint and
    clips the leftmost name. The grid is the authoritative template geometry,
    so convert preferred/cell widths to fixed twips and center the float.
    """

    grid_widths: list[int] = []
    for column in table.xpath("./w:tblGrid/w:gridCol", namespaces=NS):
        try:
            width = int(column.get(qn(W_NS, "w"), "0"))
        except ValueError:
            width = 0
        if width <= 0:
            raise WorkflowError("template signature table has an invalid grid width")
        grid_widths.append(width)
    if not grid_widths:
        raise WorkflowError("template signature table has no fixed grid")

    properties = table.find("w:tblPr", NS)
    if properties is None:
        raise WorkflowError("template signature table has no table properties")
    preferred_width = properties.find("w:tblW", NS)
    if preferred_width is None:
        preferred_width = etree.Element(qn(W_NS, "tblW"))
        properties.insert(0, preferred_width)
    preferred_width.set(qn(W_NS, "type"), "dxa")
    preferred_width.set(qn(W_NS, "w"), str(sum(grid_widths)))

    # The model positions the signature table at an absolute Y coordinate on
    # the page. Carrying that anchor into documents with different pagination
    # puts the signatures near the bottom of the following page and creates a
    # large blank gap. Make the copied table inline so it follows the closing
    # paragraph naturally, then center the fixed-width grid.
    for floating_name in ("tblpPr", "tblOverlap"):
        floating = properties.find(f"w:{floating_name}", NS)
        if floating is not None:
            properties.remove(floating)
    justification = properties.find("w:jc", NS)
    if justification is None:
        justification = etree.Element(qn(W_NS, "jc"))
        properties.insert(properties.index(preferred_width) + 1, justification)
    justification.set(W_VAL, "center")

    for row in table.xpath("./w:tr", namespaces=NS):
        grid_index = 0
        for cell in row.xpath("./w:tc", namespaces=NS):
            cell_properties = cell.find("w:tcPr", NS)
            if cell_properties is None:
                cell_properties = etree.Element(qn(W_NS, "tcPr"))
                cell.insert(0, cell_properties)
            span_element = cell_properties.find("w:gridSpan", NS)
            try:
                span = int(span_element.get(W_VAL, "1")) if span_element is not None else 1
            except ValueError:
                span = 1
            span = max(1, span)
            end = min(len(grid_widths), grid_index + span)
            if grid_index >= len(grid_widths) or end <= grid_index:
                raise WorkflowError("template signature table cell exceeds its grid")
            cell_width = cell_properties.find("w:tcW", NS)
            if cell_width is None:
                cell_width = etree.Element(qn(W_NS, "tcW"))
                cell_properties.insert(0, cell_width)
            cell_width.set(qn(W_NS, "type"), "dxa")
            cell_width.set(qn(W_NS, "w"), str(sum(grid_widths[grid_index:end])))
            grid_index = end


def signature_table_visual_hash(node: etree._Element) -> str:
    """Hash current visual geometry independently of remapped style IDs."""

    clone = copy.deepcopy(node)
    for element in clone.iter():
        if etree.QName(element).localname in STYLE_REFERENCE_NAMES:
            element.set(W_VAL, "STYLE")
    return signature_table_layout_hash(clone)


def adapted_signature_table_visual_hash(node: etree._Element) -> str:
    """Hash the template table after deterministic cross-page adaptation."""

    clone = copy.deepcopy(node)
    stabilize_signature_table_geometry(clone)
    return signature_table_visual_hash(clone)


def keep_closing_with_following_table(paragraph: etree._Element) -> None:
    """Keep the complete closing paragraph on the same page as its table."""

    paragraph_properties = paragraph.find("w:pPr", NS)
    if paragraph_properties is None:
        paragraph_properties = etree.Element(qn(W_NS, "pPr"))
        paragraph.insert(0, paragraph_properties)
    # pPr is sequence-ordered: keepNext/keepLines belong immediately after
    # pStyle and before bidi/jc/rPr. Reinsert them in schema order so Word does
    # not repair or ignore the pagination properties on open.
    for property_name in ("keepNext", "keepLines"):
        existing = paragraph_properties.find(f"w:{property_name}", NS)
        if existing is not None:
            paragraph_properties.remove(existing)
    insert_index = 1 if (
        len(paragraph_properties)
        and paragraph_properties[0].tag == qn(W_NS, "pStyle")
    ) else 0
    for property_name in ("keepNext", "keepLines"):
        paragraph_properties.insert(
            insert_index, etree.Element(qn(W_NS, property_name))
        )
        insert_index += 1


def control_type(sdt: etree._Element) -> str:
    properties = sdt.find("w:sdtPr", namespaces=NS)
    if properties is None:
        return "unknown"
    for name in CONTROL_TYPES:
        if properties.find(f"w:{name}", namespaces=NS) is not None:
            return name
    # A Word rich-text content control has no dedicated type child in
    # w:sdtPr; the absence of another type is the rich-text representation.
    return "richText"


def control_definition_signature(sdt: etree._Element) -> tuple[Any, ...]:
    properties = sdt.find("w:sdtPr", namespaces=NS)
    kind = control_type(sdt)
    items: list[tuple[str, str]] = []
    if properties is not None and kind in {"comboBox", "dropDownList"}:
        container = properties.find(f"w:{kind}", namespaces=NS)
        if container is not None:
            for item in container.findall("w:listItem", namespaces=NS):
                items.append(
                    (
                        item.get(qn(W_NS, "displayText"), ""),
                        item.get(qn(W_NS, "value"), ""),
                    )
                )
    return kind, tuple(items)


def paragraph_control_signature(paragraph: etree._Element) -> tuple[Any, ...]:
    return tuple(
        control_definition_signature(sdt)
        for sdt in paragraph.xpath(".//w:sdt", namespaces=NS)
    )


def control_counter(node: etree._Element) -> Counter[str]:
    return Counter(
        control_type(sdt) for sdt in node.xpath(".//w:sdt", namespaces=NS)
    )


def contains_subsequence(
    sequence: Sequence[Any], needle: Sequence[Any]
) -> bool:
    if not needle:
        return True
    if len(needle) > len(sequence):
        return False
    return any(
        tuple(sequence[index : index + len(needle)]) == tuple(needle)
        for index in range(len(sequence) - len(needle) + 1)
    )


def paragraph_is_safe_template(node: etree._Element) -> tuple[bool, str]:
    unsupported = node.xpath(
        ".//w:drawing | .//w:pict | .//w:object | .//w:numPr | .//w:pStyle | .//w:rStyle",
        namespaces=NS,
    )
    if unsupported:
        return (
            False,
            "template paragraph contains unsupported drawings, numbering, or style references",
        )
    for element in node.iter():
        for attribute in element.attrib:
            if etree.QName(attribute).namespace == R_NS:
                return False, "template paragraph contains an external relationship"
    return True, ""


def package_is_protected(package: ZipFile) -> bool:
    if "word/settings.xml" not in package.namelist():
        return False
    settings = parse_xml(package.read("word/settings.xml"))
    for protection in settings.xpath(".//w:documentProtection", namespaces=NS):
        value = (protection.get(qn(W_NS, "enforcement")) or "").lower()
        if value in {"1", "true", "on", "yes"}:
            return True
    return False


def read_docx_parts(path: Path) -> DocxParts:
    try:
        with ZipFile(path) as package:
            names = set(package.namelist())
            if DOCUMENT_XML not in names:
                raise WorkflowError("missing word/document.xml")
            root = parse_xml(package.read(DOCUMENT_XML))
            glossary_root = (
                parse_xml(package.read(GLOSSARY_DOCUMENT))
                if GLOSSARY_DOCUMENT in names
                else None
            )
            glossary_styles_root = (
                parse_xml(package.read(GLOSSARY_STYLES))
                if GLOSSARY_STYLES in names
                else None
            )
            return DocxParts(
                document_root=root,
                glossary_root=glossary_root,
                glossary_styles_root=glossary_styles_root,
                protected=package_is_protected(package),
                names=names,
            )
    except PermissionError as exc:
        raise WorkflowError(
            f"تعذر قراءة الملف لأنه مفتوح حصريًا. احفظه وأغلقه ثم أعد المحاولة: {path}"
        ) from exc
    except BadZipFile as exc:
        raise WorkflowError("invalid DOCX ZIP package") from exc
    except (OSError, etree.XMLSyntaxError) as exc:
        raise WorkflowError(f"cannot read DOCX: {exc}") from exc


def glossary_maps(
    glossary_root: etree._Element, styles_root: etree._Element
) -> tuple[dict[str, etree._Element], dict[str, etree._Element]]:
    docparts: dict[str, etree._Element] = {}
    for docpart in glossary_root.xpath(".//w:docPart", namespaces=NS):
        name = docpart.find("w:docPartPr/w:name", namespaces=NS)
        if name is not None and name.get(W_VAL):
            docparts[name.get(W_VAL)] = docpart
    styles: dict[str, etree._Element] = {}
    for style in styles_root.xpath("./w:style", namespaces=NS):
        style_id = style.get(W_STYLE_ID)
        if style_id:
            styles[style_id] = style
    return docparts, styles


def placeholder_names(node: etree._Element) -> list[str]:
    return [
        element.get(W_VAL)
        for element in node.xpath(
            ".//w:sdtPr/w:placeholder/w:docPart", namespaces=NS
        )
        if element.get(W_VAL)
    ]


def load_template(path: Path) -> TemplateData:
    parts = read_docx_parts(path)
    if parts.protected:
        raise WorkflowError("template is protected")
    if parts.glossary_root is None or parts.glossary_styles_root is None:
        raise WorkflowError("template is missing its Word glossary parts")
    openings: list[OpeningTemplate] = []
    closing_candidates: list[etree._Element] = []
    for paragraph in body_paragraphs(parts.document_root, nonempty=True):
        text = paragraph_text(paragraph)
        normalized = normalize_text(text)
        if normalized.startswith(SESSION_NORM):
            minutes, display_time = parse_time(text)
            safe, reason = paragraph_is_safe_template(paragraph)
            if not safe:
                raise WorkflowError(reason)
            signature = paragraph_control_signature(paragraph)
            counts = control_counter(paragraph)
            if len(signature) != 7 or counts != Counter(
                {"comboBox": 6, "dropDownList": 1}
            ):
                raise WorkflowError(
                    "each template session paragraph must contain 6 ComboBox and 1 DropDownList controls"
                )
            openings.append(
                OpeningTemplate(
                    order=len(openings) + 1,
                    minutes=minutes,
                    display_time=display_time,
                    text=text,
                    normalized_text=normalized,
                    xml_hash=semantic_xml_hash(paragraph),
                    control_signature=signature,
                    control_counts=dict(counts),
                    node=paragraph,
                )
            )
        if normalized.startswith(NOTICE_NORM):
            closing_candidates.append(paragraph)
    if not openings:
        raise WorkflowError(
            "expected at least one ordered session paragraph in the template"
        )
    if len(closing_candidates) != 1:
        raise WorkflowError(
            f"expected one decision-notice paragraph in template, found {len(closing_candidates)}"
        )
    closing = closing_candidates[0]
    safe, reason = paragraph_is_safe_template(closing)
    if not safe:
        raise WorkflowError(reason)
    closing_signature = paragraph_control_signature(closing)
    closing_counts = control_counter(closing)
    if len(closing_signature) != 2 or closing_counts != Counter({"comboBox": 2}):
        raise WorkflowError(
            "template closing paragraph must contain exactly two ComboBox controls"
        )

    body = closing.getparent()
    if body is None:
        raise WorkflowError("template closing paragraph is detached")
    closing_position = body.index(closing)
    closing_tables = [
        node
        for node in body[:closing_position]
        if etree.QName(node).localname == "tbl"
        and node.xpath(".//w:drawing | .//w:pict", namespaces=NS)
        and any(
            normalize_text(prefix) in normalize_text(paragraph_text(node))
            for prefix in SIGNATURE_PREFIXES
        )
    ]
    if len(closing_tables) != 1:
        raise WorkflowError(
            "template must contain one signature table with images before the closing paragraph"
        )
    closing_table = closing_tables[0]

    docparts, styles = glossary_maps(
        parts.glossary_root, parts.glossary_styles_root
    )
    referenced = set()
    for paragraph in [opening.node for opening in openings] + [closing]:
        referenced.update(placeholder_names(paragraph))
    for name in sorted(referenced):
        if name not in docparts:
            raise WorkflowError(
                f"template placeholder does not resolve to a glossary docPart: {name}"
            )
        style = docparts[name].find(".//w:pStyle", namespaces=NS)
        if style is not None:
            style_id = style.get(W_VAL)
            if style_id not in styles:
                raise WorkflowError(
                    f"template glossary style is missing for placeholder: {name}"
                )
    closing_relationships: dict[
        str, tuple[str, str | None, str | None, bytes | None]
    ] = {}
    closing_content_types: dict[str, str] = {}
    closing_styles: dict[str, etree._Element] = {}
    try:
        with ZipFile(path) as package:
            glossary_parts = {
                name: package.read(name) for name in GLOSSARY_PARTS
            }
            relationship_root = parse_xml(package.read(DOCUMENT_RELS))
            relationship_map = {
                relationship.get("Id"): relationship
                for relationship in relationship_root
                if relationship.get("Id")
            }
            referenced_relationships = {
                value
                for element in closing_table.iter()
                for attribute, value in element.attrib.items()
                if etree.QName(attribute).namespace == R_NS
            }
            for relationship_id in sorted(referenced_relationships):
                relationship = relationship_map.get(relationship_id)
                if relationship is None:
                    raise WorkflowError(
                        f"signature table relationship is missing: {relationship_id}"
                    )
                target = relationship.get("Target")
                target_mode = relationship.get("TargetMode")
                part_name: str | None = None
                payload: bytes | None = None
                if target and target_mode != "External":
                    part_name = str((Path("word") / target).as_posix())
                    payload = package.read(part_name)
                closing_relationships[relationship_id] = (
                    relationship.get("Type", ""),
                    target_mode,
                    part_name,
                    payload,
                )
            content_types_root = parse_xml(package.read(CONTENT_TYPES))
            for default in content_types_root.findall(qn(CT_NS, "Default")):
                extension = (default.get("Extension") or "").lower()
                content_type = default.get("ContentType")
                if extension and content_type:
                    closing_content_types[extension] = content_type
            main_styles_root = parse_xml(
                package.read(STYLES_XML), part_name=STYLES_XML
            )
            closing_styles = collect_style_closure(
                main_styles_root,
                style_reference_ids(closing_table),
                context="template signature table",
            )
    except KeyError as exc:
        raise WorkflowError(f"template glossary package is incomplete: {exc}") from exc

    return TemplateData(
        path=path,
        sha256=sha256_file(path),
        openings=tuple(openings),
        closing_text=paragraph_text(closing),
        closing_normalized_text=normalize_text(paragraph_text(closing)),
        closing_xml_hash=closing_semantic_xml_hash(closing),
        closing_control_signature=closing_signature,
        closing_control_counts=dict(closing_counts),
        closing_node=closing,
        closing_table_node=closing_table,
        closing_table_xml_hash=signature_table_layout_hash(closing_table),
        closing_table_visual_hash=adapted_signature_table_visual_hash(closing_table),
        closing_styles=closing_styles,
        closing_relationships=closing_relationships,
        closing_content_types=closing_content_types,
        glossary_docparts=docparts,
        glossary_styles=styles,
        glossary_parts=glossary_parts,
    )


class GlossaryManager:
    def __init__(
        self,
        glossary_root: etree._Element,
        styles_root: etree._Element,
        template: TemplateData,
        *,
        created: bool,
    ):
        self.root = glossary_root
        self.styles_root = styles_root
        self.template = template
        self.created = created
        self.docparts, self.styles = glossary_maps(self.root, self.styles_root)
        self.mapping: dict[str, str] = {}
        self.added_docparts = 0
        self.added_styles = 0
        containers = self.root.xpath("./w:docParts", namespaces=NS)
        if len(containers) != 1:
            raise WorkflowError("target glossary has no unique w:docParts container")
        self.docparts_container = containers[0]
        self._ensure_core_styles()

    def _append_style(self, style: etree._Element) -> None:
        clone = copy.deepcopy(style)
        remove_volatile_identity(clone)
        self.styles_root.append(clone)
        style_id = clone.get(W_STYLE_ID)
        if style_id:
            self.styles[style_id] = clone
        self.added_styles += 1

    def _ensure_core_styles(self) -> None:
        for style_id in ("Normal", "DefaultParagraphFont", "PlaceholderText"):
            if style_id in self.styles:
                continue
            template_style = self.template.glossary_styles.get(style_id)
            if template_style is None:
                raise WorkflowError(
                    f"template glossary is missing required style {style_id}"
                )
            self._append_style(template_style)

    def _new_name(self, old_name: str) -> str:
        counter = 1
        while True:
            candidate = hashlib.sha256(
                f"{old_name}:{counter}:{len(self.docparts)}".encode("ascii")
            ).hexdigest()[:32].upper()
            if candidate not in self.docparts and candidate not in self.styles:
                return candidate
            counter += 1

    @staticmethod
    def _docpart_visible_text(docpart: etree._Element) -> str:
        return "".join(docpart.xpath(".//w:t/text()", namespaces=NS))

    def ensure_placeholder(self, old_name: str) -> str:
        if old_name in self.mapping:
            return self.mapping[old_name]
        source_docpart = self.template.glossary_docparts.get(old_name)
        source_style = self.template.glossary_styles.get(old_name)
        if source_docpart is None or source_style is None:
            raise WorkflowError(
                f"template glossary definition is incomplete: {old_name}"
            )

        # A placeholder name is also used as a glossary style id.  Reusing an
        # existing name after comparing only its visible text is unsafe: the
        # target may carry different list values, formatting, or docPart
        # metadata.  Any collision is therefore remapped as a complete pair.
        if old_name in self.docparts or old_name in self.styles:
            new_name = self._new_name(old_name)
        else:
            new_name = old_name

        docpart = copy.deepcopy(source_docpart)
        remove_volatile_identity(docpart)
        name_element = docpart.find("w:docPartPr/w:name", namespaces=NS)
        if name_element is None:
            raise WorkflowError("template glossary docPart has no name")
        name_element.set(W_VAL, new_name)
        guid_element = docpart.find("w:docPartPr/w:guid", namespaces=NS)
        if guid_element is not None:
            guid_hex = hashlib.sha256(new_name.encode("ascii")).hexdigest()[:32]
            guid_element.set(W_VAL, "{" + str(uuid.UUID(guid_hex)).upper() + "}")
        for style_reference in docpart.xpath(".//w:pStyle", namespaces=NS):
            if style_reference.get(W_VAL) == old_name:
                style_reference.set(W_VAL, new_name)

        style = copy.deepcopy(source_style)
        remove_volatile_identity(style)
        style.set(W_STYLE_ID, new_name)
        style_name = style.find("w:name", namespaces=NS)
        if style_name is not None:
            style_name.set(W_VAL, new_name)

        self.docparts_container.append(docpart)
        self.styles_root.append(style)
        self.docparts[new_name] = docpart
        self.styles[new_name] = style
        self.mapping[old_name] = new_name
        self.added_docparts += 1
        self.added_styles += 1
        return new_name

    def attach_paragraph(self, paragraph: etree._Element) -> None:
        for reference in paragraph.xpath(
            ".//w:sdtPr/w:placeholder/w:docPart", namespaces=NS
        ):
            old_name = reference.get(W_VAL)
            if not old_name:
                raise WorkflowError("template placeholder reference has no value")
            reference.set(W_VAL, self.ensure_placeholder(old_name))


def clone_template_paragraph(
    node: etree._Element,
    sdt_allocator: SdtIdAllocator,
    paragraph_allocator: ParagraphIdAllocator,
    glossary: GlossaryManager,
) -> etree._Element:
    clone = copy.deepcopy(node)
    remove_volatile_identity(clone)
    sdt_allocator.apply(clone)
    paragraph_allocator.apply(clone)
    glossary.attach_paragraph(clone)
    return clone


def _next_relationship_id(root: etree._Element) -> str:
    used = {relationship.get("Id", "") for relationship in root}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def _ensure_content_type_default(
    root: etree._Element, extension: str, content_type: str
) -> None:
    for default in root.findall(qn(CT_NS, "Default")):
        if (default.get("Extension") or "").lower() == extension.lower():
            return
    etree.SubElement(
        root,
        qn(CT_NS, "Default"),
        Extension=extension,
        ContentType=content_type,
    )


def normalize_signature_table_flow(
    table: etree._Element,
    document_root: etree._Element,
    reserved_drawing_ids: set[int] | None = None,
) -> None:
    """Preserve the template table exactly and only allocate safe drawing IDs."""

    table_drawing_properties = table.xpath(".//wp:docPr", namespaces=NS)
    table_elements = set(table_drawing_properties)
    used_ids: set[int] = set(reserved_drawing_ids or ())
    for element in document_root.xpath(".//wp:docPr", namespaces=NS):
        if element in table_elements:
            continue
        try:
            used_ids.add(int(element.get("id", "")))
        except ValueError:
            continue
    candidate = max(used_ids, default=0) + 1
    for element in table_drawing_properties:
        while candidate in used_ids:
            candidate += 1
        element.set("id", str(candidate))
        used_ids.add(candidate)
        candidate += 1


def import_signature_styles(
    source: Path,
    table: etree._Element,
    template: TemplateData,
    replacements: dict[str, bytes],
    allowed_changes: set[str],
) -> None:
    """Import the table's complete main-document style closure safely."""

    try:
        with ZipFile(source) as package:
            style_bytes = replacements.get(STYLES_XML, package.read(STYLES_XML))
    except KeyError as exc:
        raise WorkflowError("source document is missing word/styles.xml") from exc
    target_root = parse_xml(style_bytes, part_name=STYLES_XML)
    target_styles = style_map(target_root)
    mapping: dict[str, str] = {}
    reserved = set(target_styles)

    for old_id, source_style in template.closing_styles.items():
        existing = target_styles.get(old_id)
        if existing is None or semantic_xml_hash(existing) == semantic_xml_hash(source_style):
            mapping[old_id] = old_id
            reserved.add(old_id)
            continue

        base = "SessionMinutes_" + hashlib.sha256(
            xml_bytes(source_style)
        ).hexdigest()[:16]
        candidate = base
        suffix = 2
        while candidate in reserved:
            candidate_style = target_styles.get(candidate)
            if (
                candidate_style is not None
                and style_definition_visual_hash(candidate_style)
                == style_definition_visual_hash(source_style)
            ):
                break
            candidate = f"{base}_{suffix}"
            suffix += 1
        mapping[old_id] = candidate
        reserved.add(candidate)

    changed = False
    for old_id, source_style in template.closing_styles.items():
        mapped_id = mapping[old_id]
        existing = target_styles.get(mapped_id)
        if existing is not None:
            continue
        clone = copy.deepcopy(source_style)
        clone.set(W_STYLE_ID, mapped_id)
        if mapped_id != old_id:
            # A remapped default style remains a complete visual base for the
            # copied table but must not compete with the document's own default.
            clone.attrib.pop(qn(W_NS, "default"), None)
        for child in clone:
            if etree.QName(child).localname not in STYLE_DEPENDENCY_NAMES:
                continue
            dependency = child.get(W_VAL)
            if dependency in mapping:
                child.set(W_VAL, mapping[dependency])
        target_root.append(clone)
        target_styles[mapped_id] = clone
        changed = True

    for element in table.iter():
        if etree.QName(element).localname not in STYLE_REFERENCE_NAMES:
            continue
        style_id = element.get(W_VAL)
        if style_id in mapping:
            element.set(W_VAL, mapping[style_id])

    if changed:
        replacements[STYLES_XML] = xml_bytes(target_root)
        allowed_changes.add(STYLES_XML)


def clone_template_signature_table(
    source: Path,
    template: TemplateData,
    document_root: etree._Element,
    sdt_allocator: SdtIdAllocator,
    paragraph_allocator: ParagraphIdAllocator,
    replacements: dict[str, bytes],
    additions: dict[str, bytes],
    allowed_changes: set[str],
) -> etree._Element:
    """Clone the signature table and import every referenced package part."""

    clone = copy.deepcopy(template.closing_table_node)
    remove_volatile_identity(clone)
    sdt_allocator.apply(clone)
    paragraph_allocator.apply(clone)

    import_signature_styles(
        source, clone, template, replacements, allowed_changes
    )

    # Preserve row properties, fonts, images, and the template's physical grid.
    # Convert percentage widths to fixed twips so A4/Letter and margin changes
    # cannot stretch the table or clip the leftmost signature name.
    stabilize_signature_table_geometry(clone)
    normalize_signature_table_flow(
        clone,
        document_root,
        reserved_drawing_ids=package_drawing_ids(source),
    )
    with ZipFile(source) as package:
        relationship_bytes = replacements.get(
            DOCUMENT_RELS, package.read(DOCUMENT_RELS)
        )
        content_type_bytes = replacements.get(
            CONTENT_TYPES, package.read(CONTENT_TYPES)
        )
        existing_names = set(package.namelist()) | set(additions)
    relationship_root = parse_xml(relationship_bytes)
    content_types_root = parse_xml(content_type_bytes)

    mapping: dict[str, str] = {}
    for old_id, definition in template.closing_relationships.items():
        relationship_type, target_mode, part_name, payload = definition
        new_id = _next_relationship_id(relationship_root)
        if target_mode == "External":
            source_relationship = next(
                element
                for element in clone.iter()
                for attribute, value in element.attrib.items()
                if etree.QName(attribute).namespace == R_NS and value == old_id
            )
            target = source_relationship.get(qn(R_NS, "target"), "")
        else:
            if part_name is None or payload is None:
                raise WorkflowError(
                    f"signature relationship has no package payload: {old_id}"
                )
            suffix = Path(part_name).suffix.lower()
            stem = hashlib.sha256(payload).hexdigest()[:16]
            target_name = f"media/session_minutes_{stem}{suffix}"
            target_part = f"word/{target_name}"
            counter = 1
            while target_part in existing_names and additions.get(target_part) != payload:
                target_name = f"media/session_minutes_{stem}_{counter}{suffix}"
                target_part = f"word/{target_name}"
                counter += 1
            additions[target_part] = payload
            existing_names.add(target_part)
            allowed_changes.add(target_part)
            target = target_name
            extension = suffix.lstrip(".")
            content_type = template.closing_content_types.get(extension)
            if not content_type:
                raise WorkflowError(
                    f"signature image content type is missing for .{extension}"
                )
            _ensure_content_type_default(
                content_types_root, extension, content_type
            )
        attributes = {"Id": new_id, "Type": relationship_type, "Target": target}
        if target_mode:
            attributes["TargetMode"] = target_mode
        etree.SubElement(
            relationship_root,
            qn(PKG_REL_NS, "Relationship"),
            **attributes,
        )
        mapping[old_id] = new_id

    for element in clone.iter():
        for attribute, value in list(element.attrib.items()):
            if etree.QName(attribute).namespace == R_NS and value in mapping:
                element.set(attribute, mapping[value])

    replacements[DOCUMENT_RELS] = xml_bytes(relationship_root)
    replacements[CONTENT_TYPES] = xml_bytes(content_types_root)
    allowed_changes.update({DOCUMENT_RELS, CONTENT_TYPES})
    return clone


def discover_batch_roots(project_root: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in project_root.iterdir()
            if path.is_dir()
            and path.name.isdigit()
            and int(path.name) > 0
        ),
        key=lambda path: int(path.name),
    )


def resolve_batch_root(project_root: Path) -> Path:
    roots = discover_batch_roots(project_root)
    if not roots:
        raise WorkflowError(
            "no top-level numeric batch folder was found"
        )
    if len(roots) != 1:
        raise WorkflowError(
            "expected exactly one top-level numeric batch folder, found: "
            + ", ".join(path.name for path in roots)
        )
    return roots[0].resolve()


def discover_numbered_subfolders(batch_root: Path) -> list[Path]:
    folders = sorted(
        (
            path
            for path in batch_root.iterdir()
            if path.is_dir()
            and path.name.isdigit()
            and int(path.name) > 0
        ),
        key=lambda path: int(path.name),
    )
    if not folders:
        raise WorkflowError(
            f"numeric batch folder {batch_root.name} has no numbered subfolders"
        )
    numbers = [int(path.name) for path in folders]
    expected = list(range(1, max(numbers) + 1))
    if numbers != expected:
        raise WorkflowError(
            "numbered subfolders must be contiguous starting at 1; found: "
            + ", ".join(str(number) for number in numbers)
        )
    return folders


def numbered_ancestor(path: Path, project_root: Path) -> int | None:
    try:
        relative = path.resolve().relative_to(project_root.resolve())
    except ValueError:
        return None
    parts = relative.parts
    if len(parts) < 3:
        return None
    batch_name, numbered_folder = parts[0], parts[1]
    if (
        not batch_name.isdigit()
        or int(batch_name) <= 0
        or not numbered_folder.isdigit()
        or int(numbered_folder) <= 0
    ):
        return None
    return int(numbered_folder)


def discover_targets(project_root: Path, template_path: Path) -> list[Path]:
    batch_root = resolve_batch_root(project_root)
    targets: list[Path] = []
    numbered_folders = discover_numbered_subfolders(batch_root)
    for numbered_folder in numbered_folders:
        for path in numbered_folder.rglob("*.docx"):
            if path.name.startswith("~$"):
                continue
            if path.resolve() == template_path.resolve():
                continue
            if numbered_ancestor(path, project_root) is None:
                continue
            targets.append(path)
    if not targets:
        raise WorkflowError(
            f"no DOCX files were found under numbered folders in {batch_root.name}"
        )
    return sorted(
        targets,
        key=lambda item: item.relative_to(project_root).as_posix().casefold(),
    )


def heading_indices(
    paragraphs: Sequence[etree._Element], heading_norm: str
) -> list[int]:
    return [
        index
        for index, paragraph in enumerate(paragraphs)
        if normalize_text(paragraph_text(paragraph)) == heading_norm
    ]


def is_notice(paragraph: etree._Element) -> bool:
    return normalize_text(paragraph_text(paragraph)).startswith(NOTICE_NORM)


def is_signature(paragraph: etree._Element) -> bool:
    normalized = normalize_text(paragraph_text(paragraph))
    return any(normalized.startswith(prefix) for prefix in SIGNATURE_NORMS)


def signature_table_matches(
    root: etree._Element, template: TemplateData
) -> list[etree._Element]:
    expected_text = normalize_text(paragraph_text(template.closing_table_node))
    return [
        table
        for table in root.xpath(".//w:tbl", namespaces=NS)
        if normalize_text(paragraph_text(table)) == expected_text
        and len(table.xpath(".//w:drawing | .//w:pict", namespaces=NS))
        == len(
            template.closing_table_node.xpath(
                ".//w:drawing | .//w:pict", namespaces=NS
            )
        )
    ]


def duplicate_drawing_ids(root: etree._Element) -> list[str]:
    values = [
        element.get("id", "")
        for element in root.xpath(".//wp:docPr", namespaces=NS)
        if element.get("id")
    ]
    return sorted(value for value, count in Counter(values).items() if count > 1)


def drawing_ids(node: etree._Element) -> set[int]:
    result: set[int] = set()
    for element in node.xpath(".//wp:docPr", namespaces=NS):
        try:
            result.add(int(element.get("id", "")))
        except ValueError:
            continue
    return result


def package_drawing_ids(path: Path, *, include_document: bool = True) -> set[int]:
    result: set[int] = set()
    with ZipFile(path) as package:
        for name in package.namelist():
            if (
                not name.startswith("word/")
                or not name.endswith(".xml")
                or (not include_document and name == DOCUMENT_XML)
            ):
                continue
            root = parse_xml(package.read(name), part_name=name)
            result.update(drawing_ids(root))
    return result


def signature_table_is_inline_after_notice(
    table: etree._Element,
    notice: etree._Element,
    root: etree._Element,
    template: TemplateData,
    reserved_package_ids: set[int] | None = None,
) -> bool:
    parent = notice.getparent()
    if (
        parent is None
        or table.getparent() is not parent
        or parent.index(table) != parent.index(notice) + 1
    ):
        return False
    if signature_table_visual_hash(table) != template.closing_table_visual_hash:
        return False
    if reserved_package_ids and drawing_ids(table) & reserved_package_ids:
        return False
    return not duplicate_drawing_ids(root)


def comment_marker_signature(root: etree._Element) -> dict[str, list[str]]:
    signature: dict[str, list[str]] = {}
    for name in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        signature[name] = sorted(
            element.get(W_ID, "")
            for element in root.xpath(f".//w:{name}", namespaces=NS)
        )
    return signature


def duplicate_sdt_ids(root: etree._Element) -> list[str]:
    values = [
        element.get(W_VAL, "")
        for element in root.xpath(".//w:sdtPr/w:id", namespaces=NS)
        if element.get(W_VAL)
    ]
    return sorted(value for value, count in Counter(values).items() if count > 1)


def paragraph_id_duplicates(root: etree._Element) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    for label, attribute in (("paraId", W14_PARA_ID), ("textId", W14_TEXT_ID)):
        values = [
            element.get(attribute, "")
            for element in root.xpath(".//w:p", namespaces=NS)
            if element.get(attribute)
        ]
        result[label] = sorted(
            value for value, count in Counter(values).items() if count > 1
        )
    return result


def footer_sdt_count(path: Path) -> int:
    total = 0
    with ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith("word/footer") and name.endswith(".xml"):
                root = parse_xml(package.read(name))
                total += len(root.xpath(".//w:sdt", namespaces=NS))
    return total


def table_count(root: etree._Element) -> int:
    return len(root.xpath(".//w:body//w:tbl", namespaces=NS))


def unresolved_token_count(root: etree._Element) -> int:
    text = "\n".join(
        paragraph_text(paragraph) for paragraph in body_paragraphs(root)
    )
    return sum(
        text.count(token)
        for token in ("Choose an item.", "000", "00/00", "…")
    )


def functional_opening_matches(
    paragraphs: Sequence[etree._Element], opening: OpeningTemplate
) -> tuple[int, int, list[int]]:
    text_count = 0
    functional_count = 0
    locations: list[int] = []
    for index, paragraph in enumerate(paragraphs):
        normalized = normalize_text(paragraph_text(paragraph))
        if opening.normalized_text not in normalized:
            continue
        text_count += normalized.count(opening.normalized_text)
        controls = paragraph_control_signature(paragraph)
        if contains_subsequence(controls, opening.control_signature):
            functional_count += 1
            locations.append(index + 1)
    return text_count, functional_count, locations


def current_notice_matches(
    notices: Sequence[etree._Element], template: TemplateData
) -> tuple[int, int]:
    text_count = 0
    functional_count = 0
    for paragraph in notices:
        if normalize_text(paragraph_text(paragraph)) != template.closing_normalized_text:
            continue
        text_count += 1
        if contains_subsequence(
            paragraph_control_signature(paragraph),
            template.closing_control_signature,
        ):
            functional_count += 1
    return text_count, functional_count


def analyze_target(
    path: Path, project_root: Path, template: TemplateData
) -> dict[str, Any]:
    relative = path.relative_to(project_root).as_posix()
    folder_number = numbered_ancestor(path, project_root)
    record: dict[str, Any] = {
        "source_path": relative,
        "source_sha256": sha256_file(path),
        "folder_number": folder_number,
        "status": "ready",
        "blocked_reasons": [],
        "warnings": [],
        "policy": POLICY,
    }
    if folder_number is None or folder_number > len(template.openings):
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "no template opening mapped to numbered folder"
        )
        return record
    opening = template.openings[folder_number - 1]
    record.update(
        {
            "selected_opening_index": folder_number,
            "selected_time": opening.display_time,
            "selected_opening_text": opening.text,
            "selected_opening_xml_hash": opening.xml_hash,
        }
    )
    try:
        parts = read_docx_parts(path)
    except WorkflowError as exc:
        record["status"] = "blocked"
        record["blocked_reasons"].append(str(exc))
        return record
    if parts.protected:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "document editing protection is enforced"
        )
        return record
    has_glossary = parts.glossary_root is not None
    if has_glossary != (parts.glossary_styles_root is not None):
        record["status"] = "blocked"
        record["blocked_reasons"].append("document has an incomplete Word glossary")
        return record
    if has_glossary:
        missing_parts = [
            name for name in GLOSSARY_PARTS if name not in parts.names
        ]
        if missing_parts:
            record["status"] = "blocked"
            record["blocked_reasons"].append(
                "document has an incomplete Word glossary package"
            )
            return record

    paragraphs = body_paragraphs(parts.document_root, nonempty=True)
    basis_indices = heading_indices(paragraphs, HEADING_BASIS_NORM)
    order_indices = heading_indices(paragraphs, HEADING_ORDER_NORM)
    record["basis_heading_count"] = len(basis_indices)
    record["order_heading_count"] = len(order_indices)
    if len(basis_indices) != 1 or len(order_indices) != 1:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "document must contain one unique الأسانيد and one unique منطوق القرار heading"
        )
        return record
    record["classification"] = "standard"
    record["opening_anchor"] = "before_basis"

    text_matches, functional_matches, locations = functional_opening_matches(
        paragraphs, opening
    )
    record["current_opening_text_count"] = text_matches
    record["current_opening_functional_count"] = functional_matches
    record["current_opening_locations"] = locations
    record["current_opening_embedded"] = any(
        paragraph_text(paragraphs[index - 1]) != opening.text
        for index in locations
    )
    if functional_matches > 1:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "current session block is duplicated in source"
        )
        return record
    if text_matches > functional_matches:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "current session text exists without its functional dropdown controls"
        )
        return record
    record["session_action"] = (
        "keep_current" if functional_matches == 1 else "insert_before_basis"
    )
    record["existing_session_paragraph_count"] = sum(
        (
            SESSION_NORM in normalize_text(paragraph_text(paragraph))
            or SESSION_ALT_NORM in normalize_text(paragraph_text(paragraph))
        )
        and normalize_text("عقدت") in normalize_text(paragraph_text(paragraph))
        for paragraph in paragraphs
    )

    order_index = order_indices[0]
    notices = [
        paragraph
        for paragraph in paragraphs[order_index + 1 :]
        if is_notice(paragraph)
    ]
    current_notice_text, current_notice_functional = current_notice_matches(
        notices, template
    )
    record["existing_notice_count"] = len(notices)
    record["current_notice_text_count"] = current_notice_text
    record["current_notice_functional_count"] = current_notice_functional
    if len(notices) == 1 and current_notice_functional == 1:
        record["closing_action"] = "keep_current"
        record["closing_anchor"] = "existing_current_notice"
    elif notices:
        record["closing_action"] = "replace_all_with_current"
        record["closing_anchor"] = "first_existing_notice"
    else:
        record["closing_action"] = "insert_current"
        record["closing_anchor"] = "after_order_before_signatures"

    matching_signature_tables = signature_table_matches(
        parts.document_root, template
    )
    record["existing_signature_table_count"] = len(
        matching_signature_tables
    )
    if len(matching_signature_tables) > 1:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "document contains duplicate current signature tables"
        )
        return record
    if len(matching_signature_tables) == 1:
        current_table_is_complete = (
            record["closing_action"] == "keep_current"
            and len(notices) == 1
            and signature_table_is_inline_after_notice(
                matching_signature_tables[0],
                notices[0],
                parts.document_root,
                template,
                reserved_package_ids=package_drawing_ids(
                    path, include_document=False
                ),
            )
        )
        record["signature_table_action"] = (
            "keep_current" if current_table_is_complete else "normalize_current"
        )
    else:
        record["signature_table_action"] = "insert_current"

    if record["closing_action"] == "insert_current":
        planned_closing_reference = closing_insertion_anchor(
            paragraphs, order_index
        )
    else:
        planned_closing_reference = notices[0]
    planned_closing_parent = planned_closing_reference.getparent()
    record["signature_cell_terminator_delta"] = int(
        record["signature_table_action"] != "keep_current"
        and planned_closing_parent is not None
        and planned_closing_parent.tag == qn(W_NS, "tc")
    )

    comment_nodes_in_notices = sum(
        len(
            paragraph.xpath(
                ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference",
                namespaces=NS,
            )
        )
        for paragraph in notices
    )
    record["comment_markers_in_notices"] = comment_nodes_in_notices
    if comment_nodes_in_notices:
        record["status"] = "blocked"
        record["blocked_reasons"].append(
            "a closing paragraph contains Word comment anchors"
        )
        return record

    main_counts = control_counter(parts.document_root)
    notice_counts = Counter()
    if record["closing_action"] == "replace_all_with_current":
        for notice in notices:
            notice_counts.update(control_counter(notice))
    expected_counts = Counter(main_counts)
    if record["session_action"] == "insert_before_basis":
        expected_counts.update(opening.control_counts)
    if record["closing_action"] == "replace_all_with_current":
        expected_counts.subtract(notice_counts)
        expected_counts.update(template.closing_control_counts)
    elif record["closing_action"] == "insert_current":
        expected_counts.update(template.closing_control_counts)
    expected_counts += Counter()

    record["main_control_types_before"] = dict(sorted(main_counts.items()))
    record["removed_notice_control_types"] = dict(
        sorted(notice_counts.items())
    )
    record["expected_control_types_after"] = dict(
        sorted(expected_counts.items())
    )
    record["date_pickers_removed"] = notice_counts.get("date", 0)
    record["table_count_before"] = table_count(parts.document_root)
    record["signature_table_paragraph_count_before"] = (
        len(matching_signature_tables[0].xpath(".//w:p", namespaces=NS))
        if matching_signature_tables
        else 0
    )
    record["expected_table_count_after"] = (
        record["table_count_before"]
        + (1 if record["signature_table_action"] == "insert_current" else 0)
    )
    record["body_paragraph_count_before"] = len(
        body_paragraphs(parts.document_root)
    )
    session_paragraph_delta = (
        1 if record["session_action"] == "insert_before_basis" else 0
    )
    if record["closing_action"] == "replace_all_with_current":
        closing_delta = 1 - len(notices)
    elif record["closing_action"] == "insert_current":
        closing_delta = 1
    else:
        closing_delta = 0
    record["expected_body_paragraph_count_after"] = (
        record["body_paragraph_count_before"]
        + session_paragraph_delta
        + closing_delta
        + (
            len(template.closing_table_node.xpath(".//w:p", namespaces=NS))
            if record["signature_table_action"] == "insert_current"
            else 0
        )
        + (
            len(template.closing_table_node.xpath(".//w:p", namespaces=NS))
            - int(record["signature_table_paragraph_count_before"])
            if record["signature_table_action"] == "normalize_current"
            else 0
        )
        + int(record["signature_cell_terminator_delta"])
    )
    record["comment_markers_before"] = comment_marker_signature(
        parts.document_root
    )
    record["footer_sdt_count_before"] = footer_sdt_count(path)
    record["glossary_state"] = "existing" if has_glossary else "create"
    record["unresolved_tokens_before"] = unresolved_token_count(
        parts.document_root
    )
    if len(notices) > 1:
        record["warnings"].append(
            f"سيتم توحيد {len(notices)} فقرات ختامية في فقرة واحدة"
        )
    if record["date_pickers_removed"]:
        record["warnings"].append(
            f"سيتم حذف {record['date_pickers_removed']} منتقي تاريخ داخل الخاتمة القديمة"
        )
    if record["existing_session_paragraph_count"] and record[
        "session_action"
    ] == "insert_before_basis":
        record["warnings"].append(
            "ستبقى جميع الجلسات أو النصوص السابقة وتضاف جلسة اليوم"
        )
    if record["signature_table_action"] == "normalize_current":
        record["warnings"].append(
            "سيتم تثبيت جدول التواقيع الحالي مباشرة بعد الخاتمة ومنع انفصاله عنها"
        )
    record["planned_classification"] = (
        "already_complete"
        if record["session_action"] == "keep_current"
        and record["closing_action"] == "keep_current"
        and record["signature_table_action"] == "keep_current"
        else "output"
    )
    return record


def add_duplicate_metadata(records: list[dict[str, Any]]) -> list[list[str]]:
    groups: defaultdict[str, list[str]] = defaultdict(list)
    for record in records:
        groups[record["source_sha256"]].append(record["source_path"])
    duplicates = [sorted(paths) for paths in groups.values() if len(paths) > 1]
    duplicates.sort()
    lookup = {path: group for group in duplicates for path in group}
    for record in records:
        record["duplicate_group"] = lookup.get(record["source_path"], [])
    return duplicates


def manifest_fingerprint(manifest: dict[str, Any]) -> str:
    payload = {
        "schema_version": manifest["schema_version"],
        "project_root": manifest["project_root"],
        "batch": manifest["batch"],
        "policy": manifest["policy"],
        "template": manifest["template"],
        "records": manifest["records"],
        "duplicates": manifest["duplicates"],
        "pilot_paths": manifest["pilot_paths"],
    }
    serialized = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return sha256_bytes(serialized.encode("utf-8"))


def build_manifest(
    project_root: Path, template_path: Path
) -> tuple[dict[str, Any], TemplateData]:
    template = load_template(template_path)
    batch_root = resolve_batch_root(project_root)
    numbered_folders = discover_numbered_subfolders(batch_root)
    if len(template.openings) != len(numbered_folders):
        raise WorkflowError(
            "template opening count does not match numbered subfolder count: "
            f"template={len(template.openings)}, folders={len(numbered_folders)}"
        )
    paths = discover_targets(project_root, template_path)
    records = [
        analyze_target(path, project_root, template) for path in paths
    ]
    duplicates = add_duplicate_metadata(records)
    summary = Counter()
    summary["files"] = len(records)
    summary["ready"] = sum(record["status"] == "ready" for record in records)
    summary["blocked"] = sum(
        record["status"] == "blocked" for record in records
    )
    summary["insert_sessions"] = sum(
        record.get("session_action") == "insert_before_basis"
        for record in records
    )
    summary["keep_current_sessions"] = sum(
        record.get("session_action") == "keep_current"
        for record in records
    )
    summary["replace_closing"] = sum(
        record.get("closing_action") == "replace_all_with_current"
        for record in records
    )
    summary["insert_closing"] = sum(
        record.get("closing_action") == "insert_current"
        for record in records
    )
    summary["keep_current_closing"] = sum(
        record.get("closing_action") == "keep_current"
        for record in records
    )
    summary["insert_signature_tables"] = sum(
        record.get("signature_table_action") == "insert_current"
        for record in records
    )
    summary["normalize_signature_tables"] = sum(
        record.get("signature_table_action") == "normalize_current"
        for record in records
    )
    summary["already_complete"] = sum(
        record.get("planned_classification") == "already_complete"
        for record in records
    )
    summary["date_pickers_removed"] = sum(
        int(record.get("date_pickers_removed", 0)) for record in records
    )
    folder_counts = Counter(record.get("folder_number") for record in records)
    # Freeze the exact fallback-resolved pilot set into the signed preview.
    # Execution and the apply gate must use this list, not select it again.
    pilot_records = select_pilot(records)
    manifest: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "created_at": dt.datetime.now(dt.timezone.utc).isoformat(),
        "project_root": str(project_root.resolve()),
        "batch": {
            "folder_name": batch_root.name,
            "path": batch_root.relative_to(project_root.resolve()).as_posix(),
            "numbered_subfolders": [
                path.name
                for path in numbered_folders
            ],
        },
        "policy": POLICY,
        "template": {
            "path": template_path.relative_to(project_root).as_posix(),
            "sha256": template.sha256,
            "openings": [
                {
                    "folder_number": opening.order,
                    "time": opening.display_time,
                    "text": opening.text,
                    "xml_hash": opening.xml_hash,
                    "control_counts": opening.control_counts,
                }
                for opening in template.openings
            ],
            "closing_text": template.closing_text,
            "closing_xml_hash": template.closing_xml_hash,
            "closing_control_counts": template.closing_control_counts,
            "closing_table_xml_hash": template.closing_table_xml_hash,
            "closing_table_relationships": len(
                template.closing_relationships
            ),
        },
        "records": records,
        "duplicates": duplicates,
        "summary": dict(summary),
        "folder_counts": {
            str(key): value
            for key, value in sorted(folder_counts.items())
            if key is not None
        },
        "pilot_paths": [record["source_path"] for record in pilot_records],
    }
    manifest["preview_id"] = manifest_fingerprint(manifest)
    return manifest, template


def escape_markdown(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def preview_markdown(manifest: dict[str, Any]) -> str:
    summary = manifest["summary"]
    lines = [
        "# معاينة تعبئة محاضر الجلسات",
        "",
        f"- معرّف المعاينة: {manifest['preview_id']}",
        f"- مجلد الدفعة المكتشف: {manifest['batch']['folder_name']}",
        f"- المجلدات الفرعية الرقمية: {', '.join(manifest['batch']['numbered_subfolders'])}",
        f"- عدد الملفات: {summary.get('files', 0)}",
        f"- الجاهزة: {summary.get('ready', 0)}",
        f"- الموقوفة: {summary.get('blocked', 0)}",
        f"- إضافة جلسة: {summary.get('insert_sessions', 0)}",
        f"- جلسة اليوم موجودة ولن تتكرر: {summary.get('keep_current_sessions', 0)}",
        f"- استبدال أو توحيد خاتمة: {summary.get('replace_closing', 0)}",
        f"- إضافة خاتمة مفقودة: {summary.get('insert_closing', 0)}",
        f"- خاتمة اليوم موجودة ولن تتكرر: {summary.get('keep_current_closing', 0)}",
        f"- منتقيات التاريخ المتوقع حذفها داخل الخواتيم القديمة: {summary.get('date_pickers_removed', 0)}",
        "",
        "## ربط المجلدات بكتل النموذج حسب ترتيبها",
        "",
    ]
    for opening in manifest["template"]["openings"]:
        lines.append(
            f"- المجلد {opening['folder_number']} ← {opening['time']}"
        )
    lines.extend(
        [
            "",
            "## ملفات التجربة المحددة",
            "",
        ]
    )
    lines.extend(f"- {path}" for path in manifest["pilot_paths"])
    lines.extend(
        [
            "",
            "## تفاصيل الملفات",
            "",
            "| الحالة | المسار | المجلد | الجلسة | الخاتمة | التحكم قبل/بعد | ملاحظات |",
            "|---|---|---:|---|---|---|---|",
        ]
    )
    for record in manifest["records"]:
        notes = list(record.get("warnings", []))
        notes.extend(record.get("blocked_reasons", []))
        before = sum(record.get("main_control_types_before", {}).values())
        after = sum(
            record.get("expected_control_types_after", {}).values()
        )
        lines.append(
            "| {status} | {path} | {folder} | {session} | {closing} | {before}/{after} | {notes} |".format(
                status=escape_markdown(record["status"]),
                path=escape_markdown(record["source_path"]),
                folder=escape_markdown(record.get("folder_number", "")),
                session=escape_markdown(record.get("session_action", "")),
                closing=escape_markdown(record.get("closing_action", "")),
                before=before,
                after=after,
                notes=escape_markdown("؛ ".join(notes)),
            )
        )
    lines.extend(
        [
            "",
            "## قاعدة التنفيذ",
            "",
            "تضاف جلسة اليوم فقط إذا لم تكن موجودة بعناصرها الوظيفية. تبقى كل الجلسات والنصوص السابقة. تبقى خاتمة اليوم إن كانت وحيدة، وإلا تستبدل جميع الخواتيم بخاتمة اليوم مرة واحدة. لا تعدل المعاينة أي ملف DOCX.",
            "",
        ]
    )
    return "\n".join(lines)


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def write_preview(
    project_root: Path, manifest: dict[str, Any]
) -> tuple[Path, Path]:
    output_dir = project_root / "outputs" / "preview"
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "preview.json"
    md_path = output_dir / "preview.md"
    write_json(json_path, manifest)
    md_path.write_text(preview_markdown(manifest), encoding="utf-8")
    return json_path, md_path


def insert_before(reference: etree._Element, node: etree._Element) -> None:
    parent = reference.getparent()
    if parent is None:
        raise WorkflowError("cannot insert before detached paragraph")
    parent.insert(parent.index(reference), node)


def insert_after(reference: etree._Element, node: etree._Element) -> None:
    parent = reference.getparent()
    if parent is None:
        raise WorkflowError("cannot insert after detached paragraph")
    parent.insert(parent.index(reference) + 1, node)


def closing_insertion_anchor(
    paragraphs: Sequence[etree._Element], order_index: int
) -> etree._Element:
    """Return the last decision-flow paragraph before any signatures.

    A signature can live in a different table cell from the decision.  Using
    the signature itself as an insertion reference would move the closing into
    that signature cell, so the closing is always inserted *after* the last
    preceding non-signature paragraph instead.
    """

    if order_index < 0 or order_index >= len(paragraphs):
        raise WorkflowError("decision heading index is invalid")
    following = list(paragraphs[order_index + 1 :])
    signature_offset = next(
        (
            index
            for index, paragraph in enumerate(following)
            if is_signature(paragraph)
        ),
        None,
    )
    decision_flow = (
        following[:signature_offset]
        if signature_offset is not None
        else following
    )
    return decision_flow[-1] if decision_flow else paragraphs[order_index]


def replace_notices(
    notices: Sequence[etree._Element], closing: etree._Element
) -> None:
    first = notices[0]
    parent = first.getparent()
    if parent is None:
        raise WorkflowError("cannot replace detached notice paragraph")
    parent.replace(first, closing)
    for paragraph in notices[1:]:
        other_parent = paragraph.getparent()
        if other_parent is not None:
            other_parent.remove(paragraph)


def add_glossary_relationship(document_rels: bytes) -> bytes:
    root = parse_xml(document_rels)
    relationships = list(root)
    for relationship in relationships:
        if relationship.get("Type") == GLOSSARY_REL_TYPE:
            return xml_bytes(root)
    used = {relationship.get("Id", "") for relationship in relationships}
    number = 1
    while f"rId{number}" in used:
        number += 1
    etree.SubElement(
        root,
        qn(PKG_REL_NS, "Relationship"),
        Id=f"rId{number}",
        Type=GLOSSARY_REL_TYPE,
        Target="glossary/document.xml",
    )
    return xml_bytes(root)


def add_glossary_content_types(content_types: bytes) -> bytes:
    root = parse_xml(content_types)
    existing = {
        element.get("PartName")
        for element in root.findall(qn(CT_NS, "Override"))
    }
    for part_name, content_type in GLOSSARY_CONTENT_TYPES.items():
        if part_name in existing:
            continue
        etree.SubElement(
            root,
            qn(CT_NS, "Override"),
            PartName=part_name,
            ContentType=content_type,
        )
    return xml_bytes(root)


def initialize_glossary(
    source: Path,
    parts: DocxParts,
    template: TemplateData,
) -> tuple[GlossaryManager, dict[str, bytes], dict[str, bytes], set[str]]:
    replacements: dict[str, bytes] = {}
    additions: dict[str, bytes] = {}
    allowed_changes = {DOCUMENT_XML}
    if parts.glossary_root is not None and parts.glossary_styles_root is not None:
        manager = GlossaryManager(
            parts.glossary_root,
            parts.glossary_styles_root,
            template,
            created=False,
        )
        return manager, replacements, additions, allowed_changes

    try:
        with ZipFile(source) as package:
            if DOCUMENT_RELS not in package.namelist():
                raise WorkflowError(
                    "document has no relationship part for creating a glossary"
                )
            if CONTENT_TYPES not in package.namelist():
                raise WorkflowError(
                    "document has no content-types part for creating a glossary"
                )
            replacements[DOCUMENT_RELS] = add_glossary_relationship(
                package.read(DOCUMENT_RELS)
            )
            replacements[CONTENT_TYPES] = add_glossary_content_types(
                package.read(CONTENT_TYPES)
            )
    except BadZipFile as exc:
        raise WorkflowError("invalid DOCX ZIP package") from exc
    additions.update(template.glossary_parts)
    glossary_root = parse_xml(additions[GLOSSARY_DOCUMENT])
    styles_root = parse_xml(additions[GLOSSARY_STYLES])
    manager = GlossaryManager(
        glossary_root, styles_root, template, created=True
    )
    additions[GLOSSARY_DOCUMENT] = xml_bytes(manager.root)
    additions[GLOSSARY_STYLES] = xml_bytes(manager.styles_root)
    allowed_changes.update(
        {
            DOCUMENT_RELS,
            CONTENT_TYPES,
            *GLOSSARY_PARTS,
        }
    )
    return manager, replacements, additions, allowed_changes


def rewrite_docx(
    source: Path,
    destination: Path,
    replacements: dict[str, bytes],
    additions: dict[str, bytes],
) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(
        prefix=destination.stem + "-",
        suffix=".tmp",
        dir=destination.parent,
    )
    os.close(fd)
    temporary = Path(temporary_name)
    try:
        with ZipFile(source, "r") as source_zip, ZipFile(
            temporary, "w"
        ) as output_zip:
            existing = set()
            for info in source_zip.infolist():
                existing.add(info.filename)
                data = replacements.get(
                    info.filename, source_zip.read(info.filename)
                )
                output_zip.writestr(info, data)
            for name, data in additions.items():
                if name in existing:
                    continue
                info = ZipInfo(name)
                info.compress_type = ZIP_DEFLATED
                info.date_time = dt.datetime.now().timetuple()[:6]
                output_zip.writestr(info, data)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def selected_template(
    template: TemplateData, record: dict[str, Any]
) -> OpeningTemplate:
    index = int(record["selected_opening_index"]) - 1
    try:
        opening = template.openings[index]
    except IndexError as exc:
        raise WorkflowError(
            "preview refers to a missing template opening"
        ) from exc
    if opening.xml_hash != record["selected_opening_xml_hash"]:
        raise WorkflowError("template opening no longer matches preview")
    return opening


def package_changed_parts(source: Path, output: Path) -> list[str]:
    with ZipFile(source) as before, ZipFile(output) as after:
        names = sorted(set(before.namelist()) | set(after.namelist()))
        return [
            name
            for name in names
            if (
                name not in before.namelist()
                or name not in after.namelist()
                or before.read(name) != after.read(name)
            )
        ]


def validate_placeholder_integrity(
    path: Path, parts: DocxParts
) -> dict[str, Any]:
    references = [
        element.get(W_VAL)
        for element in parts.document_root.xpath(
            ".//w:sdtPr/w:placeholder/w:docPart", namespaces=NS
        )
        if element.get(W_VAL)
    ]
    glossary_names = set(GLOSSARY_PARTS) & parts.names
    if not references and not glossary_names:
        return {
            "references": 0,
            "all_resolved": True,
            "package_complete": True,
        }
    missing_parts = sorted(set(GLOSSARY_PARTS) - parts.names)
    if missing_parts:
        raise WorkflowError(
            "generated output has an incomplete glossary package: "
            + ", ".join(missing_parts)
        )
    if parts.glossary_root is None or parts.glossary_styles_root is None:
        raise WorkflowError(
            "generated output contains placeholder references without a glossary"
        )
    docpart_elements = parts.glossary_root.xpath(
        "./w:docParts/w:docPart", namespaces=NS
    )
    docpart_names: list[str] = []
    docpart_guids: list[str] = []
    for element in docpart_elements:
        name_element = element.find("w:docPartPr/w:name", namespaces=NS)
        guid_element = element.find("w:docPartPr/w:guid", namespaces=NS)
        name = name_element.get(W_VAL) if name_element is not None else None
        guid = guid_element.get(W_VAL) if guid_element is not None else None
        if not name or not guid:
            raise WorkflowError(
                "generated output has a glossary docPart without a name or GUID"
            )
        try:
            uuid.UUID(guid.strip("{}"))
        except (ValueError, AttributeError) as exc:
            raise WorkflowError(
                f"generated output has an invalid glossary GUID: {guid}"
            ) from exc
        docpart_names.append(name)
        docpart_guids.append(guid.upper())
    duplicate_names = sorted(
        name for name, count in Counter(docpart_names).items() if count > 1
    )
    duplicate_guids = sorted(
        guid for guid, count in Counter(docpart_guids).items() if count > 1
    )
    if duplicate_names or duplicate_guids:
        raise WorkflowError(
            "generated output has duplicate glossary docPart names or GUIDs"
        )

    style_elements = parts.glossary_styles_root.xpath(
        "./w:style", namespaces=NS
    )
    style_ids = [element.get(W_STYLE_ID) for element in style_elements]
    if any(not style_id for style_id in style_ids):
        raise WorkflowError("generated glossary has a style without an id")
    duplicate_style_ids = sorted(
        style_id
        for style_id, count in Counter(style_ids).items()
        if count > 1
    )
    if duplicate_style_ids:
        raise WorkflowError("generated glossary has duplicate style ids")
    styles = set(style_ids)
    if "PlaceholderText" not in styles:
        raise WorkflowError("generated glossary is missing PlaceholderText style")

    docparts = {
        name: element for name, element in zip(docpart_names, docpart_elements)
    }
    unresolved_references = sorted(
        {name for name in references if name not in docparts}
    )
    unresolved_styles = sorted(
        {
            style_reference.get(W_VAL, "")
            for element in docpart_elements
            for style_reference in element.xpath(".//w:pStyle", namespaces=NS)
            if style_reference.get(W_VAL, "") not in styles
        }
    )
    if unresolved_references or unresolved_styles:
        unresolved = sorted(set(unresolved_references + unresolved_styles))
        raise WorkflowError(
            "generated output has unresolved glossary placeholders or styles: "
            + ", ".join(unresolved)
        )

    try:
        with ZipFile(path) as package:
            relationship_root = parse_xml(package.read(DOCUMENT_RELS))
            content_types_root = parse_xml(package.read(CONTENT_TYPES))
    except (BadZipFile, KeyError) as exc:
        raise WorkflowError(
            "generated output cannot resolve its glossary package metadata"
        ) from exc
    glossary_relationships = [
        relationship
        for relationship in relationship_root.findall(
            qn(PKG_REL_NS, "Relationship")
        )
        if relationship.get("Type") == GLOSSARY_REL_TYPE
    ]
    if (
        len(glossary_relationships) != 1
        or glossary_relationships[0].get("Target", "").replace("\\", "/")
        != "glossary/document.xml"
        or glossary_relationships[0].get("TargetMode")
    ):
        raise WorkflowError(
            "generated output has an invalid glossary relationship"
        )
    overrides = content_types_root.findall(qn(CT_NS, "Override"))
    for part_name, content_type in GLOSSARY_CONTENT_TYPES.items():
        matches = [
            element
            for element in overrides
            if element.get("PartName") == part_name
            and element.get("ContentType") == content_type
        ]
        if len(matches) != 1:
            raise WorkflowError(
                f"generated output has an invalid content type for {part_name}"
            )
    return {
        "references": len(references),
        "docparts": len(docpart_names),
        "styles": len(style_ids),
        "all_resolved": True,
        "package_complete": True,
    }


def validate_signature_table_package(
    path: Path, table: etree._Element, template: TemplateData
) -> dict[str, Any]:
    referenced_relationship_ids = sorted(
        {
            value
            for element in table.iter()
            for attribute, value in element.attrib.items()
            if etree.QName(attribute).namespace == R_NS and value
        }
    )
    expected_reference_count = len(template.closing_relationships)
    if len(referenced_relationship_ids) != expected_reference_count:
        raise WorkflowError(
            "signature table relationship count mismatch: "
            f"expected {expected_reference_count}, found {len(referenced_relationship_ids)}"
        )

    try:
        with ZipFile(path) as package:
            package_names = set(package.namelist())
            relationship_root = parse_xml(package.read(DOCUMENT_RELS))
            content_types_root = parse_xml(package.read(CONTENT_TYPES))
            styles_root = parse_xml(
                package.read(STYLES_XML), part_name=STYLES_XML
            )
    except (BadZipFile, KeyError) as exc:
        raise WorkflowError(
            "generated output cannot resolve signature-table package metadata"
        ) from exc

    relationships = relationship_root.findall(qn(PKG_REL_NS, "Relationship"))
    relationship_ids = [item.get("Id", "") for item in relationships]
    if any(
        count > 1 for count in Counter(relationship_ids).values()
    ):
        raise WorkflowError("generated output has duplicate relationship ids")
    relationship_map = {item.get("Id", ""): item for item in relationships}

    defaults = {
        (item.get("Extension") or "").lower(): item.get("ContentType", "")
        for item in content_types_root.findall(qn(CT_NS, "Default"))
    }
    overrides = {
        item.get("PartName", ""): item.get("ContentType", "")
        for item in content_types_root.findall(qn(CT_NS, "Override"))
    }
    resolved_parts: list[str] = []
    actual_relationship_signatures: Counter[tuple[str, str | None]] = Counter()
    for relationship_id in referenced_relationship_ids:
        relationship = relationship_map.get(relationship_id)
        if relationship is None:
            raise WorkflowError(
                f"signature table has an unresolved relationship: {relationship_id}"
            )
        relationship_type = relationship.get("Type", "")
        target_mode = relationship.get("TargetMode")
        actual_relationship_signatures[(relationship_type, target_mode)] += 1
        if target_mode == "External":
            continue
        target = (relationship.get("Target") or "").replace("\\", "/")
        target_part = posixpath.normpath(
            posixpath.join("word", target)
        ).lstrip("/")
        if target_part.startswith("../") or target_part not in package_names:
            raise WorkflowError(
                f"signature relationship target is missing: {target_part}"
            )
        extension = posixpath.splitext(target_part)[1].lstrip(".").lower()
        content_type = overrides.get("/" + target_part) or defaults.get(extension)
        expected_content_type = template.closing_content_types.get(extension)
        if not expected_content_type or content_type != expected_content_type:
            raise WorkflowError(
                f"signature image has an invalid content type: {target_part}"
            )
        resolved_parts.append(target_part)

    expected_relationship_signatures = Counter(
        (definition[0], definition[1])
        for definition in template.closing_relationships.values()
    )
    if actual_relationship_signatures != expected_relationship_signatures:
        raise WorkflowError(
            "signature table relationship types do not match the template"
        )
    expected_style_hashes = Counter(
        style_definition_visual_hash(style)
        for style in template.closing_styles.values()
    )
    actual_style_hashes = signature_style_hashes(
        styles_root, table, context="generated signature table"
    )
    if actual_style_hashes != expected_style_hashes:
        raise WorkflowError(
            "signature table Word styles do not match the template"
        )
    return {
        "references": len(referenced_relationship_ids),
        "resolved_parts": sorted(resolved_parts),
        "styles": sum(actual_style_hashes.values()),
        "all_resolved": True,
    }


def validate_output(
    source: Path,
    destination: Path,
    record: dict[str, Any],
    template: TemplateData,
    allowed_changes: set[str],
) -> dict[str, Any]:
    output_parts = read_docx_parts(destination)
    if output_parts.protected:
        raise WorkflowError(
            "generated output unexpectedly has editing protection"
        )
    opening = selected_template(template, record)
    paragraphs = body_paragraphs(output_parts.document_root, nonempty=True)
    _, current_opening_count, _ = functional_opening_matches(
        paragraphs, opening
    )
    if current_opening_count != 1:
        raise WorkflowError(
            f"expected one current session, found {current_opening_count}"
        )
    basis_indices = heading_indices(paragraphs, HEADING_BASIS_NORM)
    order_indices = heading_indices(paragraphs, HEADING_ORDER_NORM)
    if len(basis_indices) != 1 or len(order_indices) != 1:
        raise WorkflowError("generated output lost its structural headings")
    if record["session_action"] == "insert_before_basis":
        basis_index = basis_indices[0]
        if basis_index == 0:
            raise WorkflowError("inserted session has no basis anchor")
        inserted = paragraphs[basis_index - 1]
        if semantic_xml_hash(inserted) != opening.xml_hash:
            raise WorkflowError(
                "inserted session formatting or dropdown definitions do not match template"
            )

    order_index = order_indices[0]
    notices = [
        paragraph
        for paragraph in paragraphs[order_index + 1 :]
        if is_notice(paragraph)
    ]
    if len(notices) != 1:
        raise WorkflowError(
            f"expected one current closing, found {len(notices)}"
        )
    _, current_notice_functional = current_notice_matches(
        notices, template
    )
    if current_notice_functional != 1:
        raise WorkflowError(
            "output closing does not match the current functional template closing"
        )
    if record["closing_action"] != "keep_current":
        if closing_semantic_xml_hash(notices[0]) != template.closing_xml_hash:
            raise WorkflowError(
                "inserted closing formatting or dropdown definitions do not match template"
            )
    if record["closing_action"] == "insert_current":
        notice_index = paragraphs.index(notices[0])
        if notice_index <= order_index:
            raise WorkflowError("inserted closing is not after the decision heading")
        if record["signature_table_action"] not in {
            "insert_current", "normalize_current"
        }:
            prior = paragraphs[notice_index - 1]
            if notices[0].getparent() is not prior.getparent():
                raise WorkflowError(
                    "inserted closing left the decision content flow"
                )
        if record["signature_table_action"] not in {
            "insert_current", "normalize_current"
        }:
            signature_index = next(
                (
                    index
                    for index, paragraph in enumerate(paragraphs)
                    if index > order_index and is_signature(paragraph)
                ),
                None,
            )
            if signature_index is not None and notice_index >= signature_index:
                raise WorkflowError("inserted closing is not before signatures")

    actual_counts = control_counter(output_parts.document_root)
    expected_counts = Counter(record["expected_control_types_after"])
    if actual_counts != expected_counts:
        raise WorkflowError(
            f"content-control count mismatch: expected {dict(expected_counts)}, found {dict(actual_counts)}"
        )
    if len(body_paragraphs(output_parts.document_root)) != int(
        record["expected_body_paragraph_count_after"]
    ):
        raise WorkflowError("body paragraph count mismatch")
    actual_table_count = table_count(output_parts.document_root)
    expected_table_count = int(record["expected_table_count_after"])
    if actual_table_count != expected_table_count:
        raise WorkflowError(
            "signature table count changed unexpectedly: "
            f"expected {expected_table_count}, found {actual_table_count}"
        )
    current_signature_tables = signature_table_matches(
        output_parts.document_root, template
    )
    if len(current_signature_tables) != 1:
        raise WorkflowError(
            f"expected one current signature table, found {len(current_signature_tables)}"
        )
    if record["signature_table_action"] in {
        "insert_current", "normalize_current"
    }:
        signature_table = current_signature_tables[0]
        closing_parent = notices[0].getparent()
        if (
            closing_parent is None
            or signature_table.getparent() is not closing_parent
            or closing_parent.index(signature_table)
            != closing_parent.index(notices[0]) + 1
        ):
            raise WorkflowError(
                "signature table is not immediately after the closing paragraph"
            )
        if signature_table_visual_hash(signature_table) != template.closing_table_visual_hash:
            raise WorkflowError(
                "signature table layout, text, or image geometry does not match template"
            )
        if int(record.get("signature_cell_terminator_delta", 0)):
            table_parent = signature_table.getparent()
            next_index = table_parent.index(signature_table) + 1
            if (
                table_parent.tag != qn(W_NS, "tc")
                or next_index >= len(table_parent)
                or table_parent[next_index].tag != qn(W_NS, "p")
            ):
                raise WorkflowError(
                    "nested signature table has no valid trailing paragraph"
                )
    duplicate_drawing_identifiers = duplicate_drawing_ids(
        output_parts.document_root
    )
    if duplicate_drawing_identifiers:
        raise WorkflowError(
            "generated output has duplicate wp:docPr drawing ids: "
            + ", ".join(duplicate_drawing_identifiers)
        )
    cross_part_drawing_ids = drawing_ids(current_signature_tables[0]) & (
        package_drawing_ids(destination, include_document=False)
    )
    if cross_part_drawing_ids:
        raise WorkflowError(
            "signature table drawing ids collide with another DOCX part: "
            + ", ".join(str(value) for value in sorted(cross_part_drawing_ids))
        )
    signature_package_validation = validate_signature_table_package(
        destination, current_signature_tables[0], template
    )
    if footer_sdt_count(destination) != int(
        record["footer_sdt_count_before"]
    ):
        raise WorkflowError("footer content controls changed unexpectedly")
    if comment_marker_signature(output_parts.document_root) != record[
        "comment_markers_before"
    ]:
        raise WorkflowError("Word comment anchors changed unexpectedly")
    duplicate_ids = duplicate_sdt_ids(output_parts.document_root)
    if duplicate_ids:
        raise WorkflowError("generated output has duplicate SDT ids")
    duplicate_paragraph_ids = paragraph_id_duplicates(
        output_parts.document_root
    )
    # Word commonly reuses textId="77777777"; paraId is the identity that
    # must remain unique within the main document.
    if duplicate_paragraph_ids["paraId"]:
        raise WorkflowError(
            "generated output has duplicate paragraph identities"
        )
    placeholder_validation = validate_placeholder_integrity(
        destination, output_parts
    )
    changed_parts = package_changed_parts(source, destination)
    unexpected = sorted(set(changed_parts) - set(allowed_changes))
    if unexpected:
        raise WorkflowError(
            "unexpected DOCX package parts changed: " + ", ".join(unexpected)
        )
    if sha256_file(source) != record["source_sha256"]:
        raise WorkflowError("source changed during output generation")
    return {
        "opening_count_after": current_opening_count,
        "notice_count_after": len(notices),
        "control_types_after": dict(sorted(actual_counts.items())),
        "date_pickers_removed": (
            Counter(record["main_control_types_before"]).get("date", 0)
            - actual_counts.get("date", 0)
        ),
        "table_count_after": table_count(output_parts.document_root),
        "signature_table_count_after": len(current_signature_tables),
        "signature_package_integrity": signature_package_validation,
        "comment_markers_preserved": True,
        "footer_controls_preserved": True,
        "placeholder_integrity": placeholder_validation,
        "changed_package_parts": changed_parts,
        "unexpected_package_parts": [],
        "source_unchanged": True,
    }


def mutate_docx(
    source: Path,
    destination: Path,
    record: dict[str, Any],
    template: TemplateData,
) -> dict[str, Any]:
    source_hash = sha256_file(source)
    if source_hash != record["source_sha256"]:
        raise WorkflowError("source hash does not match preview")
    if record["status"] != "ready":
        raise WorkflowError("blocked record cannot be executed")
    if destination.exists():
        raise WorkflowError(f"refusing to replace existing output: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    no_changes = (
        record["session_action"] == "keep_current"
        and record["closing_action"] == "keep_current"
        and record["signature_table_action"] == "keep_current"
    )
    if no_changes:
        fd, staged_name = tempfile.mkstemp(
            prefix=f".{destination.stem}-staging-",
            suffix=".docx",
            dir=destination.parent,
        )
        os.close(fd)
        staged = Path(staged_name)
        try:
            shutil.copy2(source, staged)
            allowed_changes: set[str] = set()
            validation = validate_output(
                source, staged, record, template, allowed_changes
            )
            if destination.exists():
                raise WorkflowError(
                    f"refusing to replace existing output: {destination}"
                )
            os.replace(staged, destination)
        finally:
            staged.unlink(missing_ok=True)
        return {
            "source_path": record["source_path"],
            "source_sha256": source_hash,
            "output_path": str(destination.resolve()),
            "output_sha256": sha256_file(destination),
            "action": "copied_unchanged",
            "session_action": record["session_action"],
            "closing_action": record["closing_action"],
            "validation": validation,
        }

    parts = read_docx_parts(source)
    if parts.protected:
        raise WorkflowError("source became protected")
    glossary, replacements, additions, allowed_changes = initialize_glossary(
        source, parts, template
    )
    sdt_allocator = SdtIdAllocator(parts.document_root)
    paragraph_allocator = ParagraphIdAllocator(parts.document_root)
    opening_template = selected_template(template, record)

    paragraphs = body_paragraphs(parts.document_root, nonempty=True)
    basis_indices = heading_indices(paragraphs, HEADING_BASIS_NORM)
    order_indices = heading_indices(paragraphs, HEADING_ORDER_NORM)
    if len(basis_indices) != 1 or len(order_indices) != 1:
        raise WorkflowError("structural anchors no longer match preview")
    if record["session_action"] == "insert_before_basis":
        opening = clone_template_paragraph(
            opening_template.node,
            sdt_allocator,
            paragraph_allocator,
            glossary,
        )
        insert_before(paragraphs[basis_indices[0]], opening)

    paragraphs = body_paragraphs(parts.document_root, nonempty=True)
    order_indices = heading_indices(paragraphs, HEADING_ORDER_NORM)
    order_index = order_indices[0]
    notices = [
        paragraph
        for paragraph in paragraphs[order_index + 1 :]
        if is_notice(paragraph)
    ]
    if record["closing_action"] == "replace_all_with_current":
        closing = clone_template_paragraph(
            template.closing_node,
            sdt_allocator,
            paragraph_allocator,
            glossary,
        )
        if not notices:
            raise WorkflowError("closing notices disappeared after preview")
        replace_notices(notices, closing)
    elif record["closing_action"] == "insert_current":
        closing = clone_template_paragraph(
            template.closing_node,
            sdt_allocator,
            paragraph_allocator,
            glossary,
        )
        anchor = closing_insertion_anchor(paragraphs, order_index)
        insert_after(anchor, closing)
    elif record["closing_action"] != "keep_current":
        raise WorkflowError("unknown closing action")

    if record["signature_table_action"] in {
        "insert_current", "normalize_current"
    }:
        updated_paragraphs = body_paragraphs(
            parts.document_root, nonempty=True
        )
        updated_order_index = heading_indices(
            updated_paragraphs, HEADING_ORDER_NORM
        )[0]
        updated_notices = [
            paragraph
            for paragraph in updated_paragraphs[updated_order_index + 1 :]
            if is_notice(paragraph)
        ]
        if len(updated_notices) != 1:
            raise WorkflowError(
                "cannot anchor the current signature table to one closing paragraph"
            )
        if record["signature_table_action"] == "normalize_current":
            matching_tables = signature_table_matches(
                parts.document_root, template
            )
            if len(matching_tables) != 1:
                raise WorkflowError(
                    "cannot normalize a non-unique current signature table"
                )
            existing_table = matching_tables[0]
            existing_parent = existing_table.getparent()
            if existing_parent is None:
                raise WorkflowError("current signature table has no parent")
            existing_parent.remove(existing_table)
        signature_table = clone_template_signature_table(
            source,
            template,
            parts.document_root,
            sdt_allocator,
            paragraph_allocator,
            replacements,
            additions,
            allowed_changes,
        )
        # Anchor after the closing paragraph in the target flow. This is more
        # portable than the template's pre-paragraph floating anchor because
        # source documents have different pagination and section geometry.
        insert_after(updated_notices[0], signature_table)
        if int(record.get("signature_cell_terminator_delta", 0)):
            cell_terminator = etree.Element(qn(W_NS, "p"))
            paragraph_allocator.apply(cell_terminator)
            insert_after(signature_table, cell_terminator)
        in_memory_table_count = table_count(parts.document_root)
        if in_memory_table_count != int(record["expected_table_count_after"]):
            raise WorkflowError(
                "signature table insertion failed before package write: "
                f"expected {record['expected_table_count_after']}, "
                f"found {in_memory_table_count}"
            )
    elif record["signature_table_action"] != "keep_current":
        raise WorkflowError("unknown signature table action")

    replacements[DOCUMENT_XML] = xml_bytes(parts.document_root)
    if glossary.created:
        additions[GLOSSARY_DOCUMENT] = xml_bytes(glossary.root)
        additions[GLOSSARY_STYLES] = xml_bytes(glossary.styles_root)
    elif glossary.added_docparts or glossary.added_styles:
        replacements[GLOSSARY_DOCUMENT] = xml_bytes(glossary.root)
        replacements[GLOSSARY_STYLES] = xml_bytes(glossary.styles_root)
        allowed_changes.update({GLOSSARY_DOCUMENT, GLOSSARY_STYLES})
    fd, staged_name = tempfile.mkstemp(
        prefix=f".{destination.stem}-staging-",
        suffix=".docx",
        dir=destination.parent,
    )
    os.close(fd)
    staged = Path(staged_name)
    try:
        rewrite_docx(source, staged, replacements, additions)
        validation = validate_output(
            source, staged, record, template, allowed_changes
        )
        if destination.exists():
            raise WorkflowError(
                f"refusing to replace existing output: {destination}"
            )
        os.replace(staged, destination)
    finally:
        staged.unlink(missing_ok=True)
    return {
        "source_path": record["source_path"],
        "source_sha256": source_hash,
        "output_path": str(destination.resolve()),
        "output_sha256": sha256_file(destination),
        "action": "modified_copy",
        "session_action": record["session_action"],
        "closing_action": record["closing_action"],
        "glossary_created": glossary.created,
        "glossary_docparts_added": glossary.added_docparts,
        "glossary_styles_added": glossary.added_styles,
        "validation": validation,
    }


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise WorkflowError(
            f"cannot read preview manifest: {exc}"
        ) from exc
    if manifest.get("schema_version") != SCHEMA_VERSION:
        raise WorkflowError("unsupported preview manifest schema")
    if manifest.get("policy") != POLICY:
        raise WorkflowError("preview policy does not match current engine")
    if manifest.get("preview_id") != manifest_fingerprint(manifest):
        raise WorkflowError("preview manifest fingerprint is invalid")
    return manifest


def resolve_template(
    project_root: Path, template_argument: str | None
) -> Path:
    if template_argument:
        path = Path(template_argument)
        if not path.is_absolute():
            path = project_root / path
        return path.resolve()
    preferred = project_root / "نموذج التعبئة للمحاضر.docx"
    if preferred.exists():
        return preferred.resolve()
    root_docx = sorted(project_root.glob("*.docx"))
    if len(root_docx) != 1:
        raise WorkflowError("template is ambiguous; pass --template")
    return root_docx[0].resolve()


def validate_preview_snapshot(
    project_root: Path,
    manifest: dict[str, Any],
    approve_preview_id: str | None,
) -> TemplateData:
    if (
        not approve_preview_id
        or approve_preview_id != manifest["preview_id"]
    ):
        raise WorkflowError(
            "explicit --approve-preview-id must match the preview"
        )
    if Path(manifest["project_root"]).resolve() != project_root.resolve():
        raise WorkflowError("preview belongs to a different project root")
    current_batch = resolve_batch_root(project_root)
    current_subfolders = [
        path.name
        for path in discover_numbered_subfolders(current_batch)
    ]
    if (
        manifest.get("batch", {}).get("path")
        != current_batch.relative_to(project_root.resolve()).as_posix()
        or manifest.get("batch", {}).get("numbered_subfolders")
        != current_subfolders
    ):
        raise WorkflowError(
            "numeric batch folder changed after preview; create a new preview"
        )
    template_path = (
        project_root / manifest["template"]["path"]
    ).resolve()
    template = load_template(template_path)
    if template.sha256 != manifest["template"]["sha256"]:
        raise WorkflowError(
            "template changed after preview; create a new preview"
        )
    current_paths = discover_targets(project_root, template_path)
    current_relative = [
        path.relative_to(project_root).as_posix()
        for path in current_paths
    ]
    preview_relative = [
        record["source_path"] for record in manifest["records"]
    ]
    if current_relative != preview_relative:
        raise WorkflowError(
            "source file set changed after preview; create a new preview"
        )
    for record in manifest["records"]:
        source = project_root / record["source_path"]
        if sha256_file(source) != record["source_sha256"]:
            raise WorkflowError(
                f"source changed after preview: {record['source_path']}"
            )
    return template


def select_pilot(
    records: Sequence[dict[str, Any]], *, allow_fallback: bool = True
) -> list[dict[str, Any]]:
    ready = [record for record in records if record["status"] == "ready"]
    selected: list[dict[str, Any]] = []
    if not allow_fallback:
        return selected

    predicates = (
        lambda record: record.get("session_action") == "keep_current"
        and record.get("current_opening_embedded"),
        lambda record: record.get("existing_notice_count", 0) > 1,
        lambda record: record.get("table_count_before", 0) > 1,
        lambda record: record.get("date_pickers_removed", 0) > 0
        and any(record["comment_markers_before"].values()),
        lambda record: record.get("closing_action") == "insert_current",
    )
    for predicate in predicates:
        match = next(
            (
                record
                for record in ready
                if record not in selected and predicate(record)
            ),
            None,
        )
        if match is not None:
            selected.append(match)
    for record in ready:
        if len(selected) >= PILOT_COUNT:
            break
        if record not in selected:
            selected.append(record)
    return selected[:PILOT_COUNT]


def manifest_pilot_records(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    paths = manifest.get("pilot_paths")
    if not isinstance(paths, list) or not paths:
        raise WorkflowError("preview has no signed pilot file set")
    if len(paths) != len(set(paths)):
        raise WorkflowError("preview pilot file set contains duplicates")
    by_path = {
        record["source_path"]: record for record in manifest.get("records", [])
    }
    try:
        selected = [by_path[path] for path in paths]
    except (KeyError, TypeError) as exc:
        raise WorkflowError("preview pilot file set is invalid") from exc
    if any(record.get("status") != "ready" for record in selected):
        raise WorkflowError("preview pilot file set contains a blocked record")
    return selected


def results_fingerprint(report: dict[str, Any]) -> str:
    payload = {
        "schema_version": report["schema_version"],
        "mode": report["mode"],
        "preview_id": report["preview_id"],
        "output_root": report["output_root"],
        "results": report["results"],
        "blocked": report["blocked"],
        "all_sources_unchanged": report["all_sources_unchanged"],
        "all_validations_passed": report["all_validations_passed"],
    }
    return sha256_bytes(
        json.dumps(
            payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    )


def validate_pilot_gate(
    project_root: Path,
    manifest: dict[str, Any],
    approved: bool,
    pilot_results_argument: str | None = None,
) -> dict[str, Any]:
    if not approved:
        raise WorkflowError(
            "apply requires --approve-pilot-review after opening pilot files in Word"
        )
    preview_id = manifest["preview_id"]
    if pilot_results_argument:
        pilot_path = Path(pilot_results_argument)
        if not pilot_path.is_absolute():
            pilot_path = project_root / pilot_path
        if pilot_path.is_dir():
            pilot_path = pilot_path / "results.json"
    else:
        pilot_path = (
            project_root
            / "outputs"
            / "pilot"
            / preview_id[:12]
            / "results.json"
        )
    pilot_path = pilot_path.resolve()
    if not pilot_path.exists():
        raise WorkflowError(
            "pilot results are missing; create and review the pilot first"
        )
    try:
        report = json.loads(pilot_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise WorkflowError("cannot read pilot results") from exc
    if not isinstance(report, dict):
        raise WorkflowError("pilot results are not a JSON object")
    try:
        fingerprint = results_fingerprint(report)
    except (KeyError, TypeError, ValueError) as exc:
        raise WorkflowError("pilot results are incomplete") from exc
    if (
        report.get("schema_version") != SCHEMA_VERSION
        or report.get("preview_id") != preview_id
        or report.get("mode") != "pilot"
        or not report.get("all_sources_unchanged")
        or not report.get("all_validations_passed")
        or report.get("results_id") != fingerprint
    ):
        raise WorkflowError("pilot results are invalid or do not match preview")
    expected_records = manifest_pilot_records(manifest)
    expected_paths = [record["source_path"] for record in expected_records]
    results = report.get("results")
    if not isinstance(results, list):
        raise WorkflowError("pilot results have no file list")
    if any(not isinstance(result, dict) for result in results):
        raise WorkflowError("pilot results contain an invalid file record")
    actual_paths = [result.get("source_path") for result in results]
    if actual_paths != expected_paths:
        raise WorkflowError(
            "pilot results do not contain the exact signed pilot file set"
        )
    if not report.get("output_root"):
        raise WorkflowError("pilot results have no output folder")
    output_root = Path(report["output_root"]).resolve()
    if output_root != pilot_path.parent.resolve():
        raise WorkflowError("pilot results are not bound to their output folder")
    expected_by_path = {
        record["source_path"]: record for record in expected_records
    }
    for result in results:
        source_path = result["source_path"]
        if result.get("source_sha256") != expected_by_path[source_path][
            "source_sha256"
        ]:
            raise WorkflowError(
                f"pilot source fingerprint is invalid: {source_path}"
            )
        output = Path(result["output_path"]).resolve()
        expected_output = (output_root / source_path).resolve()
        if output != expected_output:
            raise WorkflowError(
                f"pilot output path is invalid: {output}"
            )
        if not output.exists() or sha256_file(output) != result["output_sha256"]:
            raise WorkflowError(
                f"pilot output changed after validation: {output}"
            )
    return report


def execution_markdown(report: dict[str, Any]) -> str:
    lines = [
        f"# نتائج وضع {report['mode']}",
        "",
        f"- معرّف المعاينة: {report['preview_id']}",
        f"- الملفات المنتجة: {len(report['results'])}",
        f"- الملفات الموقوفة: {len(report['blocked'])}",
        f"- الملفات الأصلية دون تغيير: {'نعم' if report['all_sources_unchanged'] else 'لا'}",
        f"- التحقق البنيوي مكتمل: {'نعم' if report['all_validations_passed'] else 'لا'}",
        "",
        "| المصدر | إجراء الجلسة | إجراء الخاتمة | الناتج |",
        "|---|---|---|---|",
    ]
    for result in report["results"]:
        lines.append(
            "| {source} | {session} | {closing} | {output} |".format(
                source=result["source_path"],
                session=result["session_action"],
                closing=result["closing_action"],
                output=result["output_path"],
            )
        )
    if report["blocked"]:
        lines.extend(["", "## الملفات الموقوفة", ""])
        for record in report["blocked"]:
            lines.append(
                f"- {record['source_path']}: {'؛ '.join(record['blocked_reasons'])}"
            )
    lines.append("")
    return "\n".join(lines)


def default_output_root(
    mode: str,
    project_root: Path,
    preview_id: str,
    *,
    now: dt.datetime | None = None,
) -> Path:
    if mode == "pilot":
        return project_root / "outputs" / "pilot" / preview_id[:12]
    if mode != "apply":
        raise WorkflowError(f"unsupported execution mode: {mode}")
    moment = now or dt.datetime.now().astimezone()
    stamp = moment.strftime("%Y%m%d-%H%M%S-%f")
    return (
        project_root
        / "outputs"
        / "final"
        / f"{stamp}-{preview_id[:12]}"
    )


def execute_mode(
    mode: str,
    project_root: Path,
    manifest_path: Path,
    approve_preview_id: str | None,
    output_root_argument: str | None,
    approve_pilot_review: bool,
    pilot_results_argument: str | None = None,
    allow_direct_apply: bool = False,
) -> tuple[dict[str, Any], Path, Path]:
    manifest = load_manifest(manifest_path)
    template = validate_preview_snapshot(
        project_root, manifest, approve_preview_id
    )
    if mode == "apply" and not allow_direct_apply:
        validate_pilot_gate(
            project_root,
            manifest,
            approve_pilot_review,
            pilot_results_argument,
        )
    ready = [
        record for record in manifest["records"] if record["status"] == "ready"
    ]
    selected = manifest_pilot_records(manifest) if mode == "pilot" else list(ready)
    if output_root_argument:
        output_root = Path(output_root_argument)
        if not output_root.is_absolute():
            output_root = project_root / output_root
    else:
        output_root = default_output_root(
            mode, project_root, manifest["preview_id"]
        )
    output_root = output_root.resolve()
    if output_root.exists():
        raise WorkflowError(
            f"refusing to replace existing output directory: {output_root}"
        )
    output_root.parent.mkdir(parents=True, exist_ok=True)
    staging_root = Path(
        tempfile.mkdtemp(
            prefix=f".{output_root.name}-staging-",
            dir=output_root.parent,
        )
    )
    try:
        results: list[dict[str, Any]] = []
        for record in selected:
            source = project_root / record["source_path"]
            staged_destination = staging_root / record["source_path"]
            result = mutate_docx(
                source, staged_destination, record, template
            )
            result["output_path"] = str(
                (output_root / record["source_path"]).resolve()
            )
            results.append(result)
        all_sources_unchanged = all(
            sha256_file(project_root / record["source_path"])
            == record["source_sha256"]
            for record in manifest["records"]
        )
        if not all_sources_unchanged:
            raise WorkflowError(
                "one or more source files changed during execution"
            )
        report: dict[str, Any] = {
            "schema_version": SCHEMA_VERSION,
            "mode": mode,
            "created_at": dt.datetime.now(dt.timezone.utc).isoformat(),
            "preview_id": manifest["preview_id"],
            "output_root": str(output_root),
            "results": results,
            "blocked": [
                record
                for record in manifest["records"]
                if record["status"] != "ready"
            ],
            "all_sources_unchanged": True,
            "all_validations_passed": all(
                result["validation"]["source_unchanged"]
                and not result["validation"]["unexpected_package_parts"]
                for result in results
            ),
        }
        report["results_id"] = results_fingerprint(report)
        write_json(staging_root / "results.json", report)
        (staging_root / "results.md").write_text(
            execution_markdown(report), encoding="utf-8"
        )
        os.replace(staging_root, output_root)
    finally:
        if staging_root.exists():
            shutil.rmtree(staging_root, ignore_errors=True)
    json_path = output_root / "results.json"
    md_path = output_root / "results.md"
    return report, json_path, md_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("preview", "pilot", "apply"))
    parser.add_argument("--project-root", default=".")
    parser.add_argument("--template")
    parser.add_argument("--preview-manifest")
    parser.add_argument("--approve-preview-id")
    parser.add_argument("--output-root")
    parser.add_argument(
        "--pilot-results",
        help="results.json (or its folder) for a pilot made with --output-root",
    )
    parser.add_argument("--approve-pilot-review", action="store_true")
    parser.add_argument(
        "--allow-direct-apply",
        action="store_true",
        help="allow final output after a valid preview without requiring a pilot",
    )
    return parser


def main(arguments: Sequence[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    args = build_parser().parse_args(arguments)
    project_root = Path(args.project_root).resolve()
    try:
        if args.mode == "preview":
            template_path = resolve_template(
                project_root, args.template
            )
            manifest, _ = build_manifest(
                project_root, template_path
            )
            json_path, markdown_path = write_preview(
                project_root, manifest
            )
            result = {
                "mode": "preview",
                "preview_id": manifest["preview_id"],
                "summary": manifest["summary"],
                "json_report": str(json_path.resolve()),
                "markdown_report": str(markdown_path.resolve()),
                "docx_outputs_created": 0,
            }
        else:
            manifest_path = (
                Path(args.preview_manifest).resolve()
                if args.preview_manifest
                else project_root
                / "outputs"
                / "preview"
                / "preview.json"
            )
            report, json_path, markdown_path = execute_mode(
                args.mode,
                project_root,
                manifest_path,
                args.approve_preview_id,
                args.output_root,
                args.approve_pilot_review,
                args.pilot_results,
                args.allow_direct_apply,
            )
            result = {
                "mode": args.mode,
                "preview_id": report["preview_id"],
                "files_created": len(report["results"]),
                "blocked": len(report["blocked"]),
                "output_root": report["output_root"],
                "json_report": str(json_path.resolve()),
                "markdown_report": str(markdown_path.resolve()),
                "results_id": report["results_id"],
            }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except WorkflowError as exc:
        print(
            json.dumps(
                {"status": "error", "message": str(exc)},
                ensure_ascii=False,
                indent=2,
            )
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
