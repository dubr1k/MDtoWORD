"""Real Word numbering (``word/numbering.xml``) for Markdown lists."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.parts.numbering import NumberingPart
from docx.shared import Length, Twips
from docx.text.paragraph import Paragraph

from ._xml import _free_partname, _max_int_attribute, _xml_attr

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument


_MAX_LIST_LEVEL = 8
_LIST_INDENT_STEP_TWIPS = 720
_LIST_HANGING_TWIPS = 360
_BULLET_SYMBOLS = ("•", "◦", "▪")  # bullet, white bullet, small square


def _clamp_level(level: int) -> int:
    return max(0, min(_MAX_LIST_LEVEL, int(level)))


def _numbering_element(document: DocxDocument) -> Any:
    """``w:numbering`` root, creating the numbering part when the document has none.

    python-docx's own ``numbering_part`` raises ``NotImplementedError`` when a
    template lacks numbering.xml, so the part is created here instead.
    """
    document_part = document.part
    try:
        part = document_part.part_related_by(RT.NUMBERING)
    except KeyError:
        package = document_part.package
        partname = _free_partname(package, "/word/numbering.xml", "/word/numbering%d.xml")
        part = NumberingPart(
            partname, CT.WML_NUMBERING, parse_xml(f"<w:numbering {nsdecls('w')}/>"), package
        )
        document_part.relate_to(part, RT.NUMBERING)
    return part.element


class ListNumbering:
    """Allocates Word numbering definitions so each Markdown list numbers independently.

    Two abstract definitions (decimal and bullet, 9 levels each) are created
    lazily, once per instance; every :meth:`start_list` call then creates a
    fresh ``w:num`` pointing at one of them. Ordered lists carry a
    ``w:startOverride`` on *every* level, so a list restarts at ``start``
    whichever level the renderer applies it at (a nested list is applied at
    level 1+ with its own numId).

    Indentation mirrors Word's built-in lists: level ``n`` has its text at
    ``720 * (n + 1)`` twips with a 360-twip hanging marker.
    """

    def __init__(self, document: DocxDocument, *, bullet_font: str | None = None) -> None:
        self._document = document
        self._bullet_font = bullet_font
        self._abstract_ids: dict[bool, int] = {}

    def start_list(self, ordered: bool, start: int = 1) -> int:
        """Return a NEW numId for one Markdown list (ordered lists restart at ``start``)."""
        numbering = _numbering_element(self._document)
        abstract_id = self._abstract_id(numbering, bool(ordered))
        num_id = _max_int_attribute(numbering.xpath("./w:num/@w:numId"), 0) + 1
        num = OxmlElement("w:num", {qn("w:numId"): str(num_id)})
        num.append(OxmlElement("w:abstractNumId", {qn("w:val"): str(abstract_id)}))
        if ordered:
            start_value = str(max(0, int(start)))
            for level in range(_MAX_LIST_LEVEL + 1):
                override = OxmlElement("w:lvlOverride", {qn("w:ilvl"): str(level)})
                override.append(OxmlElement("w:startOverride", {qn("w:val"): start_value}))
                num.append(override)
        cleanup = numbering.find(qn("w:numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(num)
        else:
            numbering.append(num)
        return num_id

    def apply(self, paragraph: Paragraph, num_id: int, level: int) -> None:
        """Make ``paragraph`` an item of list ``num_id`` at ``level`` (clamped to 0..8)."""
        num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        num_pr.get_or_add_ilvl().val = _clamp_level(level)
        num_pr.get_or_add_numId().val = int(num_id)

    def continuation_indent(self, level: int) -> Length:
        """Left indent aligning a non-numbered paragraph with the item text at ``level``."""
        return Twips(_LIST_INDENT_STEP_TWIPS * (_clamp_level(level) + 1))

    # -- internals ---------------------------------------------------------

    def _abstract_id(self, numbering: Any, ordered: bool) -> int:
        cached = self._abstract_ids.get(ordered)
        if cached is not None and numbering.xpath(
            f'./w:abstractNum[@w:abstractNumId="{cached}"]'
        ):
            return cached
        abstract_id = _max_int_attribute(numbering.xpath("./w:abstractNum/@w:abstractNumId"), -1) + 1
        abstract = parse_xml(self._abstract_xml(abstract_id, ordered))
        # Every w:abstractNum must precede every w:num.
        first_num = numbering.find(qn("w:num"))
        anchor = first_num if first_num is not None else numbering.find(qn("w:numIdMacAtCleanup"))
        if anchor is not None:
            anchor.addprevious(abstract)
        else:
            numbering.append(abstract)
        self._abstract_ids[ordered] = abstract_id
        return abstract_id

    def _abstract_xml(self, abstract_id: int, ordered: bool) -> str:
        nsid = f"{(0x4D440000 + abstract_id * 2 + int(ordered)) & 0xFFFFFFFF:08X}"
        levels = "".join(self._level_xml(level, ordered) for level in range(_MAX_LIST_LEVEL + 1))
        return (
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
            f'<w:nsid w:val="{nsid}"/>'
            '<w:multiLevelType w:val="hybridMultilevel"/>'
            f"{levels}</w:abstractNum>"
        )

    def _level_xml(self, level: int, ordered: bool) -> str:
        left = _LIST_INDENT_STEP_TWIPS * (level + 1)
        ppr = f'<w:pPr><w:ind w:left="{left}" w:hanging="{_LIST_HANGING_TWIPS}"/></w:pPr>'
        if ordered:
            return (
                f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
                f'<w:lvlText w:val="%{level + 1}."/><w:lvlJc w:val="left"/>{ppr}</w:lvl>'
            )
        symbol = _BULLET_SYMBOLS[level % len(_BULLET_SYMBOLS)]
        if self._bullet_font:
            font = _xml_attr(self._bullet_font)
            fonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/>'
        else:
            # No explicit face: the marker inherits the body font, and the WGL4
            # repertoire every common body font covers includes all three glyphs.
            fonts = '<w:rFonts w:hint="default"/>'
        return (
            f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="{symbol}"/><w:lvlJc w:val="left"/>{ppr}'
            f"<w:rPr>{fonts}</w:rPr></w:lvl>"
        )
