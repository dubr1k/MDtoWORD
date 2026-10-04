"""Инструмент inspect_docx: сводка структуры готового .docx."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Annotated, Any

import anyio
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from mcp.types import ToolAnnotations
from pydantic import Field

from .models import DocxCounts, DocxInspection, DocxProperties, OutlineEntry, PageSetup
from .server import mcp

_PAPER_SIZES_MM = {
    "A4": (210.0, 297.0),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
    "A3": (297.0, 420.0),
    "A5": (148.0, 210.0),
}
_PAPER_TOLERANCE_MM = 1.0
_TEXT_PREVIEW_LENGTH = 500
_HEADING_STYLE = re.compile(r"heading ([1-9])")
_TOC_INSTRUCTION = re.compile(r"^\s*TOC\b")


def _mm(length: Any) -> float | None:
    return None if length is None else round(length.mm, 1)


def _paper_name(width: float | None, height: float | None) -> str | None:
    if width is None or height is None:
        return None
    short, long = sorted((width, height))
    for name, (paper_short, paper_long) in _PAPER_SIZES_MM.items():
        if (
            abs(short - paper_short) <= _PAPER_TOLERANCE_MM
            and abs(long - paper_long) <= _PAPER_TOLERANCE_MM
        ):
            return name
    return "custom"


def _page_setup(section: Any) -> PageSetup:
    width, height = _mm(section.page_width), _mm(section.page_height)
    orientation = None
    if width is not None and height is not None:
        orientation = "landscape" if width > height else "portrait"
    return PageSetup(
        width_mm=width,
        height_mm=height,
        paper=_paper_name(width, height),
        orientation=orientation,
        margin_top_mm=_mm(section.top_margin),
        margin_bottom_mm=_mm(section.bottom_margin),
        margin_left_mm=_mm(section.left_margin),
        margin_right_mm=_mm(section.right_margin),
    )


class _StyleInfo:
    """Уровень заголовка и нумерация стиля с учётом цепочки base_style."""

    def __init__(self, document: Any) -> None:
        self._styles = {style.style_id: style for style in document.styles}
        self._heading: dict[str | None, int | None] = {}
        self._numbered: dict[str | None, bool] = {}

    def _chain(self, style_id: str | None) -> list[Any]:
        chain: list[Any] = []
        style = self._styles.get(style_id) if style_id else None
        while style is not None and style not in chain:
            chain.append(style)
            style = getattr(style, "base_style", None)
        return chain

    def heading_level(self, style_id: str | None) -> int | None:
        if style_id not in self._heading:
            level = None
            for style in self._chain(style_id):
                name = (style.name or "").strip().lower()
                if name == "title":
                    level = 0
                    break
                match = _HEADING_STYLE.fullmatch(name)
                if match:
                    level = int(match.group(1))
                    break
                outline = style.element.xpath("./w:pPr/w:outlineLvl/@w:val")
                if outline and outline[0].isdigit() and int(outline[0]) <= 8:
                    level = int(outline[0]) + 1
                    break
            self._heading[style_id] = level
        return self._heading[style_id]

    def numbered(self, style_id: str | None) -> bool:
        if style_id not in self._numbered:
            numbered = False
            for style in self._chain(style_id):
                num_id = style.element.xpath("./w:pPr/w:numPr/w:numId/@w:val")
                if num_id:
                    numbered = num_id[0] != "0"
                    break
            self._numbered[style_id] = numbered
        return self._numbered[style_id]


def _paragraph_text(paragraph: Any) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t"), qn("m:t")))


def _count_footnotes(document: Any) -> int:
    for relationship in document.part.rels.values():
        if relationship.reltype != RELATIONSHIP_TYPE.FOOTNOTES or relationship.is_external:
            continue
        root = parse_xml(relationship.target_part.blob)
        return sum(
            1
            for footnote in root.iterchildren(qn("w:footnote"))
            if footnote.get(qn("w:type"), "normal") == "normal"
        )
    return 0


def _has_toc(body: Any) -> bool:
    instructions = body.xpath(".//w:instrText/text()") + body.xpath(".//w:fldSimple/@w:instr")
    if any(_TOC_INSTRUCTION.match(instruction) for instruction in instructions):
        return True
    galleries = body.xpath(".//w:sdtPr/w:docPartObj/w:docPartGallery/@w:val")
    return any("table of contents" in gallery.lower() for gallery in galleries)


def _inspect(path: Path) -> DocxInspection:
    if not path.exists():
        raise ValueError(f"File not found: {path}")
    if not path.is_file():
        raise ValueError(f"Not a file: {path}")
    try:
        document = Document(str(path))
    except Exception as error:
        raise ValueError(
            f"Not a readable .docx document: {path} ({type(error).__name__}: {error})"
        ) from error

    body = document.element.body
    styles = _StyleInfo(document)
    outline: list[OutlineEntry] = []
    paragraphs = list_paragraphs = 0
    preview_parts: list[str] = []
    preview_length = 0

    for paragraph in body.iter(qn("w:p")):
        text = _paragraph_text(paragraph)
        style_ids = paragraph.xpath("./w:pPr/w:pStyle/@w:val")
        style_id = style_ids[0] if style_ids else None

        level = styles.heading_level(style_id)
        outline_level = paragraph.xpath("./w:pPr/w:outlineLvl/@w:val")
        if level is None and outline_level and outline_level[0].isdigit():
            level = int(outline_level[0]) + 1 if int(outline_level[0]) <= 8 else None
        if level is not None and text.strip():
            outline.append(OutlineEntry(level=level, text=text.strip()))
        elif level is None:
            num_id = paragraph.xpath("./w:pPr/w:numPr/w:numId/@w:val")
            if (num_id[0] != "0") if num_id else styles.numbered(style_id):
                list_paragraphs += 1

        if text.strip() or paragraph.xpath(".//w:drawing"):
            paragraphs += 1
        if text.strip() and preview_length < _TEXT_PREVIEW_LENGTH:
            preview_parts.append(text.strip())
            preview_length += len(text.strip()) + 1

    preview = "\n".join(preview_parts)
    if len(preview) > _TEXT_PREVIEW_LENGTH:
        preview = preview[:_TEXT_PREVIEW_LENGTH].rstrip() + "…"

    images = sum(
        1
        for drawing in body.xpath(".//wp:inline | .//wp:anchor")
        if drawing.xpath(".//a:blip")
    )
    core = document.core_properties
    language = document.styles.element.xpath(
        "./w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val"
    )
    return DocxInspection(
        path=str(path),
        page=_page_setup(document.sections[0]),
        section_count=len(document.sections),
        counts=DocxCounts(
            paragraphs=paragraphs,
            headings=len(outline),
            tables=len(document.tables),
            images=images,
            equations=len(body.xpath(".//m:oMath[not(ancestor::m:oMath)]")),
            footnotes=_count_footnotes(document),
            hyperlinks=len(body.xpath(".//w:hyperlink")),
            list_paragraphs=list_paragraphs,
        ),
        has_toc=_has_toc(body),
        outline=outline,
        properties=DocxProperties(
            title=core.title or None,
            author=core.author or None,
            subject=core.subject or None,
            keywords=core.keywords or None,
            language=core.language or None,
        ),
        language=language[0] if language else None,
        text_preview=preview,
    )


_INSPECT_DOCX_TITLE = "Inspect Word document"


@mcp.tool(
    title=_INSPECT_DOCX_TITLE,
    annotations=ToolAnnotations(
        title=_INSPECT_DOCX_TITLE,
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def inspect_docx(
    path: Annotated[str, Field(min_length=1, description="Path of a .docx file")],
) -> DocxInspection:
    """Summarise the structure of a Word .docx file without modifying it.

    Use it to verify a `markdown_to_word` result (or any .docx) instead of
    trusting the conversion blindly: page size, orientation and margins in
    millimetres (first section) and the section count; counts of
    paragraphs, headings, tables, images, native equations, footnotes,
    hyperlinks and list items; whether a table of contents is present; the
    heading outline (Title = level 0, Heading N = level N); core document
    properties; the default proofing language; and the first ~500
    characters of body text.

    Typical checks: `counts.equations` matches the number of formulas in
    the source (a lower number means some were kept as text); the outline
    matches the Markdown headings; `page.paper` and margins match the
    requested preset ("gost" → A4, margins 20/20/30/15 mm top/bottom/
    left/right).

    Fails with a clear error if the path does not exist or is not a
    readable .docx file.
    """
    return await anyio.to_thread.run_sync(
        _inspect, Path(path).expanduser().resolve()
    )
