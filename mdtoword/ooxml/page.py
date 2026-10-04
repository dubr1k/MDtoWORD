"""Page setup, text area, proofing language and page-number footers."""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING

from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import Emu, Length, Mm, Twips
from docx.text.paragraph import Paragraph

from ._xml import _DOC_DEFAULTS_SEQUENCE, _RPR_SEQUENCE, _STYLES_SEQUENCE, _get_or_insert
from .fields import add_field

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from docx.section import Section


PAGE_SIZES_MM: dict[str, tuple[float, float]] = {"A4": (210, 297), "Letter": (215.9, 279.4)}
# Word's own defaults when w:pgSz / w:pgMar are absent (US Letter, 1" margins).
_DEFAULT_PAGE = (Twips(12240), Twips(15840))
_DEFAULT_MARGIN = Twips(1440)


def apply_page_setup(
    section: Section,
    page_size: str,
    margins_mm: tuple[float, float, float, float] | None = None,
) -> None:
    """Portrait ``page_size`` (a PAGE_SIZES_MM key, case-insensitive) and optional
    margins given as (top, right, bottom, left) millimetres."""
    size = next((value for key, value in PAGE_SIZES_MM.items() if key.lower() == page_size.lower()), None)
    if size is None:
        raise ValueError(f"unknown page size {page_size!r}; expected one of {sorted(PAGE_SIZES_MM)}")
    section.orientation = WD_ORIENTATION.PORTRAIT
    section.page_width = Mm(size[0])
    section.page_height = Mm(size[1])
    if margins_mm is not None:
        if len(margins_mm) != 4 or any(not math.isfinite(m) or m < 0 for m in margins_mm):
            raise ValueError(f"margins must be four non-negative numbers, got {margins_mm!r}")
        top, right, bottom, left = margins_mm
        section.top_margin = Mm(top)
        section.right_margin = Mm(right)
        section.bottom_margin = Mm(bottom)
        section.left_margin = Mm(left)


def _or_default(value: Length | None, default: Length) -> int:
    return int(value) if value is not None else int(default)


def text_width(section: Section) -> Length:
    """Width available to body text: page width minus left and right margins."""
    width = _or_default(section.page_width, _DEFAULT_PAGE[0])
    width -= _or_default(section.left_margin, _DEFAULT_MARGIN)
    width -= _or_default(section.right_margin, _DEFAULT_MARGIN)
    return Emu(max(0, width))


def text_height(section: Section) -> Length:
    """Height available to body text: page height minus top and bottom margins."""
    height = _or_default(section.page_height, _DEFAULT_PAGE[1])
    height -= _or_default(section.top_margin, _DEFAULT_MARGIN)
    height -= _or_default(section.bottom_margin, _DEFAULT_MARGIN)
    return Emu(max(0, height))


_LANGUAGE_TAG = re.compile(r"^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$")


def set_document_language(document: DocxDocument, tag: str) -> None:
    """Make ``tag`` (e.g. ``"ru-RU"``) the document's default proofing language."""
    if not _LANGUAGE_TAG.match(tag):
        raise ValueError(f"invalid language tag {tag!r}")
    styles = document.styles.element
    doc_defaults = _get_or_insert(styles, "w:docDefaults", _STYLES_SEQUENCE)
    rpr_default = _get_or_insert(doc_defaults, "w:rPrDefault", _DOC_DEFAULTS_SEQUENCE)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_default.append(rpr)
    lang = _get_or_insert(rpr, "w:lang", _RPR_SEQUENCE)
    lang.set(qn("w:val"), tag)
    if lang.get(qn("w:eastAsia")) == "en-US":
        lang.set(qn("w:eastAsia"), tag)
    document.core_properties.language = tag


def _is_cyrillic(char: str) -> bool:
    return "Ѐ" <= char <= "ԯ" or "ᲀ" <= char <= "᲏" or (
        "ⷠ" <= char <= "ⷿ"
    ) or "Ꙁ" <= char <= "ꚟ"


def detect_language(text: str) -> str:
    """``"ru-RU"`` when Cyrillic makes up at least 30% of the letters, else ``"en-US"``."""
    letters = cyrillic = 0
    for char in text:
        if char.isalpha():
            letters += 1
            if _is_cyrillic(char):
                cyrillic += 1
    if letters and cyrillic / letters >= 0.3:
        return "ru-RU"
    return "en-US"


def add_page_number_footer(
    section: Section, alignment: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER
) -> Paragraph:
    """Put a PAGE field into the section's footer (unlinked from the previous one)."""
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraphs = footer.paragraphs
    if len(paragraphs) == 1 and not paragraphs[0]._p.xpath("./w:r | ./w:hyperlink | ./w:fldSimple"):
        paragraph = paragraphs[0]
    else:
        paragraph = footer.add_paragraph()
    paragraph.alignment = alignment
    add_field(paragraph, "PAGE", "1", dirty=False)
    return paragraph
