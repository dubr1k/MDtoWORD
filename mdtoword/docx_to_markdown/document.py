"""One document's conversion: the shared state, the story walk and the final assembly.

:class:`_Conversion` owns everything that outlives a single paragraph --
styles, numbering counters, the paragraph classifier, open fields, heading
anchors, notes, pictures and warnings -- and puts the front matter, the
body and the note definitions together.
"""

from __future__ import annotations

from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT

from ..errors import ConversionWarning
from .anchors import heading_anchors
from .blocks import _BlockWriter
from .classify import _Classifier
from .collect import _InlineCollector
from .diagnostics import _Diagnostics
from .media import _Media, _Pictures
from .metadata import front_matter, without_title_block
from .notes import _Notes
from .numbering import _Numbering
from .styles import _Styles
from .toc import _TOC_PLACEHOLDER, is_toc_control
from .wordml import W_CUSTOMXML, W_INS, W_MOVETO, W_P, W_SDT, W_SDTCONTENT, W_TBL, _q, part_element


class _Conversion:
    """State shared by one document conversion: styles, counters, fields, warnings."""

    def __init__(self, document: Any, media: _Media, *, front_matter: bool) -> None:
        self.document = document
        self.main_part = document.part
        self.include_front_matter = front_matter
        self.diagnostics = _Diagnostics()
        self.styles = _Styles(part_element(self.main_part, RT.STYLES))
        self.numbering = _Numbering(part_element(self.main_part, RT.NUMBERING), self.styles)
        self.classifier = _Classifier(self.styles, self.numbering)
        # Heading anchors by bookmark name, filled before anything is written.
        self.bookmark_slugs: dict[str, str] = {}
        self.pictures = _Pictures(media, self.diagnostics)
        self.notes = _Notes(self.main_part)
        self.inline = _InlineCollector(self.classifier, self.diagnostics, self.pictures,
                                       self.notes, self.bookmark_slugs)

    def warnings(self) -> list[ConversionWarning]:
        return self.diagnostics.warnings()

    def convert(self) -> str:
        body = self.document.element.body
        elements = self.block_elements(body)
        # Links may point forward, so every heading's anchor is known first.
        self.bookmark_slugs.update(
            heading_anchors(body, elements, self.classifier, self.numbering))
        blocks: list[str] = []
        matter = front_matter(self.document) if self.include_front_matter else None
        if matter:
            blocks.append(matter)
            elements = without_title_block(elements, self.document, self.classifier)
        blocks.extend(_BlockWriter(self, self.main_part).render(elements))
        blocks.extend(self.notes.definitions(self))
        text = "\n\n".join(block for block in blocks if block.strip())
        return text + "\n" if text else ""

    def block_elements(self, container: Any) -> list[Any]:
        """Paragraphs and tables of a container, content controls opened."""
        result: list[Any] = []
        for child in container:
            tag = child.tag
            if tag in (W_P, W_TBL):
                result.append(child)
            elif tag == W_SDT:
                if is_toc_control(child):
                    result.append(_TOC_PLACEHOLDER)
                    continue
                content = child.find(W_SDTCONTENT)
                if content is not None:
                    result.extend(self.block_elements(content))
            elif tag in (W_CUSTOMXML, W_INS, W_MOVETO):
                result.extend(self.block_elements(child))
            elif tag == _q("w:altChunk"):
                self.diagnostics.warn("altchunk_skipped")
        return result
