"""The Word document itself: template, page setup, styles and properties.

Creates the document (blank, or a user template stripped of its body text
and orphaned notes), applies page size, margins and language, configures
the built-in styles the renderer uses (fonts, heading sizes, GOST spacing),
adds the project's own ``Source Code`` style, fixes Word's justification of
lines ending in a manual break, and copies front-matter metadata into the
core document properties.
"""

from __future__ import annotations

from pathlib import Path
from typing import cast

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.styles.style import ParagraphStyle

from ..ooxml import (
    add_page_number_footer,
    apply_page_setup,
    set_document_language,
    set_paragraph_borders,
    set_paragraph_shading,
    text_height,
    text_width,
)
from .constants import (
    _BLACK,
    _CODE_BLOCK_BORDER,
    _CODE_BLOCK_SHADING,
    _CODE_FONT,
    _COMPAT_BEFORE_SHIFT_RETURN,
    _DEFAULT_MARGINS,
    _GOST_FIRST_LINE,
    _GOST_MARGINS,
    _GOST_USER_MARGINS,
    _HEADING_SCALE,
    _MAX_PROPERTY_LENGTH,
    _XML_INVALID,
)
from .helpers import _set_style_font
from .state import RendererState


class DocumentSetupMixin(RendererState):
    """Create and configure the document before any content goes in."""

    def _new_document(self) -> DocumentType:
        if self.options.template is None:
            return Document()
        template = Path(self.options.template)
        if not template.is_file():
            raise FileNotFoundError(f"Template not found: {template}")
        document = Document(str(template))
        # Only the styles, numbering, section setup, headers and footers of
        # the template are wanted -- never its body text.
        body = document.element.body
        for child in list(body):
            if child.tag != qn("w:sectPr"):
                body.remove(child)
        # Notes belonging to the removed text would survive as orphans.
        for relationship in document.part.rels.values():
            if not relationship.reltype.endswith(("/footnotes", "/endnotes")):
                continue
            notes = getattr(relationship.target_part, "element", None)
            if notes is None:
                continue
            for note in list(notes):
                if note.get(qn("w:type")) is None:
                    notes.remove(note)
        return document

    def _configure_document(self) -> None:
        section = self.document.sections[0]
        if self.options.template is None:
            margins = _GOST_MARGINS if self._gost else _DEFAULT_MARGINS
            if self.options.preset == "gost_user":
                margins = _GOST_USER_MARGINS
            apply_page_setup(section, self.options.page_size or "A4", margins)
        elif self.options.page_size is not None:
            apply_page_setup(section, self.options.page_size)
        set_document_language(self.document, self._language)
        if self.options.template is None:
            self._configure_styles()
            if self.options.preset == "gost":
                add_page_number_footer(section)
        self._ensure_custom_styles()
        self._compat_do_not_expand_shift_return()
        self._text_width = text_width(section)
        self._text_height = text_height(section)
        self._apply_front_matter_properties()

    def _style(self, name: str, base: str | None = "Normal") -> ParagraphStyle:
        """Return a paragraph style, creating a plain one if the document lacks it.

        A user template need not define every built-in style the renderer
        uses; asking python-docx for a missing one would raise KeyError.
        """
        try:
            return cast(ParagraphStyle, self.document.styles[name])
        except KeyError:
            style = cast(
                ParagraphStyle, self.document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            )
            if base is not None:
                try:
                    style.base_style = self.document.styles[base]
                except KeyError:
                    pass
            return style

    def _configure_styles(self) -> None:
        normal = cast(ParagraphStyle, self.document.styles["Normal"])
        _set_style_font(normal, self.font_name)
        normal.font.size = self.font_size
        normal.font.color.rgb = _BLACK
        if self._gost:
            normal_format = normal.paragraph_format
            normal_format.line_spacing = 1.5
            if self.options.preset == "gost_user":
                normal_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            normal_format.first_line_indent = _GOST_FIRST_LINE
            normal_format.space_before = Pt(0)
            normal_format.space_after = Pt(0)
            normal_format.widow_control = True

        for level in range(1, 10):
            try:
                heading = cast(ParagraphStyle, self.document.styles[f"Heading {level}"])
            except KeyError:
                continue
            heading.font.color.rgb = _BLACK
            _set_style_font(heading, self.font_name)
            if self._gost:
                heading.font.size = self.font_size
                heading.font.bold = True
                heading.font.italic = False
                heading_format = heading.paragraph_format
                heading_format.first_line_indent = _GOST_FIRST_LINE
                heading_format.space_before = Pt(12)
                heading_format.space_after = Pt(12)
                heading_format.keep_with_next = True
                heading_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                continue
            scale = _HEADING_SCALE.get(level)
            if scale is not None:
                multiplier, bold, italic = scale
                heading.font.size = Pt(round(self.font_size.pt * multiplier))
                heading.font.bold = bold
                heading.font.italic = italic

        for name in ("Quote", "List Paragraph", "List Bullet", "List Number", "Caption",
                     "Title", "Subtitle"):
            style = self._style(name)
            _set_style_font(style, self.font_name)
            style.font.color.rgb = _BLACK

        quote = self._style("Quote")
        quote.font.italic = False

        caption = self._style("Caption")
        caption.font.bold = False
        caption.font.italic = not self._gost
        caption.font.size = self.font_size if self._gost else Pt(max(8, round(self.font_size.pt * 0.9)))
        caption.paragraph_format.space_before = Pt(0 if self._gost else 3)
        caption.paragraph_format.space_after = Pt(6 if self._gost else 10)
        caption.paragraph_format.first_line_indent = Pt(0)

        title = self._style("Title")
        title.font.size = Pt(round(self.font_size.pt * 1.75))
        title.font.bold = True
        title_ppr = title.element.get_or_add_pPr()
        for border in title_ppr.findall(qn("w:pBdr")):
            title_ppr.remove(border)
        title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.paragraph_format.first_line_indent = Pt(0)
        subtitle = self._style("Subtitle")
        subtitle.font.size = Pt(round(self.font_size.pt * 1.2))
        subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.paragraph_format.first_line_indent = Pt(0)

    def _ensure_custom_styles(self) -> None:
        """Styles this project defines itself; created even under a template."""
        try:
            self.document.styles["Source Code"]
        except KeyError:
            code = self._style("Source Code")
            _set_style_font(code, _CODE_FONT)
            code.font.size = Pt(max(8, round(self.font_size.pt * 0.85)))
            code_format = code.paragraph_format
            code_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            code_format.first_line_indent = Pt(0)
            code_format.line_spacing = 1.0
            code_format.space_before = Pt(4)
            code_format.space_after = Pt(8)
            code_format.widow_control = False
            set_paragraph_shading(code, _CODE_BLOCK_SHADING)
            set_paragraph_borders(
                code,
                left=_CODE_BLOCK_BORDER, top=_CODE_BLOCK_BORDER,
                right=_CODE_BLOCK_BORDER, bottom=_CODE_BLOCK_BORDER,
            )

    def _compat_do_not_expand_shift_return(self) -> None:
        """Stop Word stretching a justified line that ends in a manual break.

        Without ``w:doNotExpandShiftReturn`` Word justifies the line before
        every Shift+Enter across the full width, so a preserved line break
        in a justified paragraph turns its line into widely spaced words.
        """
        settings = self.document.settings.element
        compat = settings.find(qn("w:compat"))
        if compat is None:
            return
        if compat.find(qn("w:doNotExpandShiftReturn")) is not None:
            return
        # CT_Compat is a sequence: the element has to follow the ones the
        # schema puts before it, or a template carrying legacy settings
        # (spaceForUL, ulTrailSpace, …) ends up with an invalid settings.xml.
        earlier = {qn(f"w:{name}") for name in _COMPAT_BEFORE_SHIFT_RETURN}
        position = 0
        for index, child in enumerate(compat):
            if child.tag in earlier:
                position = index + 1
        compat.insert(position, OxmlElement("w:doNotExpandShiftReturn"))

    def _apply_front_matter_properties(self) -> None:
        data = self._front_matter
        if not data:
            return
        properties = self.document.core_properties

        def text(key: str) -> str | None:
            value = data.get(key)
            if isinstance(value, list):
                return ", ".join(value)
            return value or None

        values = {
            "title": text("title"),
            "author": text("author") or text("authors"),
            "subject": text("subject") or text("description"),
            "keywords": text("keywords") or text("tags"),
            "comments": text("abstract") or (text("description") if text("subject") else None),
        }
        for name, value in values.items():
            if not value:
                continue
            value = _XML_INVALID.sub("", value)
            if len(value) > _MAX_PROPERTY_LENGTH:
                self._warn(
                    f"Front matter value for the document's {name} property is longer than "
                    f"{_MAX_PROPERTY_LENGTH} characters; it was shortened",
                    "front_matter_truncated",
                    line=1,
                )
                value = value[: _MAX_PROPERTY_LENGTH - 1] + "…"
            setattr(properties, name, value)
