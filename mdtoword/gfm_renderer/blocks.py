"""Block tokens: the type -> handler dispatch, and code blocks and rules.

``render`` hands every token of the top-level stream to ``_render_block``,
which looks the token type up in ``_BLOCK_HANDLERS`` and calls the handler
named there. Every handler takes ``(tokens, index, source_path)`` and
returns the index of the next token to render, so one that consumes
several tokens at once (a figure, a TOC marker, a table caption) skips
past them. Handlers open and close container state -- headings, lists,
blockquotes, definition lists -- or hand the token to the mixin that owns
the construct (tables, footnotes, math, raw HTML, special paragraphs).
Fenced/indented code blocks and thematic breaks need nothing more and are
rendered here.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from ..ooxml import add_bookmark
from .helpers import _ListState
from .state import RendererState

_TABLE_TOKENS = (
    "table_open", "thead_open", "thead_close", "tbody_open", "tbody_close", "tr_open",
    "tr_close", "th_open", "td_open", "th_close", "td_close", "table_close",
)

# Token type -> name of the method that renders it. Looked up by name on the
# instance, so a handler may live in any mixin. Types not listed here --
# dl_open/dl_close, footnote_block_close, footnote_anchor (the HTML
# back-reference) and anything unknown -- carry no content of their own and
# are skipped.
_BLOCK_HANDLERS: dict[str, str] = {
    "front_matter": "_block_front_matter",
    "heading_open": "_block_heading_open",
    "heading_close": "_block_heading_close",
    "paragraph_open": "_block_paragraph_open",
    "inline": "_block_inline",
    "paragraph_close": "_block_paragraph_close",
    "blockquote_open": "_block_blockquote_open",
    "blockquote_close": "_block_blockquote_close",
    "bullet_list_open": "_block_list_open",
    "ordered_list_open": "_block_list_open",
    "bullet_list_close": "_block_list_close",
    "ordered_list_close": "_block_list_close",
    "list_item_open": "_block_list_item_open",
    "list_item_close": "_block_list_item_close",
    "amsmath": "_block_amsmath",
    "math_block": "_block_math",
    "math_block_label": "_block_math_label",
    "fence": "_block_code",
    "code_block": "_block_code",
    "hr": "_block_rule",
    "html_block": "_block_html",
    **dict.fromkeys(_TABLE_TOKENS, "_block_table"),
    "dt_open": "_block_term_open",
    "dt_close": "_block_term_close",
    "dd_open": "_block_description_open",
    "dd_close": "_block_description_close",
    "footnote_block_open": "_block_footnote_block_open",
    "footnote_open": "_block_footnote_open",
    "footnote_close": "_block_footnote_close",
}


class BlockMixin(RendererState):
    """Dispatch block tokens to their handlers; render code blocks and rules."""

    def _render_block(self, tokens: Sequence[Any], index: int, source_path: Path | None) -> int:
        token_type = tokens[index].type

        if self._skip_footnote and token_type != "footnote_close":
            return index + 1

        handler = _BLOCK_HANDLERS.get(token_type)
        if handler is None:
            # footnote_anchor (the HTML back-reference) and anything unknown that
            # carries no content of its own.
            return index + 1
        return getattr(self, handler)(tokens, index, source_path)

    # ---- front matter, headings, paragraphs

    def _block_front_matter(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._render_title_block()
        if self.options.toc:
            self._insert_toc()
        return index + 1

    def _block_heading_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        token = tokens[index]
        if self._lists:
            self._ensure_item_started()
        level = min(int(token.tag[1:]), 9)
        self._heading_level = level
        self._current_heading_bookmark = self._heading_bookmark_at.get(index)
        self._paragraph = self._add_paragraph(f"Heading {level}")
        if self._lists or self._quotes:
            self._decorate(self._paragraph, "heading")
        return index + 1

    def _block_heading_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        if self._paragraph is not None and self._current_heading_bookmark:
            add_bookmark(self._paragraph, self._current_heading_bookmark)
        self._current_heading_bookmark = None
        self._paragraph = None
        self._heading_level = None
        return index + 1

    def _block_paragraph_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        special = self._render_special_paragraph(tokens, index, source_path)
        if special is not None:
            return special
        self._paragraph = self._new_paragraph()
        return index + 1

    def _block_inline(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        token = tokens[index]
        self._render_inline(token.children or [], source_path, token.content)
        return index + 1

    def _block_paragraph_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._paragraph = None
        return index + 1

    # ---- containers: blockquotes, lists, definition lists

    def _block_blockquote_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._open_blockquote(tokens, index)
        return index + 1

    def _block_blockquote_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        if self._quotes:
            self._quotes.pop()
        return index + 1

    def _block_list_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        token = tokens[index]
        self._ensure_item_started()
        ordered = token.type == "ordered_list_open"
        start = 1
        if ordered:
            try:
                start = int(token.attrGet("start") or 1)
            except ValueError:
                start = 1
        self._lists.append(
            _ListState(
                ordered=ordered,
                num_id=self._numbering.start_list(ordered, start),
                level=len(self._lists),
            )
        )
        return index + 1

    def _block_list_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        if self._lists:
            self._lists.pop()
        return index + 1

    def _block_list_item_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        if self._lists:
            self._lists[-1].has_paragraph = False
        return index + 1

    def _block_list_item_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._ensure_item_started()
        return index + 1

    def _block_term_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._in_term = True
        return index + 1

    def _block_term_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._in_term = False
        if self._paragraph is not None:
            for run in self._paragraph.runs:
                run.bold = True
        self._paragraph = None
        return index + 1

    def _block_description_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._definition_depth += 1
        return index + 1

    def _block_description_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._definition_depth -= 1
        return index + 1

    # ---- leaf blocks owned by other mixins: math, code, rules, HTML, tables

    def _block_amsmath(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._render_amsmath(tokens[index])
        return index + 1

    def _block_math(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._render_math(tokens[index].content, display=True)
        return index + 1

    def _block_math_label(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        token = tokens[index]
        self._render_math(token.content, display=True, label=token.info)
        return index + 1

    def _block_code(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._render_code_block(tokens[index])
        return index + 1

    def _block_rule(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._ensure_item_started()
        self._add_thematic_break()
        return index + 1

    def _block_html(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._render_html_block(tokens[index].content, source_path)
        return index + 1

    def _block_table(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._table_token(tokens[index], source_path)
        return index + 1

    # ---- footnotes

    def _block_footnote_block_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        if self._footnotes is None:
            heading = self._add_paragraph("Heading 2")
            heading.add_run(self._localized_footnotes_heading())
        return index + 1

    def _block_footnote_open(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._open_footnote(tokens[index])
        return index + 1

    def _block_footnote_close(
        self, tokens: Sequence[Any], index: int, source_path: Path | None
    ) -> int:
        self._close_footnote()
        return index + 1

    # ---- code blocks and thematic breaks

    def _render_code_block(self, token: Any) -> None:
        self._ensure_item_started()
        code_style_id = self._style("Source Code").style_id
        if self._footnote_target is None and self._last_body_style_id() == code_style_id:
            # Word merges the borders of adjacent paragraphs that share the
            # same border settings, so two code blocks in a row would read as
            # one. A small plain paragraph between them keeps them apart.
            spacer = self._add_paragraph()
            spacer.paragraph_format.space_after = Pt(0)
            spacer.paragraph_format.space_before = Pt(0)
            spacer.paragraph_format.line_spacing = Pt(4)
        paragraph = self._add_paragraph("Source Code")
        self._decorate(paragraph, "code")
        paragraph.add_run(token.content.rstrip("\n"))

    def _last_body_style_id(self) -> str | None:
        body = self.document.element.body
        children = [child for child in body if child.tag != qn("w:sectPr")]
        if not children or children[-1].tag != qn("w:p"):
            return None
        properties = children[-1].pPr
        if properties is None or properties.pStyle is None:
            return None
        return properties.pStyle.val

    def _add_thematic_break(self) -> None:
        paragraph = self._add_paragraph()
        self._decorate(paragraph, "rule")
        properties = paragraph._p.get_or_add_pPr()
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "808080")
        borders.append(bottom)
        existing = properties.find(qn("w:pBdr"))
        if existing is not None:
            existing.append(bottom)
        else:
            properties.insert_element_before(
                borders,
                "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
                "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                "w:pPrChange",
            )
