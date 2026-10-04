"""The block writer: one story's paragraphs and tables as Markdown blocks.

Each paragraph is classified -- heading, list item, code, quote, display
equation, thematic break -- and written by a small state machine that
groups consecutive code lines, quote paragraphs and list items into one
Markdown block each. Lists, quotes, tables and captions are delegated to
their own modules; this one holds the grouping state and the dispatch.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .captions import _CaptionMatcher
from .code import code_text, fenced_code, is_language_caption
from .equations import display_math
from .inline import _add_attr, _drop_attr, _Text
from .lists import _ListBuilder, split_task
from .quotes import _QuoteWriter, is_quoted
from .render import render_inline
from .tables import _TableWriter
from .toc import _TOC_PLACEHOLDER, append_toc_marker
from .wordml import W_P, W_TBL, _plain_text

if TYPE_CHECKING:
    from .classify import _ParaInfo
    from .document import _Conversion


class _BlockWriter:
    """Turn a sequence of paragraphs and tables into Markdown blocks.

    Holds the grouping state -- the open list, a pending code block, a
    pending block quote -- for one story (the body, or one footnote).
    """

    def __init__(self, conversion: _Conversion, part: Any, *, inside_quote: bool = False) -> None:
        self.c = conversion
        self.part = part
        self.inside_quote = inside_quote
        self.classifier = conversion.classifier
        self.blocks: list[str] = []
        self.lists = _ListBuilder()
        self.code_lines: list[str] | None = None
        self.code_language = ""
        self.code_indent = 0
        self.pending_language = ""
        self.captions = _CaptionMatcher(conversion.classifier)
        self.tables = _TableWriter(conversion, part)

    def render(self, elements: list[Any]) -> list[str]:
        self.feed(elements)
        self.flush()
        return self.blocks

    def feed(self, elements: list[Any]) -> None:
        """Write a run of elements, keeping open lists and code for the next run."""
        self.captions.match(elements)
        index = 0
        while index < len(elements):
            element = elements[index]
            index += 1
            if element in self.captions.skipped:
                continue
            if element is _TOC_PLACEHOLDER:
                self._table_of_contents()
                continue
            if element.tag == W_TBL:
                self.flush_code()
                table = self.tables.table(element)
                if table:
                    caption = self.captions.tables.get(element)
                    if caption is not None:
                        position, line = caption
                        table = f"{line}\n\n{table}" if position == "before" else f"{table}\n\n{line}"
                    self.place(table, self.tables.indent(element))
                continue
            info = self.classifier.paragraph_info(element)
            if not self.inside_quote and is_quoted(info):
                run = [element]
                while index < len(elements) and elements[index].tag == W_P:
                    following_info = self.classifier.paragraph_info(elements[index])
                    if (elements[index] in self.captions.skipped or not is_quoted(following_info)
                            or following_info.callout != info.callout):
                        break
                    run.append(elements[index])
                    index += 1
                self.flush()
                block = _QuoteWriter(self.classifier, self._nested_writer).render(run)
                if block:
                    self.blocks.append(block)
                continue
            following = elements[index] if index < len(elements) else None
            self.c.pictures.caption = self.captions.figures.get(element)
            try:
                self.paragraph(element, following)
            finally:
                self.c.pictures.caption = None

    def _nested_writer(self) -> _BlockWriter:
        """A writer for one level of a block quote in this story."""
        return _BlockWriter(self.c, self.part, inside_quote=True)

    # -- flushing ---------------------------------------------------------------

    def flush(self) -> None:
        self.flush_code()
        self.close_list()

    def close_list(self) -> None:
        block = self.lists.close()
        if block is not None:
            self.blocks.append(block)

    def flush_code(self) -> None:
        if self.code_lines is None:
            return
        block = fenced_code(self.code_lines, self.code_language)
        indent = self.code_indent
        self.code_lines = None
        self.code_language = ""
        self.place(block, indent)

    def place(self, block: str, indent: int) -> None:
        """Add a block, inside the open list item it is indented under if any."""
        if not self.lists.attach(block, indent):
            self.close_list()
            self.blocks.append(block)

    # -- paragraphs -------------------------------------------------------------

    def paragraph(self, element: Any, following: Any) -> None:
        info = self.classifier.paragraph_info(element)
        if info.code:
            if self.code_lines is None:
                self.code_lines = []
                self.code_language = self.pending_language
                self.code_indent = info.indent
            self.pending_language = ""
            self.code_lines.extend(code_text(element).split("\n"))
            return
        self.flush_code()
        self.pending_language = ""

        if (following is not None and following.tag == W_P
                and is_language_caption(self.classifier, element, info)
                and self.classifier.paragraph_info(following).code):
            self.flush()
            self.pending_language = _plain_text(element).strip()
            return

        if info.rule:
            self.flush()
            self.blocks.append("---")
            return
        if info.toc_heading:
            return

        display = display_math(self.c.diagnostics, element, info)
        if display is not None:
            for block in display:
                self.place(block, info.indent)
            return

        inline = self.c.inline
        was_suppressed = inline.suppressed()
        tables_of_contents = inline.toc_count
        items = inline.collect(element, self.part)
        if inline.toc_count > tables_of_contents:
            # Whatever precedes the field in its paragraph stays; the
            # listing itself becomes the marker.
            text = render_inline(items, "block")
            if text:
                self.flush()
                self.blocks.append(text)
            self._table_of_contents()
            return
        if was_suppressed and not items:
            return

        if info.heading:
            self.heading(info, items)
            return
        if info.level is not None:
            self.list_item(info, items)
            return
        if info.subtitle:
            items = _add_attr(items, "italic")
        text = render_inline(items, "block")
        if not text:
            # An empty paragraph is spacing, not content: it does not
            # interrupt a list Word keeps numbering across.
            return
        self.place(text, info.indent)

    def heading(self, info: _ParaInfo, items: list[Any]) -> None:
        self.flush()
        numbering = self.c.numbering
        if info.level is not None:
            numbering.advance(info.level)
            label = numbering.label(info.level) if info.level.ordered else ""
            if label:
                items = [_Text(label + " ", frozenset()), *items]
        texts = [item for item in items if isinstance(item, _Text) and item.text.strip()]
        for attr in ("bold", "italic"):
            # Formatting on every word of a heading is its style, not emphasis.
            if texts and all(attr in item.attrs for item in texts):
                items = _drop_attr(items, attr)
        if info.heading > 6:
            self.c.diagnostics.warn("heading_level_clamped")
            text = render_inline(_add_attr(items, "bold"), "block")
            if text:
                self.blocks.append(text)
            return
        text = render_inline(items, "heading")
        if text:
            self.blocks.append("#" * info.heading + " " + text)

    def list_item(self, info: _ParaInfo, items: list[Any]) -> None:
        self.flush_code()
        level = info.level
        assert level is not None
        value = self.c.numbering.advance(level)
        task, items = split_task(items)
        text = render_inline(items, "block")
        if not text and not task:
            return
        finished = self.lists.add(level, value, task + text, info.indent)
        if finished is not None:
            self.blocks.append(finished)

    def _table_of_contents(self) -> None:
        """Write the "[TOC]" marker, replacing a "Contents" title just before it."""
        self.flush()
        append_toc_marker(self.blocks)
