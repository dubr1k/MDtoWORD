"""GFM pipe tables from Word tables.

Merged cells are split (the text stays in the first cell), column alignment
comes from the header row's paragraphs, cell paragraphs are joined with
``<br>``, and tables nested in a cell are flattened to text.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING, Any

from .code import code_text
from .escaping import _code_span
from .inline import _drop_attr, _Text
from .render import render_inline
from .toc import _TOC_PLACEHOLDER
from .wordml import (
    W_CUSTOMXML,
    W_GRIDAFTER,
    W_GRIDBEFORE,
    W_GRIDSPAN,
    W_HMERGE,
    W_JC,
    W_P,
    W_PPR,
    W_SDT,
    W_SDTCONTENT,
    W_TBL,
    W_TC,
    W_TCPR,
    W_TR,
    W_TRPR,
    W_VAL,
    W_VMERGE,
    _int,
    _q,
)

if TYPE_CHECKING:
    from .document import _Conversion


class _TableWriter:
    """Write the tables of one story (the body, or a note) as GFM tables."""

    def __init__(self, conversion: _Conversion, part: Any) -> None:
        self.c = conversion
        self.part = part

    def table(self, element: Any) -> str:
        rows: list[list[str]] = []
        header_alignment: list[str | None] = []
        merged = False
        for row_index, row in enumerate(self._rows(element)):
            cells: list[str] = []
            alignment: list[str | None] = []
            row_properties = row.find(W_TRPR)
            before = after = 0
            if row_properties is not None:
                element_before = row_properties.find(W_GRIDBEFORE)
                element_after = row_properties.find(W_GRIDAFTER)
                before = _int(element_before.get(W_VAL)) if element_before is not None else 0
                after = _int(element_after.get(W_VAL)) if element_after is not None else 0
            cells.extend([""] * before)
            alignment.extend([None] * before)
            for cell in self._cells(row):
                properties = cell.find(W_TCPR)
                span = 1
                continued = False
                if properties is not None:
                    span_element = properties.find(W_GRIDSPAN)
                    span = max(1, _int(span_element.get(W_VAL), 1)) if span_element is not None else 1
                    vertical = properties.find(W_VMERGE)
                    horizontal = properties.find(W_HMERGE)
                    if vertical is not None and vertical.get(W_VAL) != "restart":
                        continued = True
                    if vertical is not None or horizontal is not None or span > 1:
                        merged = True
                    if horizontal is not None and horizontal.get(W_VAL) != "restart":
                        continued = True
                text = "" if continued else self.cell(cell, header=row_index == 0)
                cells.extend([text] + [""] * (span - 1))
                justification = self._cell_alignment(cell)
                alignment.extend([justification] * span)
            cells.extend([""] * after)
            alignment.extend([None] * after)
            rows.append(cells)
            if row_index == 0:
                header_alignment = alignment
        if not rows:
            return ""
        if merged:
            self.c.diagnostics.warn("table_merged_cells")
        width = max(len(row) for row in rows)
        if width == 0:
            return ""
        rows = [row + [""] * (width - len(row)) for row in rows]
        header_alignment += [None] * (width - len(header_alignment))
        separators = []
        for justification in header_alignment[:width]:
            if justification == "center":
                separators.append(":---:")
            elif justification in ("right", "end"):
                separators.append("---:")
            else:
                separators.append("---")
        lines = ["| " + " | ".join(rows[0]) + " |",
                 "| " + " | ".join(separators) + " |"]
        lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
        return "\n".join(lines)

    @staticmethod
    def indent(table: Any) -> int:
        """The table's left indent in twips, for placing it inside a list item."""
        properties = table.find(_q("w:tblPr"))
        indent = properties.find(_q("w:tblInd")) if properties is not None else None
        if indent is None or indent.get(_q("w:type"), "dxa") != "dxa":
            return 0
        return _int(indent.get(_q("w:w")))

    def _rows(self, table: Any) -> Iterable[Any]:
        for child in table:
            if child.tag == W_TR:
                yield child
            elif child.tag in (W_SDT, W_CUSTOMXML):
                content = child.find(W_SDTCONTENT) if child.tag == W_SDT else child
                if content is not None:
                    yield from self._rows(content)

    def _cells(self, row: Any) -> Iterable[Any]:
        for child in row:
            if child.tag == W_TC:
                yield child
            elif child.tag in (W_SDT, W_CUSTOMXML):
                content = child.find(W_SDTCONTENT) if child.tag == W_SDT else child
                if content is not None:
                    yield from self._cells(content)

    def _cell_alignment(self, cell: Any) -> str | None:
        paragraph = next(cell.iter(W_P), None)
        if paragraph is None:
            return None
        ppr = paragraph.find(W_PPR)
        jc = ppr.find(W_JC) if ppr is not None else None
        if jc is None:
            for style in self.c.classifier.paragraph_info(paragraph).chain:
                if style.ppr is not None and style.ppr.find(W_JC) is not None:
                    jc = style.ppr.find(W_JC)
                    break
        return jc.get(W_VAL) if jc is not None else None

    def cell(self, cell: Any, *, header: bool = False) -> str:
        pieces: list[str] = []
        inline = self.c.inline
        for element in self.c.block_elements(cell):
            if element is _TOC_PLACEHOLDER:
                # A table-of-contents control cannot be placed as a whole
                # inside a cell -- reported like a TOC field there.
                self.c.diagnostics.warn("toc_skipped")
                continue
            if element.tag == W_TBL:
                self.c.diagnostics.warn("nested_table_flattened")
                flattened = self._flatten(element)
                if flattened:
                    pieces.append(flattened)
                continue
            info = self.c.classifier.paragraph_info(element)
            if info.code:
                lines = [line for line in code_text(element).split("\n") if line.strip()]
                pieces.extend(_code_span(line.replace("\t", "    ")).replace("|", "\\|")
                              for line in lines)
                continue
            tables_of_contents = inline.toc_count
            items = inline.collect(element, self.part)
            if inline.toc_count > tables_of_contents:
                self.c.diagnostics.warn("toc_skipped")
            if header:
                # Header cells render bold anyway; bold on all of one is style.
                texts = [item for item in items if isinstance(item, _Text) and item.text.strip()]
                if texts and all("bold" in item.attrs for item in texts):
                    items = _drop_attr(items, "bold")
            text = render_inline(items, "cell")
            if info.level is not None:
                self.c.numbering.advance(info.level)
                label = (self.c.numbering.label(info.level) or "1.") if info.level.ordered else "•"
                text = f"{label} {text}" if text else ""
            if text:
                pieces.append(text)
        return "<br>".join(pieces)

    def _flatten(self, table: Any) -> str:
        rows = []
        for row in self._rows(table):
            cells = [self.cell(cell) for cell in self._cells(row)]
            cells = [cell for cell in cells if cell]
            if cells:
                rows.append("; ".join(cells))
        return "<br>".join(rows)
