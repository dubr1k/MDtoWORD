"""GFM tables: collecting the rows, then building the Word table.

Table tokens are buffered cell by cell (inline content is captured, not
rendered, while a row is open) because column widths depend on every row.
At ``table_close`` the Word table is built with direct borders, the
container indent, content-weighted column widths and repeating header
rows, and each cell's inline content is rendered into it. A table inside a
footnote, which Word cannot hold, is flattened to text.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Length, Pt
from docx.table import _Cell as _DocxCell

from ..ooxml import set_table_header_repeat
from .constants import _MAX_TABLE_COLUMNS
from .helpers import _Cell, _inline_plain_text
from .state import RendererState


class TableMixin(RendererState):
    """Buffer table tokens and render the finished table."""

    # Class-level defaults, so both exist before the first table sets them on
    # the instance (they are not part of _reset).
    _table_cell_children: list[Any] | None = None
    _table_columns: int = 0

    def _table_token(self, token: Any, source_path: Path | None) -> None:
        token_type = token.type
        if token_type == "table_open":
            self._ensure_item_started()
            self._table_rows = []
            self._table_line = self._line
            return
        if token_type == "thead_open":
            self._table_header = True
            return
        if token_type == "thead_close":
            self._table_header = False
            return
        if token_type == "tr_open":
            self._table_row = []
            return
        if token_type in {"th_open", "td_open"}:
            style_attr = token.attrGet("style") or ""
            alignment = (
                "right" if "right" in style_attr else "center" if "center" in style_attr else None
            )
            self._table_row_alignment = alignment
            self._table_cell_children = []
            return
        if token_type in {"th_close", "td_close"}:
            if self._table_row is not None:
                children = self._table_cell_children or []
                self._table_row.append(
                    _Cell(
                        children=children,
                        content="",
                        alignment=self._table_row_alignment,
                        header=self._table_header,
                        line=self._line,
                    )
                )
            self._table_cell_children = None
            return
        if token_type == "tr_close":
            if self._table_rows is not None and self._table_row is not None:
                self._table_rows.append(self._table_row)
            self._table_row = None
            return
        if token_type == "table_close":
            self._finish_table(source_path)

    def _finish_table(self, source_path: Path | None) -> None:
        rows = self._table_rows or []
        self._table_rows = None
        if not rows:
            return
        if self._pending_footnote_label is not None:
            # A section-mode note that starts with a table: its mark gets a
            # line of its own instead of drifting to the next paragraph.
            self._add_paragraph()
        columns = max(len(row) for row in rows)
        if self._footnote_target is not None:
            self._warn("Table inside a footnote flattened to text", "table_in_footnote")
            for row in rows:
                paragraph = self._new_paragraph()
                paragraph.add_run(" | ".join(_inline_plain_text(cell.children) for cell in row))
            return

        table = self.document.add_table(rows=len(rows), cols=columns)
        table.style = self.document.styles["Table Grid"] if self._has_style("Table Grid") else None
        self._apply_table_borders(table)
        indent = int(self._container_indent())
        if indent:
            self._set_table_indent(table, indent)

        if columns > _MAX_TABLE_COLUMNS:
            self._warn(
                f"Table has {columns} columns; Word supports at most {_MAX_TABLE_COLUMNS}, "
                "so it may not display every column",
                "table_too_wide",
                line=self._table_line,
            )
        # python-docx's table.cell()/row.cells rebuild the whole cell grid on
        # every call -- quadratic in the table size. A freshly created table
        # has exactly one w:tc per grid cell, so the grid is read once here.
        grid = [[_DocxCell(tc, table) for tc in tr.tc_lst] for tr in table._tbl.tr_lst]
        widths = self._column_widths(rows, columns)
        table.autofit = False
        for column_index, width in enumerate(widths):
            table.columns[column_index].width = width
        for row_cells in grid:
            for column_index, width in enumerate(widths):
                row_cells[column_index].width = width
        header_rows = sum(1 for row in rows if row and row[0].header)
        for row_index in range(min(header_rows, len(table.rows))):
            set_table_header_repeat(table.rows[row_index])

        alignments = {"right": WD_ALIGN_PARAGRAPH.RIGHT, "center": WD_ALIGN_PARAGRAPH.CENTER}
        saved_lists, saved_quotes, saved_depth = self._lists, self._quotes, self._definition_depth
        self._lists, self._quotes, self._definition_depth = [], [], 0
        self._cell_context = True
        self._table_columns = columns
        try:
            for row_index, cells in enumerate(rows):
                for column_index in range(columns):
                    paragraph = grid[row_index][column_index].paragraphs[0]
                    paragraph_format = paragraph.paragraph_format
                    paragraph_format.space_after = Pt(0)
                    paragraph_format.space_before = Pt(0)
                    if self._gost:
                        paragraph_format.first_line_indent = Pt(0)
                        paragraph_format.line_spacing = 1.0
                    if column_index >= len(cells):
                        continue
                    cell = cells[column_index]
                    self._line = cell.line
                    paragraph.alignment = alignments.get(cell.alignment or "", WD_ALIGN_PARAGRAPH.LEFT)
                    self._paragraph = paragraph
                    self._table_header = cell.header
                    self._render_inline(cell.children, source_path, "")
                    self._paragraph = None
        finally:
            self._cell_context = False
            self._table_header = False
            self._table_columns = 0
            self._lists, self._quotes, self._definition_depth = saved_lists, saved_quotes, saved_depth
        self._last_table = table

    def _has_style(self, name: str) -> bool:
        try:
            self.document.styles[name]
        except KeyError:
            return False
        return True

    def _column_widths(self, rows: list[list[_Cell]], columns: int) -> list[Length]:
        weights = []
        for column_index in range(columns):
            longest = 3
            for row in rows:
                if column_index < len(row):
                    text = _inline_plain_text(row[column_index].children)
                    longest = max(longest, max((len(word) for word in text.split()), default=0) * 2, min(len(text), 60))
            weights.append(min(longest, 60))
        total_width = int(self._available_width())
        total_weight = sum(weights) or 1
        # Every column gets a floor (1.5 cm, or an equal share when even
        # that does not fit), then the rest is shared by content length --
        # so the sum never exceeds the text width.
        minimum = min(int(Cm(1.5)), total_width // max(1, columns))
        spare = max(0, total_width - minimum * columns)
        widths = [minimum + spare * weight // total_weight for weight in weights]
        # Word stores widths in twips (635 EMU); rounding down keeps the
        # stored sum inside the text width too.
        return [Emu(width - width % 635) for width in widths]

    @staticmethod
    def _set_table_indent(table: Any, indent_emu: int) -> None:
        tbl_pr = table._tbl.tblPr
        indentation = OxmlElement("w:tblInd")
        indentation.set(qn("w:w"), str(int(indent_emu / 635)))
        indentation.set(qn("w:type"), "dxa")
        tbl_pr.insert_element_before(
            indentation,
            "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
            "w:tblCaption", "w:tblDescription", "w:tblPrChange",
        )

    @staticmethod
    def _apply_table_borders(table: Any) -> None:
        """Write borders as direct formatting so every viewer renders them."""
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element = OxmlElement(f"w:{edge}")
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "4")
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), "000000")
            borders.append(element)
        table._tbl.tblPr.insert_element_before(
            borders,
            "w:shd",
            "w:tblLayout",
            "w:tblCellMar",
            "w:tblLook",
            "w:tblCaption",
            "w:tblDescription",
            "w:tblPrChange",
        )
