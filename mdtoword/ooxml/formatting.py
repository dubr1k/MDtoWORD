"""Paragraph/run shading, paragraph borders, keep options and table-row flags."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

from ._xml import (
    _PPR_SEQUENCE,
    _RPR_SEQUENCE,
    _TRPR_SEQUENCE,
    _hex_color,
    _insert_ordered,
    _ppr_of,
    _replace_ordered,
    _rpr_of,
)

if TYPE_CHECKING:
    from docx.table import _Row


BorderSpec = tuple[str, int, int, str]


def _shading(fill_hex: str) -> Any:
    return OxmlElement(
        "w:shd", {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): _hex_color(fill_hex)}
    )


def set_paragraph_shading(paragraph_or_style: Any, fill_hex: str) -> None:
    """Solid background ``fill_hex`` behind a paragraph (or every paragraph of a style)."""
    _replace_ordered(_ppr_of(paragraph_or_style), _shading(fill_hex), _PPR_SEQUENCE)


def set_paragraph_borders(
    paragraph_or_style: Any,
    *,
    left: BorderSpec | None = None,
    top: BorderSpec | None = None,
    right: BorderSpec | None = None,
    bottom: BorderSpec | None = None,
    between: BorderSpec | None = None,
) -> None:
    """Replace the paragraph borders; each side is (val, size in 1/8 pt, space pt, color).

    Sides left as ``None`` get no border; all ``None`` removes ``w:pBdr``.
    """
    ppr = _ppr_of(paragraph_or_style)
    borders = OxmlElement("w:pBdr")
    sides = {"top": top, "left": left, "bottom": bottom, "right": right, "between": between}
    for side, spec in sides.items():
        if spec is None:
            continue
        val, size, space, color = spec
        borders.append(OxmlElement(f"w:{side}", {
            qn("w:val"): val,
            qn("w:sz"): str(int(size)),
            qn("w:space"): str(int(space)),
            qn("w:color"): _hex_color(color),
        }))
    if len(borders):
        _replace_ordered(ppr, borders, _PPR_SEQUENCE)
    else:
        for existing in ppr.findall(qn("w:pBdr")):
            ppr.remove(existing)


def set_run_shading(run_or_style: Any, fill_hex: str) -> None:
    """Background ``fill_hex`` behind a run's text (or a character style's)."""
    _replace_ordered(_rpr_of(run_or_style), _shading(fill_hex), _RPR_SEQUENCE)


def _set_row_flag(row: _Row, nsptag: str) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn(nsptag)) is None:
        _insert_ordered(tr_pr, OxmlElement(nsptag), _TRPR_SEQUENCE)


def set_table_header_repeat(row: _Row) -> None:
    """Repeat ``row`` at the top of every page the table spans."""
    _set_row_flag(row, "w:tblHeader")


def set_cant_split(row: _Row) -> None:
    """Keep ``row`` from breaking across pages."""
    _set_row_flag(row, "w:cantSplit")


def set_keep_lines(
    paragraph_or_style: Any, keep_together: bool = True, keep_with_next: bool | None = None
) -> None:
    """Set keep-lines-together (and optionally keep-with-next) on a paragraph or style."""
    paragraph_format = paragraph_or_style.paragraph_format
    paragraph_format.keep_together = keep_together
    if keep_with_next is not None:
        paragraph_format.keep_with_next = keep_with_next
