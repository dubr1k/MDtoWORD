"""Style lookup in ``w:styles`` by id or by (localized template-safe) name."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from docx.oxml.ns import qn


def _iter_styles(styles_element: Any, style_type: str) -> Iterable[Any]:
    for style in styles_element.iterchildren(qn("w:style")):
        if style.get(qn("w:type")) == style_type:
            yield style


def _find_style_id(styles_element: Any, style_type: str, name_or_id: str) -> str | None:
    """Style id of the ``style_type`` style whose id or (case-insensitive) name matches.

    Localized Word templates keep the English built-in *names* ("footnote
    text") but use opaque ids ("a5"), so a name match is as good as an id one.
    """
    folded = name_or_id.casefold()
    by_name = None
    for style in _iter_styles(styles_element, style_type):
        style_id = style.get(qn("w:styleId"))
        if style_id == name_or_id:
            return style_id
        name = style.find(qn("w:name"))
        if by_name is None and name is not None and (name.get(qn("w:val")) or "").casefold() == folded:
            by_name = style_id
    return by_name


def _default_style_id(styles_element: Any, style_type: str) -> str | None:
    for style in _iter_styles(styles_element, style_type):
        if style.get(qn("w:default")) in ("1", "true", "on"):
            return style.get(qn("w:styleId"))
    return None


def _styles_element_of(part: Any) -> Any:
    """``w:styles`` root of the package that ``part`` belongs to."""
    return part.package.main_document_part.styles.element


def _unused_style_id(styles_element: Any, preferred: str) -> str:
    taken = set(styles_element.xpath("./w:style/@w:styleId"))
    candidate, counter = preferred, 1
    while candidate in taken:
        counter += 1
        candidate = f"{preferred}{counter}"
    return candidate
