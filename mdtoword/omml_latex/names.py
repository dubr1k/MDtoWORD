"""OMML and WordprocessingML tag names, and the OOXML helpers shared by the converter."""

from __future__ import annotations

from typing import Any

_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _m(local: str) -> str:
    return f"{{{_M}}}{local}"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


_VAL = _m("val")

# Elements that carry no content of their own.
_SKIPPED_W = frozenset(
    _w(name) for name in (
        "rPr", "bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart",
        "commentRangeEnd", "permStart", "permEnd", "del", "moveFrom",
        "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart",
        "moveToRangeEnd", "lastRenderedPageBreak",
    )
)
_TRANSPARENT_W = frozenset(_w(name) for name in ("ins", "moveTo", "smartTag", "customXml"))


def _is_on(element: Any) -> bool:
    """OOXML on/off property: present without a value, or a true value."""
    if element is None:
        return False
    value = element.get(_VAL)
    return value is None or value.lower() in ("1", "on", "true")


def _local(tag: Any) -> str:
    return tag.split("}", 1)[-1] if isinstance(tag, str) else ""
