"""Tables of contents: Word's TOC control or field becomes a ``[TOC]`` marker.

Markdown tools (the forward converter included) build the listing, and its
title, from the marker themselves.
"""

from __future__ import annotations

import re
from typing import Any

from .wordml import W_DOCPARTGALLERY, W_DOCPARTOBJ, W_SDTPR, W_VAL

_TOC_TITLES = frozenset({"contents", "table of contents", "содержание", "оглавление"})
_TOC_MARKER = "[TOC]"


class _Placeholder:
    """Stands in the element list for a Word table-of-contents control."""

    tag = "toc"

    def __iter__(self) -> Any:
        return iter(())


_TOC_PLACEHOLDER = _Placeholder()


def is_toc_control(sdt: Any) -> bool:
    """Whether a content control is a table of contents (docPart gallery)."""
    properties = sdt.find(W_SDTPR)
    gallery = None
    if properties is not None:
        part_object = properties.find(W_DOCPARTOBJ)
        if part_object is not None:
            gallery = part_object.find(W_DOCPARTGALLERY)
    return gallery is not None and "table of contents" in (gallery.get(W_VAL) or "").lower()


def append_toc_marker(blocks: list[str]) -> None:
    """Append the "[TOC]" marker, replacing a "Contents" title just before it."""
    if blocks:
        title = re.sub(r"[*#_\s]+", " ", blocks[-1]).strip().lower()
        if title in _TOC_TITLES:
            blocks.pop()
    blocks.append(_TOC_MARKER)
