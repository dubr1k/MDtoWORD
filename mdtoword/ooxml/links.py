"""Heading anchors, bookmarks and hyperlinks (internal and external)."""

from __future__ import annotations

import re
import unicodedata
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from ._xml import _max_int_attribute

_BOOKMARK_NAME_MAX = 40
_BOOKMARK_COUNTER_ATTR = "_mdtoword_next_bookmark_id"
_BOOKMARK_IDS = etree.XPath(
    "//w:bookmarkStart/@w:id",
    namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
)


def github_slug(text: str, used: dict[str, int]) -> str:
    """GitHub's heading anchor for ``text`` (github-slugger algorithm).

    Lowercase; drop everything except letters, marks, digits, ``-``, ``_``
    and spaces (Cyrillic and other scripts stay); each space becomes ``-``.
    Repeats get ``-1``, ``-2``… suffixes, tracked in ``used``.
    """
    base = "".join(_slug_char(char) for char in text.strip().lower())
    slug = base
    while slug in used:
        used[base] += 1
        slug = f"{base}-{used[base]}"
    used[slug] = 0
    return slug


def _slug_char(char: str) -> str:
    if char == " ":
        return "-"
    if char in "-_" or unicodedata.category(char)[0] in "LMN":
        return char
    return ""


def bookmark_name(slug: str, used: set[str]) -> str:
    """A valid, unique (case-insensitively) Word bookmark name for ``slug``.

    Non-word characters become ``_``; a name that would not start with a
    letter gets an ``h_`` prefix (a leading ``_`` marks a *hidden* bookmark in
    Word); the result is cut to 40 characters, with ``_2``, ``_3``… keeping it
    unique. The chosen name is added to ``used``.
    """
    name = re.sub(r"\W", "_", slug)
    if not name or not name[0].isalpha():
        name = f"h_{name}"
    taken = {existing.casefold() for existing in used}
    candidate = name[:_BOOKMARK_NAME_MAX]
    counter = 1
    while candidate.casefold() in taken:
        counter += 1
        suffix = f"_{counter}"
        candidate = name[: _BOOKMARK_NAME_MAX - len(suffix)] + suffix
    used.add(candidate)
    return candidate


def _next_bookmark_id(part: Any) -> int:
    """Allocate a bookmark id unique across every story of the package."""
    package = part.package
    owner = package.main_document_part if package is not None else part
    next_id = getattr(owner, _BOOKMARK_COUNTER_ATTR, None)
    if next_id is None:
        existing: list[str] = []
        parts = package.iter_parts() if package is not None else [part]
        for candidate in parts:
            element = getattr(candidate, "_element", None)
            if element is not None:
                existing.extend(_BOOKMARK_IDS(element))
        next_id = _max_int_attribute(existing, -1) + 1
    setattr(owner, _BOOKMARK_COUNTER_ATTR, next_id + 1)
    return next_id


def add_bookmark(paragraph: Paragraph, name: str) -> int:
    """Wrap the paragraph's current content in bookmark ``name``; return its id.

    Call it after the paragraph's runs are added so the bookmark spans them
    (an empty paragraph yields a collapsed bookmark, which still works as a
    link target).
    """
    if not name or len(name) > _BOOKMARK_NAME_MAX:
        raise ValueError(f"invalid bookmark name {name!r}; use bookmark_name()")
    bookmark_id = str(_next_bookmark_id(paragraph.part))
    start = OxmlElement("w:bookmarkStart", {qn("w:id"): bookmark_id, qn("w:name"): name})
    end = OxmlElement("w:bookmarkEnd", {qn("w:id"): bookmark_id})
    p = paragraph._p
    ppr = p.pPr
    if ppr is not None:
        ppr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)
    return int(bookmark_id)


def add_internal_hyperlink(paragraph: Paragraph, anchor: str) -> Any:
    """Append an empty ``<w:hyperlink w:anchor=…>`` to ``paragraph`` and return it."""
    hyperlink = OxmlElement("w:hyperlink", {qn("w:anchor"): anchor, qn("w:history"): "1"})
    paragraph._p.append(hyperlink)
    return hyperlink


def add_external_hyperlink(paragraph: Paragraph, url: str) -> Any:
    """Append an empty external ``w:hyperlink`` (r:id on ``paragraph.part``)."""
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink", {qn("r:id"): r_id, qn("w:history"): "1"})
    paragraph._p.append(hyperlink)
    return hyperlink


def add_hyperlink_run(hyperlink: Any, paragraph: Paragraph, text: str = "") -> Run:
    """Append a run to a ``w:hyperlink`` element and return it as a python-docx Run."""
    element = OxmlElement("w:r")
    hyperlink.append(element)
    run = Run(element, paragraph)
    if text:
        run.text = text
    return run
