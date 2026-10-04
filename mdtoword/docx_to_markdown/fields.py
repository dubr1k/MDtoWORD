"""Word fields: what an instruction such as ``HYPERLINK``, ``REF`` or ``TOC`` becomes.

A field's result is either kept as text, suppressed (page furniture,
generated listings), replaced by a ``[TOC]`` marker (a heading table of
contents) or turned into a link (``HYPERLINK``, ``REF \\h``).
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field

from .inline import _Link

# Fields whose result is generated page furniture, not document text.
_DROPPED_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION",
                             "PAGEREF", "XE", "TC", "TA", "RD"})
# Generated listings with no Markdown form. A heading TOC is not among them:
# it becomes a "[TOC]" marker that Markdown tools expand themselves.
_GENERATED_FIELDS = {"INDEX": "index_skipped", "TOA": "index_skipped"}
# TOC switches that make it a table of figures/tables, not of headings.
_CAPTION_TOC_SWITCHES = frozenset({"\\c", "\\a"})

_FIELD_TOKEN = re.compile(r'"([^"]*)"|(\S+)')


@dataclass
class _Field:
    """An open field: its instruction so far, phase and what its result becomes."""

    instr: list[str] = field(default_factory=list)
    phase: str = "instr"
    suppress: bool = False
    link: _Link | None = None


@dataclass
class _FieldMeaning:
    """How a field's result is written, from its instruction."""

    suppress: bool = False
    warning: str | None = None
    table_of_contents: bool = False
    target: str | None = None
    title: str | None = None


def interpret_field(instr: list[str], anchor_target: Callable[[str], str]) -> _FieldMeaning:
    """Read a field instruction; ``anchor_target`` maps a bookmark to its anchor."""
    tokens = [quoted or bare for quoted, bare in _FIELD_TOKEN.findall(" ".join(instr))]
    kind = tokens[0].upper() if tokens else ""
    if kind == "TOC":
        if any(token.lower() in _CAPTION_TOC_SWITCHES for token in tokens[1:]):
            return _FieldMeaning(suppress=True, warning="toc_skipped")
        # The block writer replaces it with a "[TOC]" marker.
        return _FieldMeaning(suppress=True, table_of_contents=True)
    if kind in _GENERATED_FIELDS:
        return _FieldMeaning(suppress=True, warning=_GENERATED_FIELDS[kind])
    if kind in _DROPPED_FIELDS:
        return _FieldMeaning(suppress=True)
    target = title = None
    if kind == "HYPERLINK":
        url, anchor = "", None
        index = 1
        while index < len(tokens):
            token = tokens[index]
            if token.lower() in ("\\l", "\\o", "\\t", "\\m", "\\n") and index + 1 < len(tokens):
                if token.lower() == "\\l":
                    anchor = tokens[index + 1]
                elif token.lower() == "\\o":
                    title = tokens[index + 1]
                index += 2
                continue
            if not token.startswith("\\") and not url:
                url = token
            index += 1
        target = url + (f"#{anchor_target(anchor)}" if anchor else "")
    elif kind == "REF" and len(tokens) > 1 and any(t.lower() == "\\h" for t in tokens):
        target = "#" + anchor_target(tokens[1])
    return _FieldMeaning(target=target, title=title)
