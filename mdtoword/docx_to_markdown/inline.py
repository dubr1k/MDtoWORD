"""The inline content model: formatted text segments, links, raw Markdown and spans.

A paragraph's runs become a flat list of formatted text segments
(:class:`_Text`) plus atomic items -- links, and :class:`_Raw` Markdown for
images, math and footnote references. Adjacent segments with identical
formatting are merged, and the list is folded into a tree of emphasis
spans (:class:`_Span`), longest span outermost, for :mod:`.render` to print.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any

# Inline formatting attributes, in the order a tie is nested (outermost first).
_ATTRS = ("bold", "italic", "strike", "underline", "highlight", "sup", "sub")


@dataclass
class _Text:
    text: str
    attrs: frozenset[str]
    code: bool = False


@dataclass
class _Raw:
    markdown: str
    plain: str = ""


@dataclass
class _Link:
    url: str
    title: str | None
    children: list[Any]
    attrs: frozenset[str] = frozenset()


@dataclass
class _Span:
    attr: str
    children: list[Any]


def _item_attrs(item: Any) -> frozenset[str]:
    if isinstance(item, (_Text, _Link)):
        return item.attrs
    return frozenset()


def _plain(items: Iterable[Any]) -> str:
    parts = []
    for item in items:
        if isinstance(item, _Text):
            parts.append(item.text)
        elif isinstance(item, _Link):
            parts.append(_plain(item.children))
        elif isinstance(item, _Raw):
            parts.append(item.plain)
    return "".join(parts)


def _prepare(items: list[Any]) -> list[Any]:
    """Normalise segments: tabs, breaks inside code, merging, link attrs."""
    prepared: list[Any] = []
    for item in items:
        if isinstance(item, _Link):
            children = _prepare(item.children)
            texts = [child.attrs for child in children
                     if isinstance(child, _Text) and child.text.strip()]
            attrs = frozenset.intersection(*texts) if texts else frozenset()
            prepared.append(_Link(item.url, item.title, children, attrs))
            continue
        if not isinstance(item, _Text):
            prepared.append(item)
            continue
        if item.code:
            pieces = re.split(r"(\n)", item.text.replace("\t", "    "))
            for piece in pieces:
                if not piece:
                    continue
                if piece == "\n" or not piece.strip():
                    prepared.append(_Text(piece, item.attrs, False))
                else:
                    prepared.append(_Text(piece, item.attrs, True))
        else:
            prepared.append(_Text(item.text.replace("\t", " "), item.attrs, False))
    merged: list[Any] = []
    for item in prepared:
        previous = merged[-1] if merged else None
        if (isinstance(item, _Text) and isinstance(previous, _Text)
                and previous.attrs == item.attrs and previous.code == item.code):
            merged[-1] = _Text(previous.text + item.text, item.attrs, item.code)
        else:
            merged.append(item)
    return merged


def _build_tree(items: list[Any], active: frozenset[str]) -> list[Any]:
    """Fold formatted segments into nested spans, longest span outermost."""
    nodes: list[Any] = []
    index = 0
    while index < len(items):
        extra = _item_attrs(items[index]) - active
        if not extra:
            nodes.append(items[index])
            index += 1
            continue
        best, best_length = "", 0
        for attr in _ATTRS:
            if attr not in extra:
                continue
            end = index
            while end < len(items) and attr in _item_attrs(items[end]):
                end += 1
            if end - index > best_length:
                best, best_length = attr, end - index
        children = _build_tree(items[index:index + best_length], active | {best})
        nodes.append(_Span(best, children))
        index += best_length
    return nodes


def _add_attr(items: list[Any], attr: str) -> list[Any]:
    result = []
    for item in items:
        if isinstance(item, _Text):
            result.append(_Text(item.text, item.attrs | {attr}, item.code))
        elif isinstance(item, _Link):
            result.append(_Link(item.url, item.title, _add_attr(item.children, attr), item.attrs))
        else:
            result.append(item)
    return result


def _drop_attr(items: list[Any], attr: str) -> list[Any]:
    result = []
    for item in items:
        if isinstance(item, _Text):
            result.append(_Text(item.text, item.attrs - {attr}, item.code))
        elif isinstance(item, _Link):
            result.append(_Link(item.url, item.title, _drop_attr(item.children, attr), item.attrs))
        else:
            result.append(item)
    return result
