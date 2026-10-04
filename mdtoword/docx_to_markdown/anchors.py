"""Heading anchors: GitHub slugs, and the bookmarks on headings that links point to."""

from __future__ import annotations

import unicodedata
from typing import TYPE_CHECKING, Any

from .wordml import W_BOOKMARKSTART, W_NAME, W_P, _plain_text

if TYPE_CHECKING:
    from .classify import _Classifier
    from .numbering import _Numbering


def github_slug(text: str, seen: dict[str, int]) -> str:
    """GitHub's heading anchor: lower case, punctuation dropped, spaces to ``-``.

    Letters of every script are kept, so Cyrillic headings get Cyrillic
    anchors exactly as GitHub renders them; repeats get ``-1``, ``-2``.
    """
    slug = "".join(
        "-" if ch == " " else ch
        for ch in text.strip().lower()
        if ch in " -_" or unicodedata.category(ch)[0] in "LNM"
    )
    result = slug
    while result in seen:
        seen[slug] += 1
        result = f"{slug}-{seen[slug]}"
    seen[result] = 0
    return result


def heading_anchors(body: Any, elements: list[Any], classifier: _Classifier,
                    numbering: _Numbering) -> dict[str, str]:
    """Map every bookmark on a heading to that heading's GitHub anchor.

    Done before writing anything, since links may point forward. A
    numbered heading's anchor includes its number ("1-intro" for
    "1. Intro"), so list counters are replayed on a scratch copy in
    document order -- table cells included, as Word counts them.
    """
    anchors: dict[str, str] = {}
    seen: dict[str, int] = {}
    top_level = {element for element in elements if element.tag == W_P}
    scratch = numbering.fresh()
    for element in body.iter(W_P):
        info = classifier.paragraph_info(element)
        label = ""
        if info.level is not None:
            scratch.advance(info.level)
            if info.heading and info.level.ordered:
                label = scratch.label(info.level)
        if element not in top_level or not 1 <= info.heading <= 6:
            continue
        text = " ".join(part for part in (label, _plain_text(element).strip()) if part)
        if not text:
            continue
        slug = github_slug(text, seen)
        for bookmark in element.iter(W_BOOKMARKSTART):
            name = bookmark.get(W_NAME)
            if name:
                anchors[name] = slug
    return anchors
