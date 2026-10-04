"""Raw HTML, inline and as whole blocks, rendered as formatted text.

Tags that map onto run formatting (``<b>``, ``<sub>``, ``<kbd>``, …),
links, images and line breaks are honoured; presentational wrappers are
unwrapped silently; block-level tags start new paragraphs; script-like tags
drop their content. Anything else is removed with one warning per tag
name, keeping its text.
"""

from __future__ import annotations

import html
from pathlib import Path
import re
from typing import Any

from .constants import (
    _HTML_BLOCK_BREAKS,
    _HTML_COMMENT,
    _HTML_DROP_CONTENT,
    _HTML_FORMAT,
    _HTML_PIECES,
    _HTML_TAG,
    _HTML_TRANSPARENT,
    _XML_INVALID,
)
from .helpers import _html_attributes, _HtmlImage, _plain_formatting
from .state import RendererState


class RawHtmlMixin(RendererState):
    """Render inline HTML tags and HTML blocks."""

    def _html_inline(
        self,
        raw: str,
        formatting: dict[str, bool],
        link: Any,
        source_path: Path | None,
    ) -> Any:
        raw = raw.strip()
        if _HTML_COMMENT.match(raw):
            return link
        match = _HTML_TAG.match(raw)
        if match is None:
            self._warn_html(raw[:20], raw)
            return link
        name = match.group("name").lower()
        closing = bool(match.group("close"))
        attributes = _html_attributes(match.group("attrs") or "")

        if name == "br":
            self._append_break(link)
            return link
        if name == "img" and not closing:
            self._append_image(_HtmlImage(attributes), source_path)
            return link
        if name == "a":
            if closing:
                return None
            href = attributes.get("href")
            if href:
                return self._open_link(href, attributes.get("title"))
            return link
        flag = _HTML_FORMAT.get(name)
        if flag is not None:
            formatting[flag] = not closing
            return link
        if name in _HTML_TRANSPARENT or name in _HTML_BLOCK_BREAKS:
            return link
        if not closing:
            self._warn_html(name, raw)
        return link

    def _warn_html(self, name: str, raw: str, kept: bool = True) -> None:
        if name in self._warned_html:
            return
        self._warned_html.add(name)
        what = "its text is kept" if kept else "its content is not document text and is dropped"
        self._warn(f"HTML tag not supported and removed ({what}): {raw[:60]}", "html_dropped")

    def _render_html_block(self, content: str, source_path: Path | None) -> None:
        """Render a raw HTML block as text, keeping the formatting we understand."""
        self._ensure_item_started()
        formatting = _plain_formatting()
        self._link = None
        # An <a> may wrap several block elements (a "card"): its target is
        # remembered and a link reopened in every paragraph it spans.
        link_target: tuple[str, str | None] | None = None
        drop_depth = 0
        preformatted = 0
        self._paragraph = None

        def ensure_paragraph() -> None:
            if self._paragraph is None:
                self._paragraph = self._new_paragraph()
                self._link = (
                    self._open_link(*link_target) if link_target is not None else None
                )

        for piece in _HTML_PIECES.split(content):
            if not piece:
                continue
            match = _HTML_TAG.match(piece) if piece.startswith("<") else None
            if piece.startswith("<!--") and _HTML_COMMENT.match(piece):
                continue
            if match is not None:
                name = match.group("name").lower()
                closing = bool(match.group("close"))
                if name in _HTML_DROP_CONTENT:
                    if match.group("self"):
                        continue
                    drop_depth = max(0, drop_depth - 1) if closing else drop_depth + 1
                    if not closing:
                        self._warn_html(name, piece, kept=False)
                    continue
                if drop_depth:
                    continue
                if name == "a":
                    attributes = _html_attributes(match.group("attrs") or "")
                    if closing or not attributes.get("href"):
                        link_target = None
                        self._link = None
                    else:
                        link_target = (attributes["href"], attributes.get("title"))
                        self._link = (
                            self._open_link(*link_target) if self._paragraph is not None else None
                        )
                    continue
                if name == "pre":
                    preformatted += -1 if closing else 1
                if name in _HTML_BLOCK_BREAKS or name == "br":
                    if name == "br" and self._paragraph is not None:
                        self._append_break(self._link)
                        continue
                    self._paragraph = None
                    self._link = None
                    if name == "summary":
                        formatting["bold"] = not closing
                    if name in {"h1", "h2", "h3", "h4", "h5", "h6", "dt", "th"}:
                        formatting["bold"] = not closing
                    if name == "li" and not closing:
                        ensure_paragraph()
                        self._paragraph.add_run("• ")
                    if name == "hr" and not closing:
                        self._add_thematic_break()
                    continue
                if name == "img" and not closing:
                    ensure_paragraph()
                    self._append_image(_HtmlImage(_html_attributes(match.group("attrs") or "")), source_path)
                    continue
                ensure_paragraph()
                self._link = self._html_inline(piece, formatting, self._link, source_path)
                continue
            if drop_depth:
                continue
            text = _XML_INVALID.sub("", html.unescape(piece))
            if not preformatted:
                text = re.sub(r"\s+", " ", text)
                if self._paragraph is None:
                    text = text.lstrip()
            if not text.strip() and self._paragraph is None:
                continue
            ensure_paragraph()
            self._append_text(
                text, formatting if not preformatted else {**formatting, "code": True}, self._link
            )
        self._paragraph = None
        self._link = None
