"""Printing inline trees as Markdown.

The tree is printed with whitespace kept *outside* the markers. A span
whose ``*``/``~~``/``==`` delimiters would not parse back -- CommonMark's
flanking rules reject e.g. ``x**(a)**`` -- falls back to its HTML tag.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any

from .escaping import _code_span, _escape_line_start, _escape_text, _link_destination
from .inline import _build_tree, _Link, _plain, _prepare, _Raw, _Span, _Text

_SCHEME = re.compile(r"[A-Za-z][A-Za-z0-9+.-]{1,31}:")

_MD_DELIMITERS = {"bold": "**", "italic": "*", "strike": "~~", "highlight": "=="}
_HTML_TAGS = {"bold": "strong", "italic": "em", "strike": "del", "highlight": "mark",
              "underline": "u", "sup": "sup", "sub": "sub"}


def _is_punctuation(ch: str) -> bool:
    return unicodedata.category(ch)[0] in "PS"


def _is_boundary(ch: str | None) -> bool:
    """Start/end of text, whitespace or punctuation: where a delimiter may sit."""
    return ch is None or ch.isspace() or _is_punctuation(ch)


def _split_outer_whitespace(text: str) -> tuple[str, str, str]:
    core = text.strip()
    if not core:
        return text, "", ""
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    return lead, core, trail


class _InlineRenderer:
    """Print a list of inline items as Markdown in one of three modes.

    ``block`` -- a paragraph: hard breaks as backslash-newline, line starts
    escaped; ``cell`` -- a GFM table cell: breaks as ``<br>``, every ``|``
    escaped; ``heading`` -- one line, breaks as spaces.
    """

    def __init__(self, mode: str) -> None:
        self.mode = mode

    def render(self, items: list[Any]) -> str:
        tree = _build_tree(_prepare(items), frozenset())
        return self._finish(self._nodes(tree))

    def _finish(self, text: str) -> str:
        text = text.strip(" \t\n")
        text = re.sub(r"[ \t]*\n[ \t]*", "\n", text)
        if self.mode == "heading":
            text = text.replace("\n", " ")
            match = re.search(r"(?:^|[ \t])(#+)[ \t]*$", text)
            if match:
                text = text[: match.start(1)] + "\\" + text[match.start(1):]
            return text
        if self.mode == "cell":
            return text.replace("\n", "<br>").replace("|", "\\|")
        return "\\\n".join(_escape_line_start(line) for line in text.split("\n"))

    def _nodes(self, nodes: list[Any]) -> str:
        parts: list[tuple[Any, ...]] = []
        for node in nodes:
            if isinstance(node, _Span):
                lead, core, trail = _split_outer_whitespace(self._nodes(node.children))
                parts.append(("span", node.attr, lead, core, trail))
            else:
                parts.append(("leaf", self._leaf(node)))
        out: list[str] = []
        for index, part in enumerate(parts):
            if part[0] == "leaf":
                text = part[1]
                # `[^1](x)` would read as a link: keep a following "(" literal.
                if text.startswith("(") and out and out[-1].endswith("]"):
                    text = "\\" + text
                out.append(text)
                continue
            _, attr, lead, core, trail = part
            if lead:
                out.append(lead)
            if core:
                before = next((chunk[-1] for chunk in reversed(out) if chunk), None)
                after = trail[0] if trail else self._next_char(parts, index + 1)
                out.append(self._wrap(attr, core, before, after))
            if trail:
                out.append(trail)
        return "".join(out)

    @staticmethod
    def _next_char(parts: list[tuple[Any, ...]], start: int) -> str | None:
        for part in parts[start:]:
            if part[0] == "leaf":
                if part[1]:
                    return part[1][0]
                continue
            _, attr, lead, core, trail = part
            if lead:
                return lead[0]
            if core:
                return _MD_DELIMITERS.get(attr, "<")[0]
            if trail:
                return trail[0]
        return None

    @staticmethod
    def _wrap(attr: str, core: str, before: str | None, after: str | None) -> str:
        tag = _HTML_TAGS[attr]
        html = f"<{tag}>{core}</{tag}>"
        delimiter = _MD_DELIMITERS.get(attr)
        if delimiter is None:
            return html
        mark = delimiter[0]
        if mark == "*":
            # Leading/trailing unescaped `*` belong to nested spans and merge
            # with ours into one delimiter run; flanking is decided past them.
            first_text = core.lstrip("*")
            last_text = core.rstrip("*")
            if core.endswith("\\*"):
                last_text = core
            first = first_text[:1] or None
            last = last_text[-1:] or None
        else:
            if core[0] == mark or core[-1] == mark:
                return html
            first, last = core[0], core[-1]
        if first is None or last is None:
            return html
        if _is_punctuation(first) and not _is_boundary(before):
            return html
        if _is_punctuation(last) and not _is_boundary(after):
            return html
        if before == mark or after == mark:
            return html
        return f"{delimiter}{core}{delimiter}"

    def _leaf(self, node: Any) -> str:
        if isinstance(node, _Text):
            if node.code:
                return _code_span(node.text)
            return _escape_text(node.text, cell=self.mode == "cell")
        if isinstance(node, _Raw):
            return node.markdown
        if isinstance(node, _Link):
            return self._link(node)
        return ""

    def _link(self, link: _Link) -> str:
        text = self._nodes(_build_tree(link.children, link.attrs))
        plain = _plain(link.children)
        url = link.url
        if not text.strip():
            return text
        if (plain == url and _SCHEME.match(url)
                and not any(ch in url for ch in " <>") and not link.title):
            return f"<{url}>"
        if (url.lower().startswith("mailto:") and plain == url[7:] and "@" in plain
                and not any(ch in plain for ch in " <>\\")):
            return f"<{plain}>"
        title = ""
        if link.title:
            title = ' "' + link.title.replace("\\", "\\\\").replace('"', '\\"') + '"'
        return f"[{text}]({_link_destination(url)}{title})"


def render_inline(items: list[Any], mode: str = "block") -> str:
    return _InlineRenderer(mode).render(items)
