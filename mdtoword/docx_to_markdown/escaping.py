"""Markdown escaping: plain text, line starts, code spans and link destinations.

Plain text is escaped so it re-parses to the same text: emphasis and link
punctuation everywhere, ``$`` (the forward converter's math delimiter), and
block syntax (``#``, ``>``, ``-``, ``1.``) wherever a line starts --
including after a hard line break.
"""

from __future__ import annotations

import re

_ENTITY = re.compile(r"&(?:#[0-9]{1,7}|#[xX][0-9A-Fa-f]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")
_ORDERED_START = re.compile(r"(\d{1,9})([.)])(?=[ \t]|$)")
_LINE_START_RULES = (
    re.compile(r"#{1,6}(?=[ \t]|$)"),
    re.compile(r">"),
    re.compile(r"[-+](?=[ \t]|$)"),
    re.compile(r"=+[ \t]*$"),
    re.compile(r"-+[ \t]*$"),
)


def _escape_text(text: str, *, cell: bool = False) -> str:
    """Escape plain text so Markdown reads it back as the same text."""
    out: list[str] = []
    length = len(text)
    for index, ch in enumerate(text):
        previous = text[index - 1] if index else ""
        following = text[index + 1] if index + 1 < length else ""
        if ch in "\\`*[]$~^":
            out.append("\\" + ch)
        elif ch == "_":
            out.append("_" if previous.isalnum() and following.isalnum() else "\\_")
        elif ch == "<":
            out.append("\\<" if following.isalpha() or following in "/!?" else "<")
        elif ch == "|":
            out.append("|" if cell else "\\|")
        elif ch == "&":
            out.append("\\&" if _ENTITY.match(text, index) else "&")
        elif ch == "=" and ("=" in (previous, following)):
            # `==` is the highlight delimiter; a lone `=` is just text.
            run_start = index
            while run_start > 0 and text[run_start - 1] == "=":
                run_start -= 1
            run_end = index
            while run_end + 1 < length and text[run_end + 1] == "=":
                run_end += 1
            before = text[run_start - 1] if run_start else ""
            after = text[run_end + 1] if run_end + 1 < length else ""
            spaced = before.isspace() and after.isspace()
            out.append("=" if spaced else "\\=")
        else:
            out.append(ch)
    return "".join(out)


def _escape_line_start(line: str) -> str:
    """Escape what would make a paragraph line start a block of its own."""
    match = _ORDERED_START.match(line)
    if match:
        return match.group(1) + "\\" + line[len(match.group(1)):]
    for rule in _LINE_START_RULES:
        if rule.match(line):
            return "\\" + line
    return line


def _code_span(text: str) -> str:
    longest = max((len(run) for run in re.findall(r"`+", text)), default=0)
    fence = "`" * (longest + 1)
    if (text.startswith("`") or text.endswith("`")
            or (text.startswith(" ") and text.endswith(" ") and text.strip())):
        text = f" {text} "
    return f"{fence}{text}{fence}"


def _link_destination(url: str) -> str:
    """A link destination, in ``<...>`` form when the bare form would break."""
    url = url.replace("\\", "%5C")
    if not url or any(ch in url for ch in " <>()") or any(ord(ch) < 0x20 for ch in url):
        return "<" + url.replace("<", "%3C").replace(">", "%3E").replace("\n", "%0A") + ">"
    return url
