"""Small markdown-it extensions and helpers the Word renderer relies on.

``mdit_py_plugins`` already ships ``~subscript~`` but has no counterpart for
``^superscript^`` or ``==highlight==``, and the project deliberately has no
YAML dependency for front matter. Each piece here is narrow on purpose: it
recognises exactly the syntax the renderer documents and nothing more, so an
unusual construct stays literal text instead of being guessed at.
"""

from __future__ import annotations

import re

from markdown_it import MarkdownIt
from markdown_it.rules_inline import StateInline

_WHITESPACE = re.compile(r"(^|[^\\])(\\\\)*\s")
_UNESCAPE = re.compile(r'\\([ \\!"#$%&\'()*+,.\/:;<=>?@[\]^_`{|}~-])')


def _superscript(state: StateInline, silent: bool) -> bool:
    """``^text^`` -- the markdown-it-sup rule: no unescaped whitespace inside."""
    start = state.pos
    maximum = state.posMax
    if silent or state.src[start] != "^" or start + 2 >= maximum:
        return False

    state.pos = start + 1
    found = False
    while state.pos < maximum:
        if state.src[state.pos] == "^":
            found = True
            break
        state.md.inline.skipToken(state)

    if not found or start + 1 == state.pos:
        state.pos = start
        return False

    content = state.src[start + 1 : state.pos]
    if _WHITESPACE.search(content) is not None:
        state.pos = start
        return False

    state.posMax = state.pos
    state.pos = start + 1
    token = state.push("sup_open", "sup", 1)
    token.markup = "^"
    token = state.push("text", "", 0)
    token.content = _UNESCAPE.sub(r"\1", content)
    token = state.push("sup_close", "sup", -1)
    token.markup = "^"
    state.pos = state.posMax + 1
    state.posMax = maximum
    return True


def sup_plugin(md: MarkdownIt) -> None:
    """Register ``^superscript^``."""
    md.inline.ruler.after("emphasis", "sup", _superscript)


def _mark(state: StateInline, silent: bool) -> bool:
    """``==text==`` -- content may contain spaces but not start or end with one."""
    start = state.pos
    maximum = state.posMax
    if silent or not state.src.startswith("==", start):
        return False
    if start > 0 and state.src[start - 1] == "=":
        return False

    # Look for the closing "==" token by token, so one inside a code span
    # or a link destination is not mistaken for it.
    end = -1
    state.pos = start + 2
    while state.pos < maximum:
        if state.src.startswith("==", state.pos):
            end = state.pos
            break
        state.md.inline.skipToken(state)
    state.pos = start
    if end == -1:
        return False
    content = state.src[start + 2 : end]
    if not content or content != content.strip() or "\n\n" in content:
        return False
    # "a === b" or "x ==== y" is an operator run, not a highlight.
    if end + 2 < maximum and state.src[end + 2] == "=":
        return False

    old_max = state.posMax
    token = state.push("mark_open", "mark", 1)
    token.markup = "=="
    state.pos = start + 2
    state.posMax = end
    state.md.inline.tokenize(state)
    state.posMax = old_max
    token = state.push("mark_close", "mark", -1)
    token.markup = "=="
    state.pos = end + 2
    return True


def mark_plugin(md: MarkdownIt) -> None:
    """Register ``==highlight==``."""
    md.inline.ruler.before("emphasis", "mark", _mark)


# --- front matter ---------------------------------------------------------

_KEY_VALUE = re.compile(r"^(?P<key>[A-Za-z_][\w-]*)\s*:(?:\s+(?P<value>.*))?$")
_LIST_ITEM = re.compile(r"^\s*-\s+(?P<value>.*)$")


def parse_front_matter(text: str) -> tuple[dict[str, str | list[str]], list[str]]:
    """Parse the flat YAML subset documents actually use for metadata.

    Supported: ``key: value`` with plain, single- or double-quoted scalars;
    inline lists ``[a, b]``; block lists (``key:`` then ``- item`` lines);
    folded/literal block scalars (``key: >`` / ``key: |``). Anything else --
    nested mappings, anchors, multi-document streams -- is skipped and named
    in the second element of the result so the caller can warn about it.
    Keys are lower-cased.
    """
    data: dict[str, str | list[str]] = {}
    skipped: list[str] = []
    lines = text.splitlines()
    index = 0
    while index < len(lines):
        raw = lines[index]
        index += 1
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        if raw[0] in " \t":
            skipped.append(raw.strip())
            continue
        match = _KEY_VALUE.match(raw.rstrip())
        if match is None:
            skipped.append(raw.strip())
            continue
        key = match.group("key").lower()
        value = (match.group("value") or "").strip()

        if value in {"|", ">", "|-", ">-", "|+", ">+"}:
            block: list[str] = []
            while index < len(lines) and (not lines[index].strip() or lines[index][0] in " \t"):
                block.append(lines[index].strip())
                index += 1
            joiner = "\n" if value.startswith("|") else " "
            data[key] = joiner.join(part for part in block if part).strip()
            continue

        if not value:
            items: list[str] = []
            nested = False
            while index < len(lines) and (not lines[index].strip() or lines[index][0] in " \t-"):
                item = _LIST_ITEM.match(lines[index])
                if item is not None and not _KEY_VALUE.match(item.group("value").strip()):
                    items.append(_scalar(item.group("value")))
                elif lines[index].strip():
                    # A mapping, or a list of mappings ("- name: A"): not
                    # flat metadata. Reported whole rather than half-read.
                    nested = True
                index += 1
            if nested:
                skipped.append(key)
            elif items:
                data[key] = items
            continue

        if value.startswith("[") and value.endswith("]"):
            data[key] = [
                _scalar(part) for part in _split_inline_list(value[1:-1]) if part.strip()
            ]
            continue
        if value.startswith("{"):
            skipped.append(key)
            continue
        data[key] = _scalar(value)
    return data, skipped


def _split_inline_list(body: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    quote: str | None = None
    for char in body:
        if quote:
            current.append(char)
            if char == quote:
                quote = None
        elif char in "\"'":
            quote = char
            current.append(char)
        elif char == ",":
            parts.append("".join(current))
            current = []
        else:
            current.append(char)
    parts.append("".join(current))
    return parts


_DOUBLE_QUOTED_ESCAPES = {"n": "\n", "t": "\t", "r": "\r", "0": "\0", '"': '"', "\\": "\\", "/": "/"}


def _scalar(value: str) -> str:
    """One YAML scalar: plain, 'single-quoted' or "double-quoted", maybe with a # comment."""
    value = value.strip()
    if value[:1] in {'"', "'"}:
        closing = _closing_quote(value)
        if closing is not None:
            rest = value[closing + 1 :].strip()
            if not rest or rest.startswith("#"):
                inner = value[1:closing]
                if value[0] == "'":
                    return inner.replace("''", "'")
                return re.sub(
                    r"\\(.)", lambda m: _DOUBLE_QUOTED_ESCAPES.get(m.group(1), m.group(0)), inner
                )
    # A trailing " # comment" is a YAML comment, not part of the value.
    comment = value.find(" #")
    if comment != -1:
        value = value[:comment].rstrip()
    return value


def _closing_quote(value: str) -> int | None:
    """Index of the quote closing the scalar that opens *value*, or None."""
    quote = value[0]
    index = 1
    while index < len(value):
        char = value[index]
        if quote == '"' and char == "\\":
            index += 2
            continue
        if char == quote:
            if quote == "'" and value[index + 1 : index + 2] == "'":
                index += 2
                continue
            return index
        index += 1
    return None
