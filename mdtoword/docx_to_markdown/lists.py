"""Markdown lists: nesting, markers, loose lists and blocks attached to items.

Word has no list containment, only numbering and indentation. Nesting is
decided by the item's text indent -- which also works for style-based lists
such as ``List Bullet 2`` that live on level 0 of a separate list -- and a
block indented under an open item's text belongs to that item.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .inline import _Text
from .numbering import _Level

_TASK_GLYPHS = {"☐": "[ ]", "□": "[ ]", "☒": "[x]", "☑": "[x]", "✓": "[x]",
                "✔": "[x]", "✅": "[x]"}


@dataclass
class _OpenItem:
    key: tuple[int, int]
    num_id: str
    ordered: bool
    content_indent: int
    text_indent: int
    marker: str


def _other_marker(marker: str) -> str:
    """The alternative list marker: switching it starts a separate list."""
    return {"-": "*", "*": "-", ".": ")", ")": "."}.get(marker, marker)


def split_task(items: list[Any]) -> tuple[str, list[Any]]:
    """``("[ ] ", rest)`` for an item starting with a check-box glyph, else ``("", items)``."""
    task = ""
    first = next((i for i, item in enumerate(items)
                  if not (isinstance(item, _Text) and not item.text.strip())), None)
    if first is not None and isinstance(items[first], _Text):
        glyph_item = items[first]
        stripped = glyph_item.text.lstrip()
        if stripped[:1] in _TASK_GLYPHS:
            task = _TASK_GLYPHS[stripped[0]] + " "
            rest = stripped[1:].lstrip()
            items = items[first + 1:]
            if rest:
                items.insert(0, _Text(rest, glyph_item.attrs, glyph_item.code))
    return task, items


class _ListBuilder:
    """The open Markdown list of a block writer: its lines and the stack of open items."""

    def __init__(self) -> None:
        self.lines: list[str] | None = None
        self.stack: list[_OpenItem] = []
        # Depths (positions in `stack`) whose list holds blocks inside its
        # items, so its items are separated by blank lines; and whether the
        # last thing written into the list was such a block.
        self.loose_depths: set[int] = set()
        self.after_block = False

    def close(self) -> str | None:
        """End the list; return its Markdown, or ``None`` if nothing was written."""
        block = "\n".join(self.lines) if self.lines else None
        self.lines = None
        self.stack = []
        self.loose_depths = set()
        self.after_block = False
        return block

    def add(self, level: _Level, value: int, text: str, text_indent: int) -> str | None:
        """Add one item; return the finished list it ended by starting another, if any."""
        finished = None
        key = (text_indent, level.ilvl)
        while self.stack and self.stack[-1].key > key:
            self.stack.pop()
        marker_char = None
        if self.stack and self.stack[-1].key == key:
            previous = self.stack.pop()
            same = previous.num_id == level.num_id
            if previous.ordered == level.ordered:
                # Markdown would merge two adjacent lists of one kind; a
                # different marker is what keeps them apart.
                marker_char = previous.marker if same else _other_marker(previous.marker)
            if not self.stack and not (same and previous.ordered == level.ordered):
                # Another list at the top level: give it a block of its own.
                if self.lines:
                    finished = "\n".join(self.lines)
                self.lines = []
                self.loose_depths = set()
                self.after_block = False
        if marker_char is None:
            marker_char = "." if level.ordered else "-"
        indent = self.stack[-1].content_indent if self.stack else 0
        marker = f"{value}{marker_char}" if level.ordered else marker_char
        content_indent = indent + len(marker) + 1
        body = text.rstrip().replace("\n", "\n" + " " * content_indent)
        depth = len(self.stack)
        if self.lines is None:
            self.lines = []
        elif depth in self.loose_depths or self.after_block:
            # This list has blocks inside its items, so it is loose: a blank
            # line keeps the next item from reading as part of them.
            self.lines.append("")
        self.loose_depths = {loose for loose in self.loose_depths if loose <= depth}
        self.after_block = False
        self.lines.append(" " * indent + marker + " " + body)
        self.stack.append(_OpenItem(key, level.num_id, level.ordered, content_indent,
                                    text_indent, marker_char))
        return finished

    def attach(self, block: str, indent: int) -> bool:
        """Put a block inside the open list item whose text it is indented under.

        Word has no containment, only indentation: a paragraph, code block,
        table or equation indented at least as far as an item's text belongs
        to that item. Returns ``False`` when no open item qualifies.
        """
        if indent <= 0 or self.lines is None or not self.stack:
            return False
        candidates = [index for index, item in enumerate(self.stack)
                      if 0 < item.text_indent <= indent]
        if not candidates:
            return False
        del self.stack[candidates[-1] + 1:]
        item = self.stack[-1]
        self.lines.append("")
        self.lines.extend(
            " " * item.content_indent + line if line else ""
            for line in block.split("\n")
        )
        self.loose_depths.add(len(self.stack) - 1)
        self.after_block = True
        return True
