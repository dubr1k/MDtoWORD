"""Footnotes and endnotes: ``[^N]`` references in the text, definitions at the end.

Notes are read from their own parts -- relationships (images, links) are
resolved against *those* parts -- and appended as ``[^N]:`` definitions
(``[^eN]:`` for endnotes) in order of first reference.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .blocks import _BlockWriter
from .inline import _Raw
from .wordml import W_ENDNOTE, W_FOOTNOTE, W_ID, W_TYPE, element_of, related_part

if TYPE_CHECKING:
    from .document import _Conversion


class _Notes:
    """The document's footnotes and endnotes, and the order they are referenced in."""

    def __init__(self, main_part: Any) -> None:
        self.notes: dict[str, tuple[Any, dict[str, Any]]] = {}
        self.order: list[tuple[str, str]] = []
        for kind, reltype in (("footnote", RT.FOOTNOTES), ("endnote", RT.ENDNOTES)):
            part = related_part(main_part, reltype)
            if part is None:
                continue
            element = element_of(part)
            tag = W_FOOTNOTE if kind == "footnote" else W_ENDNOTE
            notes = {}
            for note in element.iterchildren(tag):
                if note.get(W_TYPE) in (None, "normal"):
                    notes[note.get(W_ID) or ""] = note
            self.notes[kind] = (part, notes)

    def reference(self, kind: str, element: Any) -> _Raw | None:
        """The ``[^N]`` marker of a note reference, recording the note's order."""
        note_id = element.get(W_ID)
        if note_id is None:
            return None
        key = (kind, note_id)
        if key not in self.order:
            self.order.append(key)
        label = note_id if kind == "footnote" else f"e{note_id}"
        return _Raw(f"[^{label}]")

    def definitions(self, conversion: _Conversion) -> list[str]:
        """Every referenced note as a definition block, nested references included."""
        blocks: list[str] = []
        index = 0
        while index < len(self.order):
            kind, note_id = self.order[index]
            index += 1
            part, notes = self.notes.get(kind, (None, {}))
            note = notes.get(note_id)
            label = note_id if kind == "footnote" else f"e{note_id}"
            if note is None:
                conversion.diagnostics.warn("footnote_missing")
                continue
            body = _BlockWriter(conversion, part).render(conversion.block_elements(note))
            blocks.append(self._definition(label, body))
        return blocks

    @staticmethod
    def _definition(label: str, body: list[str]) -> str:
        if not body:
            return f"[^{label}]: "
        lines: list[str] = []
        for block_index, block in enumerate(body):
            block_lines = block.split("\n")
            if block_index == 0:
                lines.append(f"[^{label}]: {block_lines[0]}")
                block_lines = block_lines[1:]
            else:
                lines.append("")
            lines.extend("    " + line if line else "" for line in block_lines)
        return "\n".join(lines)
