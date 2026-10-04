"""Footnotes: native Word footnotes, or a closing "Footnotes" section.

In native mode each reference reserves a Word footnote and the definition's
blocks are rendered straight into it (containers put aside meanwhile); a
note referenced twice is copied into a second Word footnote, its pictures
renumbered; a note referenced from inside another footnote -- which Word
cannot nest -- is appended to that footnote behind a visible mark. In
section mode the notes become ordinary paragraphs under a localized
heading, each starting with its mark.
"""

from __future__ import annotations

from collections.abc import Sequence
import copy
from typing import Any

from docx.oxml.ns import qn

from .state import RendererState


class FootnoteMixin(RendererState):
    """Reference, open, fill, close and duplicate footnotes."""

    def _localized_footnotes_heading(self) -> str:
        if self.footnotes_heading == "Footnotes" and self._russian:
            return "Сноски"
        return self.footnotes_heading

    def _footnote_reference(self, token: Any) -> None:
        meta = token.meta or {}
        footnote_id = meta.get("id", -1)
        if self._footnotes is not None and self._footnote_target is None:
            word_id = self._footnotes.reserve()
            self._footnote_word_ids.setdefault(footnote_id, []).append(word_id)
            self._footnotes.add_reference(self._paragraph, word_id)
            return
        if self._footnotes is not None:
            # Word has no footnotes inside footnotes (a reference in
            # footnotes.xml makes the file unreadable): the inner note's text
            # is appended to this footnote, behind the same mark.
            self._nested_footnote_parent.setdefault(footnote_id, self._footnote_target)
            self._warn(
                f"Footnote [^{meta.get('label') or footnote_id + 1}] is referenced inside another "
                "footnote; Word cannot nest footnotes, so its text is appended to that footnote",
                "footnote_nested",
            )
        run = self._paragraph.add_run(self._footnote_mark(meta))
        run.font.superscript = True
        if self._link is not None:
            self._link.append(run._r)

    def _footnote_mark(self, meta: dict[str, Any]) -> str:
        """The visible mark of a footnote that is not a native Word footnote.

        Section-mode notes are numbered in order of first reference, like
        Word's own numbering; a note nested in another footnote keeps its
        Markdown label so the reader can match mark and text.
        """
        if self._footnotes is None:
            return str(meta.get("id", 0) + 1)
        return str(meta.get("label") or meta.get("id", 0) + 1)

    def _open_footnote(self, token: Any) -> None:
        meta = token.meta or {}
        if self._footnotes is None:
            # The mark is written into the footnote's first paragraph by
            # _add_paragraph, so mark and text share one line.
            self._pending_footnote_label = self._footnote_mark(meta)
            self._footnote_section_open = True
            return
        targets = self._footnote_word_ids.get(meta.get("id", -1), [])
        parent = self._nested_footnote_parent.get(meta.get("id", -1))
        if not targets and parent is not None:
            self._saved_containers = (self._lists, self._quotes, self._definition_depth)
            self._lists, self._quotes, self._definition_depth = [], [], 0
            self._footnote_target = parent
            self._pending_footnote_label = self._footnote_mark(meta)
            return
        if not targets:
            # Defined but never referenced: Word has nowhere to anchor it.
            self._warn(
                f"Footnote [^{meta.get('label', '?')}] is defined but never referenced; skipped",
                "footnote_unreferenced",
            )
            self._skip_footnote = True
            return
        self._saved_containers = (self._lists, self._quotes, self._definition_depth)
        self._lists, self._quotes, self._definition_depth = [], [], 0
        self._footnote_target = targets[0]

    def _close_footnote(self) -> None:
        if self._skip_footnote:
            self._skip_footnote = False
            return
        if self._footnotes is None:
            self._paragraph = None
            self._footnote_section_open = False
            return
        first = self._footnote_target
        self._footnote_target = None
        self._paragraph = None
        self._lists, self._quotes, self._definition_depth = self._saved_containers
        if first is None:
            return
        for word_ids in self._footnote_word_ids.values():
            if word_ids and word_ids[0] == first:
                self._copy_footnote(first, word_ids[1:])
                break

    def _copy_footnote(self, source_id: int, target_ids: Sequence[int]) -> None:
        """A footnote referenced twice becomes two Word footnotes (as pandoc does)."""
        if not target_ids or self._footnotes is None:
            return
        part_element = self._footnotes_part_element()
        if part_element is None:
            return
        source = self._footnote_element(part_element, source_id)
        if source is None:
            return
        for target_id in target_ids:
            target = self._footnote_element(part_element, target_id)
            if target is None:
                continue
            for child in list(target):
                target.remove(child)
            for child in source:
                duplicate = copy.deepcopy(child)
                self._renumber_drawings(duplicate, part_element)
                target.append(duplicate)

    def _renumber_drawings(self, element: Any, footnotes_element: Any) -> None:
        """Give every picture in *element* a document-unique ``wp:docPr`` id."""
        pictures = list(element.iter(qn("wp:docPr")))
        if not pictures:
            return
        used = [
            int(value)
            for root in (self.document.element, footnotes_element)
            for picture in root.iter(qn("wp:docPr"))
            if (value := picture.get("id", "")).isdigit()
        ]
        next_id = max(used, default=0) + 1
        for picture in pictures:
            picture.set("id", str(next_id))
            next_id += 1

    def _footnotes_part_element(self) -> Any:
        for relationship in self.document.part.rels.values():
            if relationship.reltype.endswith("/footnotes"):
                return relationship.target_part.element
        return None

    @staticmethod
    def _footnote_element(part_element: Any, footnote_id: int) -> Any:
        for footnote in part_element.findall(qn("w:footnote")):
            if footnote.get(qn("w:id")) == str(footnote_id):
                return footnote
        return None
