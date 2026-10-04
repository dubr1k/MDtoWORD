"""Numbered captions, paired with the picture or table they label.

A figure caption becomes the picture's title; a table caption becomes a
"Table: ..." line next to the table -- the forms Markdown renderers (and
the forward converter) number and caption themselves, so the label and
number are not repeated on the way back.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .escaping import _escape_text
from .wordml import (
    A_BLIP,
    M_OMATH,
    R_ID,
    V_IMAGEDATA,
    W_FLDSIMPLE,
    W_INSTR,
    W_INSTRTEXT,
    W_P,
    W_TBL,
    _plain_text,
    _q,
)

if TYPE_CHECKING:
    from .classify import _Classifier

# "Figure 1: Text", "Рисунок 1 — Текст", "Table 2.1. Text" -- the number being
# a SEQ field's result.
_CAPTION = re.compile(
    r"(?P<label>[^\W\d_]+\.?)\s*(?P<number>\d+(?:[.\-–]\d+)*)"
    r"(?:\s*[—–]\s*|\s+-\s+|\s*:\s*|\.\s+)(?P<text>\S.*)",
    re.DOTALL,
)
_TABLE_LABELS = frozenset({"table", "tab", "tbl", "таблица", "табл"})
_FIGURE_LABELS = frozenset({
    "figure", "fig", "image", "picture", "illustration", "рисунок", "рис",
    "иллюстрация", "изображение", "схема", "диаграмма",
})


def caption_parts(classifier: _Classifier, paragraph: Any) -> tuple[str | None, str, str] | None:
    """``(kind hint, label, text)`` of a numbered caption paragraph, else ``None``.

    A caption is a paragraph in the ``Caption`` style reading
    "<label> <number><separator><text>", where the number is a SEQ
    field (as Word and the forward converter write it) or the label is
    a known figure/table word. Anything else is an ordinary paragraph.
    """
    info = classifier.paragraph_info(paragraph)
    if not any(style.name == "caption" for style in info.chain):
        return None
    match = _CAPTION.fullmatch(_plain_text(paragraph).strip())
    if match is None:
        return None
    instructions = " ".join(
        [t.text or "" for t in paragraph.iter(W_INSTRTEXT)]
        + [f.get(W_INSTR) or "" for f in paragraph.iter(W_FLDSIMPLE)]
    )
    sequence = re.search(r"\bSEQ\s+(\"[^\"]+\"|\S+)", instructions, re.IGNORECASE)
    label = match.group("label")
    words = {label.rstrip(".").lower()}
    if sequence:
        words.add(sequence.group(1).strip('"').lower())
    if sequence is None and not words & (_TABLE_LABELS | _FIGURE_LABELS):
        return None
    hint = ("table" if words & _TABLE_LABELS
            else "figure" if words & _FIGURE_LABELS else None)
    return hint, label, match.group("text").strip()


def single_picture(element: Any) -> bool:
    """A paragraph holding exactly one picture and no text."""
    if element is None or element.tag != W_P or _plain_text(element).strip():
        return False
    if next(element.iter(M_OMATH), None) is not None:
        return False
    pictures = sum(1 for _ in element.iter(A_BLIP))
    pictures += sum(1 for data in element.iter(V_IMAGEDATA)
                    if data.get(R_ID) or data.get(_q("r:href")))
    return pictures == 1


class _CaptionMatcher:
    """Captions of one run of elements, paired with their pictures and tables.

    ``figures`` maps a picture paragraph to its caption text, ``tables`` a
    table to ``(side, "Table: ..." line)``; the caption paragraphs used are
    collected in ``skipped``.
    """

    def __init__(self, classifier: _Classifier) -> None:
        self.classifier = classifier
        self.skipped: set[Any] = set()
        self.figures: dict[Any, str] = {}
        self.tables: dict[Any, tuple[str, str]] = {}

    def match(self, elements: list[Any]) -> None:
        """Pair numbered caption paragraphs with the picture or table they label."""
        for position, element in enumerate(elements):
            if element.tag != W_P or element in self.skipped:
                continue
            parts = caption_parts(self.classifier, element)
            if parts is None:
                continue
            hint, label, text = parts
            previous = elements[position - 1] if position > 0 else None
            following = elements[position + 1] if position + 1 < len(elements) else None
            figure_after = ("figure", previous)
            figure_before = ("figure", following)
            table_before = ("table", following)
            table_after = ("table", previous)
            if hint == "table":
                order = (table_before, table_after, figure_after, figure_before)
            elif hint == "figure":
                order = (figure_after, figure_before, table_before, table_after)
            else:
                order = (figure_after, table_before, table_after, figure_before)
            for kind, target in order:
                if target is None or target in self.skipped:
                    continue
                if (kind == "figure" and target not in self.figures
                        and single_picture(target)):
                    self.figures[target] = text
                elif (kind == "table" and target.tag == W_TBL
                        and target not in self.tables):
                    word = "Таблица" if re.search("[А-Яа-яЁё]", label) else "Table"
                    side = "before" if target is following else "after"
                    self.tables[target] = (side, f"{word}: {_escape_text(text)}")
                else:
                    continue
                self.skipped.add(element)
                break
