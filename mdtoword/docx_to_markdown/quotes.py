"""Block quotes and call-outs: nesting depth, GitHub alert titles, ``>`` prefixes.

A run of quoted paragraphs -- a quote style, a bar on the left, or a shaded
call-out box -- becomes one block quote, nested by indentation; each level
is written by its own block writer, so lists and code keep working inside.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from .wordml import W_T, _plain_text, text_runs

if TYPE_CHECKING:
    from .blocks import _BlockWriter
    from .classify import _Classifier, _ParaInfo

_ALERT_TITLES = {
    "note": "note", "примечание": "note", "tip": "tip", "совет": "tip",
    "important": "important", "важно": "important", "warning": "warning",
    "внимание": "warning", "caution": "caution", "осторожно": "caution",
}
# A block quote's depth, as the forward converter writes it: 18pt of extra
# left indent per level, on top of whatever indent the paragraph already has.
_QUOTE_STEP = 360


def is_quoted(info: _ParaInfo) -> bool:
    return bool(info.quote or info.callout or info.quote_bar)


class _QuoteWriter:
    """Write a run of quoted paragraphs; ``nested_writer`` makes the writer of one level."""

    def __init__(self, classifier: _Classifier,
                 nested_writer: Callable[[], _BlockWriter]) -> None:
        self.classifier = classifier
        self.nested_writer = nested_writer

    def render(self, run: list[Any]) -> str:
        """A run of quoted paragraphs as one (possibly nested) block quote."""
        infos = [self.classifier.paragraph_info(element) for element in run]
        # Plain "Quote" paragraphs without the bar nest by rank of their
        # direct indent; style indents are ignored -- "Intense Quote" is
        # indented by its style without being nested.
        ranked = sorted({info.direct_indent for info in infos
                         if not (info.quote_bar or info.callout) and info.quote == 1
                         and info.direct_indent is not None})
        depths = []
        for info in infos:
            if info.quote > 1:
                depths.append(info.quote)
            elif (info.quote_bar or info.callout) and info.direct_indent is not None:
                # Barred paragraphs carry one quote step of indent per level
                # on top of their own -- a list item's own indent included.
                own = info.level.indent if info.level is not None and info.level.indent else 0
                depths.append(max(1, round((info.direct_indent - own) / _QUOTE_STEP)))
            elif len(ranked) > 1 and info.direct_indent in ranked:
                depths.append(ranked.index(info.direct_indent) + 1)
            else:
                depths.append(1)
        return self._level(list(zip(run, depths, strict=True)), 1)

    def _level(self, entries: list[tuple[Any, int]], depth: int) -> str:
        """Write the entries at ``depth`` (deeper ones nested) and prefix ``>``."""
        writer = self.nested_writer()
        marker = ""
        index = 0
        first, first_depth = entries[0]
        kind = self._alert_title(first) if first_depth == depth else None
        if kind:
            marker = f"[!{kind.upper()}]"
            index = 1
        while index < len(entries):
            end = index
            if entries[index][1] > depth:
                while end < len(entries) and entries[end][1] > depth:
                    end += 1
                nested = self._level(entries[index:end], depth + 1)
                writer.flush()
                if nested:
                    writer.blocks.append(nested)
            else:
                while end < len(entries) and entries[end][1] <= depth:
                    end += 1
                writer.feed([element for element, _ in entries[index:end]])
            index = end
        writer.flush()
        body = "\n\n".join(block for block in writer.blocks if block.strip())
        if marker:
            # GitHub wants the alert marker on the line right above its text.
            body = marker + ("\n" + body if body else "")
        if not body:
            return ""
        return "\n".join("> " + line if line else ">" for line in body.split("\n"))

    def _alert_title(self, element: Any) -> str | None:
        """The alert kind when a call-out paragraph is just its bold title."""
        if not self.classifier.paragraph_info(element).callout:
            return None
        kind = _ALERT_TITLES.get(_plain_text(element).strip().rstrip(":").lower())
        runs = [run for run in text_runs(element)
                if "".join(t.text or "" for t in run.iter(W_T)).strip()]
        if kind and runs and all(
                "bold" in self.classifier.run_format(run, in_link=False).attrs for run in runs):
            return kind
        return None
