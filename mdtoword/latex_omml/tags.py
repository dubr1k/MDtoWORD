"""Equation numbering at the source level: `split_equation_tag` takes
``\\tag``/``\\label``/``\\nonumber`` out of a display formula before it
is parsed, so the caller can number the equation itself."""

from __future__ import annotations

from typing import Optional


def _matching_brace(latex: str, opening: int) -> Optional[int]:
    """Index of the ``}`` closing the ``{`` at `opening`, or None.

    ``\\{`` and ``\\}`` are escaped characters, not braces.
    """
    depth = 0
    position = opening
    while position < len(latex):
        character = latex[position]
        if character == "\\":
            position += 2
            continue
        if character == "{":
            depth += 1
        elif character == "}":
            depth -= 1
            if depth == 0:
                return position
        position += 1
    return None


def split_equation_tag(latex: str) -> tuple[str, Optional[str]]:
    r"""Take the equation-numbering commands out of a display formula.

    Returns ``(remaining_latex, tag_text)``.  ``\label{...}``,
    ``\nonumber`` and ``\notag`` are removed and produce nothing.  A single
    ``\tag{...}`` or ``\tag*{...}`` is removed too, and its brace-balanced
    argument comes back as `tag_text` -- ``"3"`` for ``\tag{3}``, ``"A"``
    for ``\tag*{A}`` -- so the caller can set the number beside the
    equation; without one, `tag_text` is None.

    Two or more ``\tag`` commands mean per-line numbers in a multi-line
    formula: those stay in place -- ``latex_to_omml`` renders each one at
    the end of its own line -- and `tag_text` is None, since no single
    number belongs to the whole equation.  Anything malformed (a ``\tag``
    without a braced argument) is left in place as well, for
    ``latex_to_omml`` to report.  This function never raises.
    """
    removals: list = []
    tags: list = []
    position = 0
    while position < len(latex):
        if latex[position] != "\\":
            position += 1
            continue
        start = position
        end = position + 1
        while end < len(latex) and latex[end].isascii() and latex[end].isalpha():
            end += 1
        if end == position + 1:
            # A control symbol such as `\\` or `\{`: skip both characters,
            # so `\\tag` is a line break followed by the letters "tag".
            position += 2
            continue
        name = latex[position + 1:end]
        position = end
        if name in ("nonumber", "notag"):
            removals.append((start, end))
            continue
        if name not in ("tag", "label"):
            continue
        probe = end
        while probe < len(latex) and latex[probe].isspace():
            probe += 1
        if name == "tag" and probe < len(latex) and latex[probe] == "*":
            probe += 1
            while probe < len(latex) and latex[probe].isspace():
                probe += 1
        if probe >= len(latex) or latex[probe] != "{":
            continue
        closing = _matching_brace(latex, probe)
        if closing is None:
            continue
        span = (start, closing + 1)
        if name == "label":
            removals.append(span)
        else:
            tags.append((span, latex[probe + 1:closing].strip()))
        position = closing + 1
    tag_text = None
    if len(tags) == 1:
        span, tag_text = tags[0]
        removals.append(span)
    remaining = []
    cursor = 0
    for start, end in sorted(removals):
        remaining.append(latex[cursor:start])
        cursor = end
    remaining.append(latex[cursor:])
    return "".join(remaining).strip(), tag_text
