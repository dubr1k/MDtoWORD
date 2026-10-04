"""Command handlers for spacing and colour.

Explicit horizontal space (``\\hspace``, ``\\kern`` and its relatives),
the named spacing commands (``\\,``, ``\\quad``...) and ``\\textcolor``.
"""

from __future__ import annotations

from typing import Any

from . import parser
from .builders import _run
from .colors import _colorize, _read_color
from .errors import UnsupportedLatexError
from .style import _Style
from .symbols import _SPACING
from .tokenizer import Stop, _parse_length, _read_raw_group, _skip_space


def _space_run(ems: float) -> list[Any]:
    """Space about `ems` wide, in the units this module spaces with: one
    space per em, as ``\\quad`` and ``\\qquad`` give.  Negative space has
    no OMML equivalent and, like ``\\!``, becomes nothing."""
    if ems <= 0:
        return []
    return [_run(" " * min(50, max(1, round(ems))))]


def _parse_hspace(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    """``\\hspace{1em}`` and ``\\hspace*{1em}``."""
    probe = _skip_space(tokens, index)
    if probe < len(tokens) and tokens[probe] == ("other", "*"):
        index = probe + 1
    text, index = _read_raw_group(tokens, index, name, verbatim=True)
    ems = _parse_length(text)
    if ems is None:
        raise UnsupportedLatexError(
            f"\\hspace needs an explicit length such as \\hspace{{1em}}, "
            f"not {text!r}"
        )
    return _space_run(ems), index


def _parse_kern(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    r"""``\kern3pt``, ``\mkern-2mu``, ``\hskip 1em``, ``\mspace{3mu}``:
    an explicit horizontal space, braced or not."""
    probe = _skip_space(tokens, index)
    if probe < len(tokens) and tokens[probe][0] == "open":
        text, index = _read_raw_group(tokens, probe, name, verbatim=True)
    else:
        # Unbraced: an optional sign, a number, then a two-letter unit.
        pieces = []
        position = probe
        while (position < len(tokens) and tokens[position][0] == "other"
               and tokens[position][1] in "+-."):
            pieces.append(tokens[position][1])
            position = _skip_space(tokens, position + 1)
        if position < len(tokens) and tokens[position][0] == "number":
            pieces.append(tokens[position][1])
            position = _skip_space(tokens, position + 1)
        for _ in range(2):
            if position < len(tokens) and tokens[position][0] == "letter":
                pieces.append(tokens[position][1])
                position += 1
        text = "".join(pieces)
        index = position
    ems = _parse_length(text)
    if ems is None:
        raise UnsupportedLatexError(
            f"\\{name} needs an explicit length such as 3mu or 1em"
        )
    return _space_run(ems), index


def _parse_spacing(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> tuple:
    """``\\,``, ``\\quad`` and the other named spaces (`_SPACING`)."""
    spacing = _SPACING[name]
    if not spacing:
        return [], index
    return [_run(spacing)], index


def _parse_textcolor(tokens: list, index: int, name: str, style: _Style,
                     stop: Stop) -> tuple:
    """``\\textcolor{red}{...}``: the argument in one colour."""
    color, index = _read_color(tokens, index, name)
    content, index = parser._parse_group(tokens, index, style, name)
    _colorize(content, color)
    return content, index
