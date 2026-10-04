"""Command handlers for delimiters.

``\\left`` ... [``\\middle`` ...] ``\\right`` pairs, the fixed-size
``\\big(`` family, and the bra-ket notation (``\\bra``, ``\\ket``,
``\\braket``).
"""

from __future__ import annotations

from typing import NoReturn

from . import parser
from .builders import _delimiter, _styled_run
from .errors import UnsupportedLatexError
from .style import _Style
from .tokenizer import (
    _MIDDLE_COMMAND,
    _RIGHT_COMMAND,
    Stop,
    _read_delimiter,
    _skip_space,
)


def _parse_left(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    r"""Parse ``\left`` ... [``\middle`` ...] ``\right``. Returns (elements, index).

    ``\middle`` splits the content into segments that OMML separates with
    one stretchy `m:sepChr`, so every ``\middle`` of a pair must use the
    same character.
    """
    begin, index = _read_delimiter(tokens, index, "left")
    segments: list = []
    separators: list = []
    while True:
        children, index = parser._parse_sequence(
            tokens, index,
            stop=lambda kind, value: (kind, value) in (_RIGHT_COMMAND,
                                                       _MIDDLE_COMMAND),
            style=style,
        )
        segments.append(children)
        if index >= len(tokens) or tokens[index] not in (_RIGHT_COMMAND,
                                                          _MIDDLE_COMMAND):
            raise UnsupportedLatexError("\\left without a matching \\right")
        if tokens[index] == _RIGHT_COMMAND:
            break
        separator, index = _read_delimiter(tokens, index + 1, "middle")
        separators.append(separator)
    end, index = _read_delimiter(tokens, index + 1, "right")
    if not separators:
        return [_delimiter(begin, end, segments[0])], index
    if len(set(separators)) > 1 or not separators[0]:
        shown = " and ".join(repr(separator or ".") for separator in separators)
        raise UnsupportedLatexError(
            "\\middle must use one and the same visible delimiter throughout "
            f"a \\left...\\right pair; got {shown}"
        )
    return [_delimiter(begin, end, *segments, separator=separators[0])], index


def _refuse_right(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> NoReturn:
    raise UnsupportedLatexError("\\right without a matching \\left")


def _refuse_middle(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        "\\middle is only allowed between \\left and \\right")


def _parse_big_delimiter(tokens: list, index: int, name: str, style: _Style,
                         stop: Stop) -> tuple:
    """``\\big(`` and its siblings: a plain character, since OMML has no
    "one size larger" fence."""
    character, index = _read_delimiter(tokens, index, name)
    return ([_styled_run(character, style, "number")] if character else []), index


def _parse_ket(tokens: list, index: int, name: str, style: _Style,
               stop: Stop) -> tuple:
    """``\\ket{a}``: |a⟩."""
    content, index = parser._parse_group(tokens, index, style, name)
    return [_delimiter("|", "⟩", content)], index


def _parse_bra(tokens: list, index: int, name: str, style: _Style,
               stop: Stop) -> tuple:
    """``\\bra{a}``: ⟨a|."""
    content, index = parser._parse_group(tokens, index, style, name)
    return [_delimiter("⟨", "|", content)], index


def _parse_braket(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    r"""``\braket{a|b}`` (braket package): ⟨a|b⟩, every top-level ``|``
    a stretchy separator."""
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe][0] != "open":
        raise UnsupportedLatexError(
            "\\braket needs a braced argument, as in \\braket{a|b}")
    index = probe + 1
    segments: list = []
    while True:
        segment, index = parser._parse_sequence(
            tokens, index, stop=lambda kind, value: (kind, value) == ("other", "|"),
            style=style)
        segments.append(segment)
        if index < len(tokens) and tokens[index] == ("other", "|"):
            index += 1
            continue
        break
    if index >= len(tokens) or tokens[index][0] != "close":
        raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
    separator = "|" if len(segments) > 1 else None
    return [_delimiter("⟨", "⟩", *segments, separator=separator)], index + 1
