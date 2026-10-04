"""Command handlers for marks over and under an expression, and boxes.

Accents, over/underlines, braces and arrows drawn over or under their
argument (``\\overbrace``, ``\\xrightarrow``), ``\\overset`` and its
relatives; frames (``\\boxed``, ``\\fbox``), strikes (``\\cancel``),
phantoms, ``\\smash`` and struts.
"""

from __future__ import annotations

from . import arguments, parser
from .builders import (
    _accent,
    _border_box,
    _group_character,
    _limit_low,
    _limit_upp,
    _overline,
    _phantom,
    _run,
)
from .errors import UnsupportedLatexError
from .style import _Style
from .tables import (
    _ACCENTS,
    _CANCELS,
    _EXTENSIBLE_ARROWS,
    _GROUP_CHARACTERS,
    _PHANTOMS,
)
from .tokenizer import Stop, _read_raw_bracket


def _parse_accent(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    """``\\hat``, ``\\vec`` and the other accents (`_ACCENTS`)."""
    base, index = parser._parse_group(tokens, index, style, name)
    return [_accent(_ACCENTS[name], base)], index


def _parse_overline(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> tuple:
    """``\\overline`` and ``\\underline``."""
    base, index = parser._parse_group(tokens, index, style, name)
    position = "top" if name == "overline" else "bot"
    return [_overline(base, position=position)], index


def _parse_group_character(tokens: list, index: int, name: str,
                           style: _Style, stop: Stop) -> tuple:
    """``\\overbrace``, ``\\underrightarrow`` and the other stretchy
    characters (`_GROUP_CHARACTERS`)."""
    character, position, annotated = _GROUP_CHARACTERS[name]
    base, index = parser._parse_group(tokens, index, style, name)
    group = [_group_character(
        base, character, position, "bot" if position == "top" else "top")]
    if not annotated:
        return group, index
    return arguments._operator_scripts(tokens, index, group, style, True)


def _parse_extensible_arrow(tokens: list, index: int, name: str,
                            style: _Style, stop: Stop) -> tuple:
    """``\\xrightarrow[below]{above}`` and its siblings."""
    character = _EXTENSIBLE_ARROWS[name]
    below, index = arguments._read_bracket_argument(tokens, index, style)
    above, index = parser._parse_group(tokens, index, style, name)
    if above and below:
        arrow = _group_character(above, character, "bot", "bot")
        return [_limit_low([arrow], below)], index
    if above:
        return [_group_character(above, character, "bot", "bot")], index
    if below:
        return [_group_character(below, character, "top", "top")], index
    return [_run(character)], index


def _parse_overset(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> tuple:
    """``\\overset``, ``\\stackrel`` and ``\\underset``: an annotation
    stacked over or under a base."""
    annotation, index = parser._parse_group(tokens, index, style, name)
    base, index = parser._parse_group(tokens, index, style, name)
    if name == "underset":
        return [_limit_low(base, annotation)], index
    return [_limit_upp(base, annotation)], index


def _parse_boxed(tokens: list, index: int, name: str, style: _Style,
                 stop: Stop) -> tuple:
    """``\\boxed{...}``: math in a frame."""
    base, index = parser._parse_group(tokens, index, style, name)
    return [_border_box(base)], index


def _parse_fbox(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    """``\\fbox{...}`` and ``\\framebox{...}``: text in a frame."""
    pieces, index = arguments._read_text_group(tokens, index, name)
    make_run = arguments._text_run_factory(None, style)
    return [_border_box(arguments._text_elements(pieces, make_run))], index


def _parse_cancel(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    """``\\cancel``, ``\\bcancel`` and ``\\xcancel``."""
    base, index = parser._parse_group(tokens, index, style, name)
    return [_border_box(base, hide=True, strikes=_CANCELS[name])], index


def _parse_phantom(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> tuple:
    """``\\phantom``, ``\\hphantom`` and ``\\vphantom``."""
    show, zero_width, zero_ascent, zero_descent = _PHANTOMS[name]
    base, index = parser._parse_group(tokens, index, style, name)
    return [_phantom(base, show=show, zero_width=zero_width,
                     zero_ascent=zero_ascent,
                     zero_descent=zero_descent)], index


def _parse_smash(tokens: list, index: int, name: str, style: _Style,
                 stop: Stop) -> tuple:
    """``\\smash``, ``\\smash[t]`` and ``\\smash[b]``."""
    option, index = _read_raw_bracket(tokens, index)
    option = (option or "").strip()
    if option not in ("", "t", "b"):
        raise UnsupportedLatexError(
            f"\\smash takes [t] or [b], not [{option}]")
    base, index = parser._parse_group(tokens, index, style, name)
    return [_phantom(base, show=True, zero_ascent=option != "b",
                     zero_descent=option != "t")], index


def _parse_strut(tokens: list, index: int, name: str, style: _Style,
                 stop: Stop) -> tuple:
    """``\\mathstrut`` and ``\\strut``: an invisible, zero-width parenthesis."""
    return [_phantom([_run("(")], show=False, zero_width=True)], index
