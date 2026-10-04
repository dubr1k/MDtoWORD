"""Command handlers for alphabets, text and plain characters.

Math alphabets (``\\mathbf``, ``\\mathbb``, ``\\boldsymbol``...), the
spacing-class wrappers (``\\mathrel``...), text-mode commands
(``\\text``...) and the commands that stand for a single character
(``\\alpha``, ``\\varGamma``, ``\\%``).
"""

from __future__ import annotations

from . import arguments, parser
from .builders import _styled_run
from .style import _MATH_ALPHABETS, _Style
from .symbols import _ESCAPED, _ITALIC_SYMBOLS, _SYMBOLS
from .tables import _SCRIPT_ALPHABETS, _TEXT_COMMANDS
from .tokenizer import Stop


def _parse_text(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    """``\\text{...}`` and its siblings: literal text (`_TEXT_COMMANDS`)."""
    pieces, index = arguments._read_text_group(tokens, index, name)
    return arguments._text_elements(
        pieces, arguments._text_run_factory(_TEXT_COMMANDS[name], style)), index


def _parse_math_class(tokens: list, index: int, name: str, style: _Style,
                      stop: Stop) -> tuple:
    """``\\mathrel{...}`` and the other spacing classes: just the content."""
    return parser._parse_group(tokens, index, style, name)


def _parse_math_alphabet(tokens: list, index: int, name: str, style: _Style,
                         stop: Stop) -> tuple:
    """``\\mathbf``, ``\\mathit``, ``\\mathbfit``, ``\\mathrm``/``\\mathup``
    and ``\\mathnormal``: the argument in one fixed alphabet."""
    return parser._parse_group(tokens, index, _MATH_ALPHABETS[name], name)


def _parse_script_alphabet(tokens: list, index: int, name: str,
                           style: _Style, stop: Stop) -> tuple:
    """``\\mathbb``, ``\\mathcal`` and friends: an `<m:scr>` alphabet,
    bold when the surrounding style is."""
    alphabet = _Style("bold" if style.bold else None, _SCRIPT_ALPHABETS[name])
    return parser._parse_group(tokens, index, alphabet, name)


def _parse_bold_italic(tokens: list, index: int, name: str, style: _Style,
                       stop: Stop) -> tuple:
    """``\\boldsymbol``, ``\\bm`` and ``\\pmb``: the surrounding alphabet,
    made bold."""
    if style.script is not None:
        heavy = _Style("bold", style.script)
    elif style.face in ("bold", "upright"):
        heavy = _Style("bold")
    else:
        heavy = _Style("bolditalic")
    return parser._parse_group(tokens, index, heavy, name)


def _parse_symbol(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    """A command that stands for one character, such as ``\\alpha``."""
    return [_styled_run(_SYMBOLS[name], style, "symbol")], index


def _parse_italic_symbol(tokens: list, index: int, name: str, style: _Style,
                         stop: Stop) -> tuple:
    """``\\varGamma`` and the other italic capital Greek letters."""
    return [_styled_run(_ITALIC_SYMBOLS[name], style, "letter")], index


def _parse_escaped(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> tuple:
    """``\\%``, ``\\{`` and the other escaped syntax characters."""
    return [_styled_run(_ESCAPED[name], style, "number")], index
