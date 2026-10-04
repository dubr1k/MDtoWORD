"""Characters and run text to math-mode LaTeX, and joining LaTeX fragments."""

from __future__ import annotations

import re
import unicodedata

from .symbols import (
    _ALPHANUMERIC_STYLES,
    _FUNCTIONS,
    _LITERALS,
    _MULTIWORD_FUNCTIONS,
    _RELATIONS,
    _SYMBOLS,
    _TEXT_ESCAPES,
)

_CONTROL_WORD_END = re.compile(r"\\[A-Za-z]+$")
_SIMPLE_ATOM = re.compile(
    r"\\[A-Za-z]+|\\[^A-Za-z\s]|[0-9]+(?:\.[0-9]+)?|[^\s{}\\^_&%#$~]"
)


def _join(pieces: list[str], *, separate_numbers: bool = True) -> str:
    """Concatenate LaTeX fragments, adding a space only where TeX needs one.

    ``\\alpha`` followed by ``x`` must not become ``\\alphax``, and two
    numbers from separate runs must not fuse into one when read back.
    """
    out = ""
    for piece in pieces:
        if not piece:
            continue
        if out and (
            (_CONTROL_WORD_END.search(out) and piece[0].isalnum())
            or (separate_numbers and out[-1].isdigit() and piece[0].isdigit())
        ):
            out += " "
        out += piece
    return re.sub(" {2,}", " ", out)


def _braced_base(latex: str) -> str:
    """A script base: bare when it is one token, braced otherwise."""
    if _SIMPLE_ATOM.fullmatch(latex):
        return latex
    return "{" + latex + "}"


def _escape_text_mode(text: str) -> str:
    return "".join(_TEXT_ESCAPES.get(ch, ch) for ch in text)


def _function_command(text: str) -> str | None:
    """``\\sin`` for the text ``sin``, or ``None`` if it is no known function."""
    word = text.strip()
    if word in _FUNCTIONS:
        return "\\" + word
    if word in _MULTIWORD_FUNCTIONS:
        return "\\" + _MULTIWORD_FUNCTIONS[word]
    return None


def _alphanumeric(ch: str) -> tuple[str | None, str] | None:
    """Split a Unicode mathematical alphanumeric into (style command, base)."""
    name = unicodedata.name(ch, "")
    if not name or not any(
        marker in name
        for marker in ("MATHEMATICAL", "DOUBLE-STRUCK", "SCRIPT CAPITAL",
                       "SCRIPT SMALL", "BLACK-LETTER", "PLANCK CONSTANT")
    ):
        return None
    base = unicodedata.normalize("NFKC", ch)
    if base == ch or len(base) != 1:
        return None
    for marker, command in _ALPHANUMERIC_STYLES:
        if marker in name:
            return command, base
    return None, base


def _char_latex(ch: str) -> tuple[str | None, str]:
    """LaTeX for one character, as (style wrapper or None, body)."""
    if ch in _SYMBOLS:
        command = "\\" + _SYMBOLS[ch]
        return None, (f" {command} " if ch in _RELATIONS else command)
    if ch in _RELATIONS:
        return None, f" {ch} "
    if ch in _LITERALS:
        return None, _LITERALS[ch]
    if ch == " ":
        return None, r"\ "
    styled = _alphanumeric(ch)
    if styled is not None:
        style, base = styled
        _, body = _char_latex(base)
        return style, body
    return None, ch


def _text_latex(text: str) -> str:
    """Math-mode LaTeX for plain run text, grouping styled characters."""
    groups: list[tuple[str | None, list[str]]] = []
    for ch in text:
        style, body = _char_latex(ch)
        if groups and groups[-1][0] == style:
            groups[-1][1].append(body)
        else:
            groups.append((style, [body]))
    pieces = []
    for style, bodies in groups:
        body = _join(bodies, separate_numbers=False)
        pieces.append(f"\\{style}{{{body.strip()}}}" if style else body)
    return _join(pieces, separate_numbers=False)
