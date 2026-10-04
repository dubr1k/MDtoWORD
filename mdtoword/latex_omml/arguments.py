"""Argument readers that parse math or text mode.

Text-mode groups (with ``$...$`` math inside them), optional ``[...]``
math arguments and the ``_``/``^`` scripts after a base or an operator.
They call back into :mod:`.parser` for the math they contain.
"""

from __future__ import annotations

from typing import Any, Callable, Optional

from . import parser
from .builders import _limits, _run, _word_format
from .errors import UnsupportedLatexError
from .style import _PLAIN, _Style
from .symbols import _ESCAPED, _SPACING, _TEXT_SYMBOLS
from .tokenizer import _DOLLAR, Stop, _read_limits_modifier, _skip_space


def _read_text_material(tokens: list, index: int, owner: str,
                        stop: Stop = None) -> tuple:
    r"""Read text-mode material up to an unmatched ``}`` or `stop`.

    Returns (pieces, index) with `index` left ON the terminator.  Each
    piece is ("text", str) or ("math", elements): ``$...$`` and
    ``\(...\)`` switch back to math inside text, as they do in LaTeX.
    Braces only group here, so ``\text{a{b}c}`` reads "abc".
    """
    pieces: list = []
    buffer: list = []

    def flush() -> None:
        if buffer:
            pieces.append(("text", "".join(buffer)))
            buffer.clear()

    depth = 0
    while index < len(tokens):
        kind, value = tokens[index]
        if kind == "open":
            depth += 1
            index += 1
            continue
        if kind == "close":
            if depth == 0:
                break
            depth -= 1
            index += 1
            continue
        if depth == 0 and stop is not None and stop(kind, value):
            break
        if (kind, value) in (_DOLLAR, ("command", "\\(")):
            closer = _DOLLAR if (kind, value) == _DOLLAR else ("command", "\\)")
            end = index + 1
            while end < len(tokens) and tokens[end] != closer:
                end += 1
            if end >= len(tokens):
                raise UnsupportedLatexError(
                    f"Math inside \\{owner} is not closed: {value!r} has no "
                    "partner"
                )
            inner = list(tokens[index + 1:end])
            elements, consumed = parser._parse_sequence(inner, 0)
            if consumed != len(inner):
                raise UnsupportedLatexError(
                    f"Could not parse the math inside \\{owner}"
                )
            flush()
            pieces.append(("math", elements))
            index = end + 1
            continue
        if kind == "command":
            name = value[1:]
            if name in _ESCAPED:
                buffer.append(_ESCAPED[name])
            elif name in _SPACING:
                buffer.append(_SPACING[name])
            elif name in _TEXT_SYMBOLS:
                buffer.append(_TEXT_SYMBOLS[name])
            else:
                raise UnsupportedLatexError(
                    f"Commands are not supported inside text: \\{name}"
                )
            index += 1
            continue
        buffer.append(" " if (kind, value) == ("other", "~") else value)
        index += 1
    flush()
    return pieces, index


def _read_text_group(tokens: list, index: int, owner: str) -> tuple:
    """Read a braced text-mode argument. Returns (pieces, index_after_group)."""
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe][0] != "open":
        raise UnsupportedLatexError(
            f"\\{owner} needs a braced argument, as in \\{owner}{{...}}"
        )
    pieces, index = _read_text_material(tokens, probe + 1, owner)
    if index >= len(tokens) or tokens[index][0] != "close":
        raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
    return pieces, index + 1


def _text_elements(pieces: list, make_run: Callable[[str], Any]) -> list:
    """Turn text-mode pieces into OMML: `make_run` for each text piece,
    the parsed elements for each math piece."""
    elements: list = []
    for kind, value in pieces:
        if kind == "text":
            elements.append(make_run(value))
        else:
            elements.extend(value)
    return elements


def _text_run_factory(variant: Optional[str], style: _Style) -> Callable[[str], Any]:
    """How `\\text`-family command `variant` writes its literal text."""
    if variant in ("sans-serif", "monospace"):
        return lambda text: _run(text, script=variant, bold=style.bold)
    if variant == "bold":
        return lambda text: _word_format(_run(text, upright=True), bold=True)
    if variant == "italic":
        return lambda text: _word_format(_run(text, upright=True), italic=True)
    return lambda text: _run(text, upright=True, bold=style.bold)


def _read_bracket_argument(tokens: list, index: int,
                           style: _Style = _PLAIN) -> tuple:
    """Read an optional `[...]` argument, as used by `\\sqrt[3]{8}`.

    Returns (elements, index) or (None, index) when no bracket follows.
    """
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe] != ("bracket", "["):
        return None, index
    elements, probe = parser._parse_sequence(
        tokens, probe + 1, stop=lambda k, v: k == "bracket" and v == "]", style=style
    )
    if probe >= len(tokens) or tokens[probe] != ("bracket", "]"):
        raise UnsupportedLatexError("Unclosed '[' in an optional argument")
    return elements, probe + 1


def _read_scripts(tokens: list, index: int, style: _Style = _PLAIN) -> tuple:
    """Read any `_`/`^` groups that follow, in either order.

    Returns (sub, sup, index); each script is None when absent.  `x_i^2` and
    `x^2_i` both give a sub and a sup.
    """
    sub = None
    sup = None
    while True:
        probe = _skip_space(tokens, index)
        if probe >= len(tokens):
            return sub, sup, index
        kind = tokens[probe][0]
        if kind == "sub":
            if sub is not None:
                raise UnsupportedLatexError("Two subscripts on one base")
            sub, index = parser._parse_group(tokens, probe + 1, style)
        elif kind == "sup":
            if sup is not None:
                raise UnsupportedLatexError("Two superscripts on one base")
            sup, index = parser._parse_group(tokens, probe + 1, style)
        else:
            return sub, sup, index


def _operator_scripts(tokens: list, index: int, base: list,
                      style: _Style, limits_default: bool) -> tuple:
    r"""Attach the scripts that follow an operator name such as ``\lim``.

    Stacked as limits (<m:limLow>/<m:limUpp>) when the operator takes
    limits -- `limits_default`, or an explicit ``\limits`` -- and as
    ordinary scripts otherwise, which `_parse_lines` adds itself.
    """
    placement, index = _read_limits_modifier(tokens, index)
    stacked = placement == "limits" or (limits_default and placement != "nolimits")
    if not stacked:
        return base, index
    sub, sup, index = _read_scripts(tokens, index, style)
    return _limits(base, sub, sup), index
