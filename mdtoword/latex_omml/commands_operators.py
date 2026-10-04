"""Command handlers for operators and functions.

Fractions, roots, binomials and ``\\substack``; big operators (``\\sum``,
``\\int``...), limit-style operators (``\\lim``), ``\\operatorname`` and
``\\mathop``, the upright function names (``\\sin``); ``\\mod`` and its
relatives, and ``\\not``.
"""

from __future__ import annotations

import unicodedata
from typing import NoReturn

from docx.oxml.ns import qn

from . import arguments, parser
from .builders import _delimiter, _fraction, _matrix, _nary, _radical, _run
from .errors import UnsupportedLatexError
from .style import _Style
from .symbols import _SPACING
from .tables import _LIMIT_OPERATORS, _NARY
from .tokenizer import (
    Stop,
    _read_limits_modifier,
    _read_raw_bracket,
    _read_raw_group,
    _skip_space,
)


def _parse_fraction(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> tuple:
    """``\\frac``, ``\\dfrac``, ``\\tfrac`` and ``\\cfrac``."""
    if name == "cfrac":
        # `\cfrac[l]` only moves the numerator, which an OMML fraction
        # always centres.
        option, index = _read_raw_bracket(tokens, index)
        if option is not None and option.strip() not in ("l", "c", "r"):
            raise UnsupportedLatexError(
                f"\\cfrac takes [l], [c] or [r], not [{option}]")
    numerator, index = parser._parse_group(tokens, index, style, name)
    denominator, index = parser._parse_group(tokens, index, style, name)
    return [_fraction(numerator, denominator)], index


def _parse_sqrt(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    """``\\sqrt{x}`` and ``\\sqrt[n]{x}``."""
    degree, index = arguments._read_bracket_argument(tokens, index, style)
    if not degree:
        degree = None
    radicand, index = parser._parse_group(tokens, index, style, name)
    return [_radical(degree, radicand)], index


def _parse_binomial(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> tuple:
    """``\\binom`` and its display/text variants: a barless fraction in
    parentheses."""
    top, index = parser._parse_group(tokens, index, style, name)
    bottom, index = parser._parse_group(tokens, index, style, name)
    return [_delimiter("(", ")", [_fraction(top, bottom, no_bar=True)])], index


def _parse_substack(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> tuple:
    """``\\substack{a \\\\ b}``: lines stacked in one column."""
    # One column, one line per `\\` -- the shape a stacked n-ary limit
    # such as `\sum_{\substack{i < j \\ i \in S}}` needs.
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe][0] != "open":
        raise UnsupportedLatexError(
            "\\substack needs a brace group, as in \\substack{a \\\\ b}"
        )
    lines, index = parser._parse_lines(tokens, probe + 1, style=style)
    if index >= len(tokens) or tokens[index][0] != "close":
        raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
    return [_matrix([[line] for line in lines])], index + 1


def _parse_operatorname(tokens: list, index: int, name: str, style: _Style,
                        stop: Stop) -> tuple:
    """``\\operatorname{sgn}`` and the limit-taking ``\\operatorname*``."""
    # `\operatorname*` is the limit-taking form, like `\lim`.
    probe = _skip_space(tokens, index)
    starred = probe < len(tokens) and tokens[probe] == ("other", "*")
    if starred:
        index = probe + 1
    text, index = _read_raw_group(
        tokens, index, "operatorname*" if starred else "operatorname")
    base = [_run(text, upright=True, bold=style.bold)]
    return arguments._operator_scripts(tokens, index, base, style, starred)


def _parse_mathop(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> tuple:
    """``\\mathop{...}``: any content, taking limits like an operator."""
    base, index = parser._parse_group(tokens, index, style, name)
    return arguments._operator_scripts(tokens, index, base, style, True)


def _parse_nary(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    """``\\sum``, ``\\int`` and the other big operators (`_NARY`)."""
    # The limits bind to the operator itself; everything after them, up
    # to the end of the enclosing construct -- or the next line break,
    # whichever comes first -- is the operand.
    placement, index = _read_limits_modifier(tokens, index)
    location = {"limits": "undOvr", "nolimits": "subSup"}.get(placement)
    sub, sup, index = arguments._read_scripts(tokens, index, style)
    body, index = parser._parse_sequence(
        tokens, index, stop=parser._stop_at_segment_end(stop), style=style)
    return [_nary(_NARY[name], sub, sup, body, location)], index


def _parse_limit_operator(tokens: list, index: int, name: str, style: _Style,
                          stop: Stop) -> tuple:
    """``\\lim``, ``\\max`` and the other limit-style names."""
    base = [_run(_LIMIT_OPERATORS[name], upright=True, bold=style.bold)]
    return arguments._operator_scripts(tokens, index, base, style, True)


def _parse_upright_function(tokens: list, index: int, name: str,
                            style: _Style, stop: Stop) -> tuple:
    """``\\sin``, ``\\log`` and the other function names: upright text with
    scripts beside it."""
    return arguments._operator_scripts(
        tokens, index, [_run(name, upright=True, bold=style.bold)], style, False)


def _parse_bmod(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> tuple:
    """``\\bmod``: the binary "mod" between two operands."""
    return [_run(" "), _run("mod", upright=True), _run(" ")], index


def _parse_mod(tokens: list, index: int, name: str, style: _Style,
               stop: Stop) -> tuple:
    """``\\mod{n}``, ``\\pmod{n}`` and ``\\pod{n}``."""
    modulus, index = parser._parse_group(tokens, index, style, name)
    if name == "mod":
        return [_run(_SPACING["quad"]), _run("mod", upright=True),
                _run(" "), *modulus], index
    inside = modulus
    if name == "pmod":
        inside = [_run("mod", upright=True), _run(" "), *modulus]
    return [_run(_SPACING["quad"]), _delimiter("(", ")", inside)], index


def _parse_negation(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> tuple:
    r"""``\not`` before a relation: the precomposed negated character
    where Unicode has one (``\not=`` is ≠, ``\not\in`` is ∉), otherwise the
    relation with U+0338 COMBINING LONG SOLIDUS OVERLAY -- which is exactly
    what Unicode canonical composition (NFC) of the two gives."""
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe][0] not in (
            "command", "other", "letter", "bracket"):
        raise UnsupportedLatexError(
            "\\not needs a relation after it, as in \\not\\in or \\not=")
    atom, index = parser._parse_atom(tokens, probe, style, stop)
    text_element = atom[0].find(qn("m:t")) if len(atom) == 1 else None
    if (text_element is None or atom[0].tag != qn("m:r")
            or len(text_element.text or "") != 1):
        raise UnsupportedLatexError(
            "\\not can only negate a single relation symbol, as in \\not\\in")
    text_element.text = unicodedata.normalize("NFC", text_element.text + "̸")
    return atom, index


def _refuse_sideset(tokens: list, index: int, name: str, style: _Style,
                    stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        "\\sideset is not supported: OMML cannot attach scripts to both "
        "sides of a big operator; write the operator with ordinary limits"
    )


def _refuse_limits(tokens: list, index: int, name: str, style: _Style,
                   stop: Stop) -> NoReturn:
    """``\\limits`` and friends anywhere but straight after an operator."""
    raise UnsupportedLatexError(
        f"\\{name} must follow a big operator such as \\sum or \\int"
    )
