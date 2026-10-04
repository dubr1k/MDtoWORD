"""Tokenizer and the token-level readers that never parse math.

Besides splitting the source into (kind, value) tokens, this holds the
readers that only glue raw token values back together or look a few
tokens ahead -- raw ``{...}``/``[...]`` arguments, the fence after
``\\left``, ``\\limits`` modifiers, the spacing after ``\\\\`` and TeX
lengths -- so they can be used anywhere without touching the parser.
"""

from __future__ import annotations

import re
from typing import Callable, Optional

from .errors import UnsupportedLatexError
from .symbols import (
    _DELIMITER_CHARACTERS,
    _DELIMITER_COMMANDS,
    _ESCAPED,
    _SPACING,
)

_ROW_SEPARATOR = ("command", "\\\\")
_END_COMMAND = ("command", "\\end")
_RIGHT_COMMAND = ("command", "\\right")
_MIDDLE_COMMAND = ("command", "\\middle")
_DOLLAR = ("other", "$")


# TeX lengths, in points per unit.  An em is taken as 10pt, which is what
# a 10pt LaTeX document's math font gives it.
_LENGTH_UNITS = {
    "pt": 1.0, "pc": 12.0, "in": 72.27, "bp": 72.27 / 72, "cm": 28.4528,
    "mm": 2.84528, "dd": 1238 / 1157, "cc": 14856 / 1157, "sp": 1 / 65536,
    "em": 10.0, "ex": 4.3, "mu": 10.0 / 18, "px": 72.27 / 96,
}
_LENGTH_RE = re.compile(
    r"^\s*(?P<value>[-+]?\s*(?:\d+(?:\.\d*)?|\.\d+))\s*"
    r"(?P<unit>pt|pc|in|bp|cm|mm|dd|cc|sp|em|ex|mu|px)"
    r"(?:\s*(?:plus|minus)\s*[-+]?\s*(?:\d+(?:\.\d*)?|\.\d+)\s*"
    r"(?:pt|pc|in|bp|cm|mm|dd|cc|sp|em|ex|mu|px|fil+))*\s*$"
)

_TOKEN_RE = re.compile(
    r"""
    (?P<command>\\[A-Za-z]+ | \\.)     # \frac, \alpha, \\, \{, \,
  | (?P<number>[0-9]+(?:\.[0-9]+)?)
  | (?P<letter>[A-Za-z])
  | (?P<open>\{) | (?P<close>\})
  | (?P<sup>\^) | (?P<sub>_)
  | (?P<bracket>\[|\])
  | (?P<amp>&)
  | (?P<space>\s+)
  | (?P<other>[^\s])
    """,
    re.VERBOSE,
)

Stop = Optional[Callable[[str, str], bool]]


def _tokenize(latex: str) -> list:
    """Split `latex` into (kind, value) pairs and check brace balance.

    Whitespace is kept as `space` tokens: the math parser skips them, but
    `_read_raw_group` needs them so `\\text{если да}` keeps its space.
    """
    tokens: list = []
    depth = 0
    position = 0
    for match in _TOKEN_RE.finditer(latex):
        if match.start() != position:
            gap = latex[position:match.start()]
            raise UnsupportedLatexError(
                f"Could not read this part of the formula: {gap!r}"
            )
        position = match.end()
        kind = match.lastgroup
        value = match.group()
        if kind == "open":
            depth += 1
        elif kind == "close":
            depth -= 1
            if depth < 0:
                raise UnsupportedLatexError(
                    f"Unbalanced braces: unexpected '}}' in {latex!r}"
                )
        tokens.append((kind, value))
    if position != len(latex):
        raise UnsupportedLatexError(
            f"Could not read this part of the formula: {latex[position:]!r}"
        )
    if depth > 0:
        raise UnsupportedLatexError(f"Unbalanced braces: unclosed '{{' in {latex!r}")
    return tokens


def _skip_space(tokens: list, index: int) -> int:
    """Index of the next non-whitespace token at or after `index`."""
    while index < len(tokens) and tokens[index][0] == "space":
        index += 1
    return index


def _parse_length(text: str) -> Optional[float]:
    """A TeX length such as ``1.5em`` or ``-3mu``, in ems; None if `text`
    is not one.  Glue (``plus``/``minus`` stretch) is accepted and ignored."""
    match = _LENGTH_RE.match(text)
    if match is None:
        return None
    value = float(match.group("value").replace(" ", ""))
    return value * _LENGTH_UNITS[match.group("unit")] / _LENGTH_UNITS["em"]


def _skip_row_spacing(tokens: list, index: int) -> int:
    r"""Skip what may follow a ``\\``: a ``*`` and a ``[2pt]`` spacing.

    Only a bracket that holds a real length is taken -- ``\\ [a, b)`` is a
    line that starts with a bracket, not an argument.
    """
    if index < len(tokens) and tokens[index] == ("other", "*"):
        index += 1
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe] != ("bracket", "["):
        return index
    close = probe + 1
    while close < len(tokens) and tokens[close] != ("bracket", "]"):
        if tokens[close][0] not in ("number", "letter", "other", "space"):
            return index
        close += 1
    if close >= len(tokens):
        return index
    text = "".join(value for _, value in tokens[probe + 1:close])
    if _parse_length(text) is None:
        return index
    return close + 1


def _read_raw_group(tokens: list, index: int, owner: str = "",
                    verbatim: bool = False) -> tuple:
    """Read a `{...}` group as plain characters instead of as math.

    The tokenizer splits words letter by letter, so `\\begin{pmatrix}` and
    `\\operatorname{sgn}` both need the raw values glued back together
    rather than parsed.  Returns (text, index_after_group).  `owner` names
    the command the group belongs to, for error messages; `verbatim` keeps
    commands as their source text instead of refusing them (a `\\label`
    key is never displayed, so it may hold anything).
    """
    index = _skip_space(tokens, index)
    if index >= len(tokens) or tokens[index][0] != "open":
        if owner:
            raise UnsupportedLatexError(
                f"\\{owner} needs a braced argument, as in \\{owner}{{...}}"
            )
        raise UnsupportedLatexError("Expected '{' here")
    index += 1
    depth = 1
    pieces: list = []
    while index < len(tokens):
        kind, value = tokens[index]
        if kind == "open":
            depth += 1
        elif kind == "close":
            depth -= 1
            if depth == 0:
                return "".join(pieces), index + 1
        if kind == "command" and not verbatim:
            name = value[1:]
            if name in _ESCAPED:
                value = _ESCAPED[name]
            elif name in _SPACING:
                value = _SPACING[name]
            else:
                raise UnsupportedLatexError(
                    f"Commands are not supported inside text: \\{name}"
                )
        pieces.append(value)
        index += 1
    raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")


def _read_raw_bracket(tokens: list, index: int) -> tuple:
    """Read an optional `[...]` argument as plain characters.

    Returns (text, index) or (None, index) when no bracket follows.
    """
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe] != ("bracket", "["):
        return None, index
    close = probe + 1
    while close < len(tokens) and tokens[close] != ("bracket", "]"):
        close += 1
    if close >= len(tokens):
        raise UnsupportedLatexError("Unclosed '[' in an optional argument")
    return "".join(value for _, value in tokens[probe + 1:close]), close + 1


def _read_limits_modifier(tokens: list, index: int) -> tuple:
    r"""Consume ``\limits`` / ``\nolimits`` / ``\displaylimits`` after an
    operator.  Returns (placement, index): "limits", "nolimits", or None
    for the operator's own default.  As in TeX, the last one wins."""
    placement = None
    while True:
        probe = _skip_space(tokens, index)
        if probe >= len(tokens) or tokens[probe][0] != "command":
            return placement, index
        name = tokens[probe][1][1:]
        if name not in ("limits", "nolimits", "displaylimits"):
            return placement, index
        placement = None if name == "displaylimits" else name
        index = probe + 1


def _read_delimiter(tokens: list, index: int, command: str) -> tuple:
    """Read the fence character that follows `\\left`, `\\right`,
    `\\middle` or a `\\big` size command.

    Returns (character, index); the character is "" for `.`, LaTeX's
    "there is no fence on this side".
    """
    index = _skip_space(tokens, index)
    if index >= len(tokens):
        raise UnsupportedLatexError(
            f"\\{command} is missing the delimiter that should follow it"
        )
    kind, value = tokens[index]
    if kind == "command":
        candidate = _DELIMITER_COMMANDS.get(value[1:])
        if candidate is None:
            raise UnsupportedLatexError(
                f"Not a delimiter after \\{command}: {value}"
            )
        return candidate, index + 1
    if value in _DELIMITER_CHARACTERS:
        return _DELIMITER_CHARACTERS[value], index + 1
    raise UnsupportedLatexError(f"Not a delimiter after \\{command}: {value!r}")
