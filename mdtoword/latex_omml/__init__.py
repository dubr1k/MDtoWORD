"""Convert a LaTeX math string into Office Math Markup Language (OMML).

The result is a real Word equation that stays editable in Word's equation
editor, not a text fallback with dollar signs around it.

Only constructs with an entry in the command table (``commands._REGISTRY``)
or a branch in ``environments._parse_environment`` -- plus what
``parser._parse_lines`` has to handle itself because it acts on the
sequence around it rather than on one atom: the infix commands, the ``\\``
line break, the ``&`` alignment point, the declaration switches
(``\\color``, ``\\bf``, ``\\displaystyle``...) and the equation-numbering
commands (``\\tag``, ``\\label``, ``\\nonumber``, ``\\notag``) -- are
understood.  Anything else raises :class:`UnsupportedLatexError` naming the
offending construct, so the caller can report it instead of silently
emitting wrong output.

Package layout, leaves first:

``errors``
    :class:`UnsupportedLatexError`.
``symbols``
    Character tables: symbols, spacing, escapes, delimiters.
``tables``
    Command and environment families: operators, accents, alphabets,
    colours, matrix and line environments.
``style``
    `_Style`, the math alphabet runs are written in, and the font switches.
``tokenizer``
    The tokenizer and the token-level readers that never parse math (raw
    groups, delimiters, ``\\limits``, row spacing, lengths).
``builders``
    OMML element constructors.
``colors``
    Colour specifications, and colouring OMML already built.
``parser``
    The recursive descent: lines, sequences, groups, atoms.
``arguments``
    Math and text-mode argument readers, and ``_``/``^`` scripts.
``commands``
    The ``\\command`` dispatch table, in precedence order.
``commands_*``
    The command handlers, grouped by concern: ``_alphabets`` (alphabets,
    text, plain characters), ``_operators`` (fractions, roots, big and
    named operators), ``_decorations`` (marks over and under, boxes,
    phantoms), ``_delimiters`` (``\\left``/``\\right``, sizes, bra-ket)
    and ``_spacing`` (spaces and ``\\textcolor``).
``environments``
    ``\\begin{...}`` ... ``\\end{...}``: matrices, cases, arrays and the
    line-by-line environments.
``tags``
    `split_equation_tag`: equation numbering at the source level.

``parser``, ``arguments``, ``environments`` and the ``commands*`` modules
call one another recursively, so among themselves they import modules
(``from . import parser``) and look functions up at call time; everything
else is imported by name.
"""

from __future__ import annotations

from typing import Any

# `parser` must be the first module of the recursive group imported here
# -- never a handler module -- so that `commands`, which it imports, builds
# its dispatch table only after every handler module has finished loading.
from .parser import _parse_sequence
from .builders import _el
from .errors import UnsupportedLatexError
from .tags import split_equation_tag
from .tokenizer import _tokenize

# Tables the test-suite inspects directly.
from .symbols import (
    _ESCAPED as _ESCAPED,
    _SPACING as _SPACING,
    _SYMBOLS as _SYMBOLS,
)
from .tables import (
    _ACCENTS as _ACCENTS,
    _ENVIRONMENT_ONLY as _ENVIRONMENT_ONLY,
    _LIMIT_OPERATORS as _LIMIT_OPERATORS,
    _NARY as _NARY,
    _UPRIGHT_FUNCTIONS as _UPRIGHT_FUNCTIONS,
)

__all__ = [
    "UnsupportedLatexError",
    "latex_to_omml",
    "omml_children",
    "split_equation_tag",
]


def omml_children(latex: str) -> list[Any]:
    """Parse a LaTeX math string into a list of OMML elements."""
    tokens = _tokenize(latex)
    elements, index = _parse_sequence(tokens, 0, stop=None)
    if index != len(tokens):
        raise UnsupportedLatexError(f"Could not parse the whole formula: {latex!r}")
    return elements


def latex_to_omml(latex: str) -> Any:
    """Parse a LaTeX math string into a single <m:oMath> element."""
    math = _el("oMath")
    for child in omml_children(latex):
        math.append(child)
    return math
