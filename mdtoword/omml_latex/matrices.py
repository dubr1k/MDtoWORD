"""Delimiters, matrices and equation arrays.

An ``m:d`` delimiter becomes ``\\left ... \\right``, unless what it wraps is
really ``\\binom``, a matrix environment (``pmatrix``, ``bmatrix`` ...) or
``cases``. Bare matrices become ``matrix`` or ``array``, equation arrays
``aligned`` or ``gathered``.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .names import _VAL, _local, _m
from .symbols import _COLUMN_LETTERS, _DELIMITERS, _MATRIX_ENVIRONMENTS
from .text import _join, _text_latex

if TYPE_CHECKING:
    from .converter import _Converter

_ALIGN_MARK = re.compile(r"\s*(?<!\\)&\s*")
_ALIGN_REPLACEMENT = " &"


def equation_environment(rows: list[str]) -> str:
    """``aligned`` when any row has an alignment point, else ``gathered``.

    Both stack their rows; only ``aligned`` lines them up on ``&``, which is
    exactly the difference between an ``m:eqArr`` with and without
    ``m:aln`` markers.
    """
    return "aligned" if any(_ALIGN_MARK.search(row) for row in rows) else "gathered"


# -- delimiters -----------------------------------------------------------------


def delimiter(converter: _Converter, element: Any) -> tuple[str, str]:
    begin_el = converter.prop(element, "dPr", "begChr")
    end_el = converter.prop(element, "dPr", "endChr")
    begin = "(" if begin_el is None else begin_el.get(_VAL, "")
    end = ")" if end_el is None else end_el.get(_VAL, "")
    separator = converter.prop_value(element, "dPr", "sepChr", "|")
    operands = element.findall(_m("e"))
    if len(operands) == 1:
        special = _special_delimiter(converter, operands[0], begin, end)
        if special is not None:
            return special, "atom"
    middle = _DELIMITERS.get(separator or "")
    separator_latex = (
        f" \\middle{middle} " if middle and middle != "." else f" {_text_latex(separator or '')} "
    )
    inner = separator_latex.join(converter.seq(operand) for operand in operands)
    left = _DELIMITERS.get(begin)
    right = _DELIMITERS.get(end)
    # A fence LaTeX cannot stretch is kept as an ordinary character.
    if left is None:
        inner, left = _join([_text_latex(begin), inner]), "."
    if right is None:
        inner, right = _join([inner, _text_latex(end)]), "."
    return _join([r"\left" + left, inner, r"\right" + right]), "atom"


def _special_delimiter(converter: _Converter, operand: Any, begin: str, end: str) -> str | None:
    """\\binom, matrix environments and cases hidden behind an m:d."""
    children = converter.children(operand)
    if len(children) != 1:
        return None
    only = children[0]
    if (only.tag == _m("f") and (begin, end) == ("(", ")")
            and converter.prop_value(only, "fPr", "type") == "noBar"):
        top = converter.seq(only.find(_m("num")))
        bottom = converter.seq(only.find(_m("den")))
        return f"\\binom{{{top}}}{{{bottom}}}"
    environment = _MATRIX_ENVIRONMENTS.get((begin, end))
    if environment is None:
        return None
    if only.tag == _m("m"):
        columns = _array_columns(only, wrapped=True)
        # `cases` columns are left-aligned by definition, so a
        # specification saying exactly that is part of the environment.
        if columns is None or (environment == "cases" and set(columns) == {"l"}):
            return _matrix_body(converter, only, environment)
        return None
    if only.tag == _m("eqArr") and environment == "cases":
        rows = equation_array_rows(converter, only)
        return r"\begin{cases}" + r" \\ ".join(rows) + r"\end{cases}"
    return None


# -- matrices -------------------------------------------------------------------


def _array_columns(matrix: Any, *, wrapped: bool) -> str | None:
    """``lcr`` column letters when the matrix is really an ``array``."""
    properties = matrix.find(_m("mPr"))
    if properties is None:
        return None
    columns = properties.find(_m("mcs"))
    if columns is None:
        return None
    letters: list[str] = []
    all_single = True
    for column in columns.findall(_m("mc")):
        column_properties = column.find(_m("mcPr"))
        count_el = column_properties.find(_m("count")) if column_properties is not None else None
        justify_el = column_properties.find(_m("mcJc")) if column_properties is not None else None
        try:
            count = int(count_el.get(_VAL, "1")) if count_el is not None else 1
        except ValueError:
            count = 1
        all_single = all_single and count_el is not None and count == 1
        justification = justify_el.get(_VAL, "center") if justify_el is not None else "center"
        letters.extend(_COLUMN_LETTERS.get(justification, "c") * max(count, 1))
    if not letters:
        return None
    width = max((len(row.findall(_m("e"))) for row in matrix.findall(_m("mr"))), default=0)
    letters.extend("c" * max(0, width - len(letters)))
    if any(letter != "c" for letter in letters):
        return "".join(letters)
    # An all-centred specification only says "array" when it has the
    # forward converter's shape: one m:mc per column and nothing else.
    only_columns = all(
        _local(child.tag) == "mcs" for child in properties if isinstance(child.tag, str)
    )
    if not wrapped and all_single and only_columns:
        return "".join(letters)
    return None


def _matrix_body(converter: _Converter, matrix: Any, environment: str, columns: str = "") -> str:
    rows = []
    for row in matrix.findall(_m("mr")):
        rows.append(" & ".join(converter.seq(cell) for cell in row.findall(_m("e"))))
    body = r" \\ ".join(rows)
    argument = f"{{{columns}}}" if columns else ""
    return f"\\begin{{{environment}}}{argument} {body} \\end{{{environment}}}"


def matrix(converter: _Converter, element: Any) -> tuple[str, str]:
    columns = _array_columns(element, wrapped=False)
    if columns is not None:
        return _matrix_body(converter, element, "array", columns), "atom"
    return _matrix_body(converter, element, "matrix"), "atom"


# -- equation arrays --------------------------------------------------------------


def equation_array_rows(converter: _Converter, element: Any) -> list[str]:
    """The rows of an ``m:eqArr``, each ``m:aln`` point written as `` &``."""
    rows = []
    for row in element.findall(_m("e")):
        latex = _ALIGN_MARK.sub(_ALIGN_REPLACEMENT, converter.seq(row)).strip()
        rows.append(latex)
    return rows


def equation_array(converter: _Converter, element: Any) -> tuple[str, str]:
    rows = equation_array_rows(converter, element)
    environment = equation_environment(rows)
    body = r" \\ ".join(rows)
    return f"\\begin{{{environment}}}{body}\\end{{{environment}}}", "atom"
