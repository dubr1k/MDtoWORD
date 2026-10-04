"""``\\begin{...}`` ... ``\\end{...}``: matrices, ``cases``, ``array``
and the line-by-line environments (``aligned``, ``gathered``...), plus the
command handlers for environment syntax met out of place."""

from __future__ import annotations

from typing import NoReturn, Optional

from . import arguments, parser
from .builders import _delimiter, _equation_array, _matrix
from .errors import UnsupportedLatexError
from .style import _PLAIN, _Style
from .tables import (
    _CASES_DELIMITERS,
    _COLUMN_COUNT_ENVIRONMENTS,
    _COLUMN_JUSTIFICATION,
    _LINE_ENVIRONMENTS,
    _MATRIX_DELIMITERS,
    _POSITIONED_ENVIRONMENTS,
)
from .tokenizer import (
    _END_COMMAND,
    _ROW_SEPARATOR,
    Stop,
    _read_raw_bracket,
    _read_raw_group,
    _skip_row_spacing,
    _skip_space,
)


def _read_column_alignments(tokens: list, index: int,
                            environment: str = "array") -> tuple:
    r"""Read `\begin{array}`'s ``{lcr}`` argument. Returns (alignments, index).

    Only ``l``, ``c`` and ``r`` survive: OMML's matrix has no vertical rule,
    no fixed-width paragraph column and no ``@{...}`` insert, so honouring
    ``{c|c}`` is impossible and dropping the rule would turn an augmented
    matrix into an ordinary one.  Both are refused instead.
    """
    missing = UnsupportedLatexError(
        f"\\begin{{{environment}}} needs a column specification, as in "
        f"\\begin{{{environment}}}{{cc}}"
    )
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe][0] != "open":
        raise missing
    specification, index = _read_raw_group(tokens, probe)
    alignments = []
    for character in specification:
        if character.isspace():
            continue
        justification = _COLUMN_JUSTIFICATION.get(character)
        if justification is None:
            raise UnsupportedLatexError(
                f"Column specification is not supported in "
                f"\\begin{{{environment}}}: {character!r} (only 'l', 'c' "
                "and 'r' columns are)"
            )
        alignments.append(justification)
    if not alignments:
        raise missing
    return alignments, index


def _close_environment(tokens: list, index: int, environment: str) -> int:
    r"""Consume the ``\end{environment}`` at `index`. Returns the index after."""
    if index >= len(tokens) or tokens[index] != _END_COMMAND:
        raise UnsupportedLatexError(
            f"\\begin{{{environment}}} without a matching \\end"
        )
    closing, index = _read_raw_group(tokens, index + 1, "end")
    if closing.strip() != environment:
        raise UnsupportedLatexError(
            f"\\begin{{{environment}}} is closed by \\end{{{closing}}}"
        )
    return index


def _parse_line_environment(tokens: list, index: int, environment: str,
                            style: _Style) -> tuple:
    r"""Parse ``aligned``, ``gathered``, ``align*`` and the other
    line-by-line environments into one equation array.

    ``aligned``-like ones turn ``&`` into alignment points, exactly as a
    bare multi-line formula does; ``gathered``-like ones centre their lines
    and refuse ``&``.  A single line needs no array and is returned flat.
    """
    if environment in _COLUMN_COUNT_ENVIRONMENTS:
        count, index = _read_raw_group(tokens, index, f"begin{{{environment}}}")
        if not count.strip().isdigit() or int(count) < 1:
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} needs a column count, as in "
                f"\\begin{{{environment}}}{{2}}, not {count!r}"
            )
    if environment in _POSITIONED_ENVIRONMENTS:
        position, index = _read_raw_bracket(tokens, index)
        if position is not None and position.strip() not in ("t", "b", "c"):
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} takes [t], [b] or [c] as its "
                f"position, not [{position}]"
            )
    mode = _LINE_ENVIRONMENTS[environment]
    alignment = {"align": "explicit", "gather": "forbid"}.get(mode, "auto")
    lines, index = parser._parse_lines(
        tokens, index, stop=lambda kind, value: (kind, value) == _END_COMMAND,
        style=style, alignment=alignment, context=environment,
    )
    index = _close_environment(tokens, index, environment)
    if len(lines) == 1:
        return lines[0], index
    return [_equation_array(lines)], index


def _parse_environment(tokens: list, index: int,
                       style: _Style = _PLAIN) -> tuple:
    """Parse `\\begin{env} ... \\end{env}`. Returns (elements, index)."""
    environment, index = _read_raw_group(tokens, index, "begin")
    environment = environment.strip()
    if environment in ("array", "subarray"):
        alignments, index = _read_column_alignments(tokens, index, environment)
        rows, index = _read_matrix_rows(tokens, index, environment, style)
        used = max((len(row) for row in rows), default=0)
        if used > len(alignments):
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} declares {len(alignments)} columns "
                f"but a row uses {used}"
            )
        return [_matrix(rows, alignments)], index
    if environment in _LINE_ENVIRONMENTS:
        return _parse_line_environment(tokens, index, environment, style)
    starred = environment.endswith("*")
    base = environment[:-1] if starred else environment
    if base in _MATRIX_DELIMITERS:
        justification = None
        if starred:
            # mathtools' `pmatrix*` and friends: one alignment for every
            # column, `[c]` when left out.
            option, index = _read_raw_bracket(tokens, index)
            option = (option or "c").strip()
            justification = _COLUMN_JUSTIFICATION.get(option)
            if justification is None:
                raise UnsupportedLatexError(
                    f"\\begin{{{environment}}} takes [l], [c] or [r], not "
                    f"[{option}]"
                )
        rows, index = _read_matrix_rows(tokens, index, environment, style)
        width = max((len(row) for row in rows), default=0)
        matrix = _matrix(rows, [justification] * width if justification else None)
        begin, end = _MATRIX_DELIMITERS[base]
        if begin or end:
            return [_delimiter(begin, end, [matrix])], index
        return [matrix], index
    if base in _CASES_DELIMITERS:
        # The starred forms set everything after the first `&` as text,
        # with `$...$` for any math in it.
        rows, index = _read_matrix_rows(
            tokens, index, environment, style,
            text_columns_from=1 if starred else None,
        )
        width = max((len(row) for row in rows), default=0)
        begin, end = _CASES_DELIMITERS[base]
        return [_delimiter(begin, end, [_matrix(rows, ["left"] * width)])], index
    raise UnsupportedLatexError(
        f"LaTeX environment is not supported: \\begin{{{environment}}}"
    )


def _read_matrix_rows(tokens: list, index: int, environment: str,
                      style: _Style = _PLAIN,
                      text_columns_from: Optional[int] = None) -> tuple:
    """Read cells split by `&` and rows split by `\\\\`, up to `\\end{env}`.

    From column `text_columns_from` on (when given), cells are text mode.
    """

    def stop(kind: str, value: str) -> bool:
        return kind == "amp" or (kind, value) in (_ROW_SEPARATOR, _END_COMMAND)

    rows: list = []
    row: list = []
    while True:
        if text_columns_from is not None and len(row) >= text_columns_from:
            pieces, index = arguments._read_text_material(tokens, index, environment, stop)
            if pieces and pieces[0][0] == "text":
                pieces[0] = ("text", pieces[0][1].lstrip())
            if pieces and pieces[-1][0] == "text":
                pieces[-1] = ("text", pieces[-1][1].rstrip())
            pieces = [piece for piece in pieces if piece != ("text", "")]
            cell = arguments._text_elements(pieces, arguments._text_run_factory(None, style))
        else:
            cell, index = parser._parse_sequence(tokens, index, stop=stop, style=style)
        row.append(cell)
        if index >= len(tokens):
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} without a matching \\end"
            )
        kind, value = tokens[index]
        if kind == "amp":
            index += 1
            continue
        if (kind, value) == _ROW_SEPARATOR:
            rows.append(row)
            row = []
            index = _skip_row_spacing(tokens, index + 1)
            continue
        if (kind, value) != _END_COMMAND:  # a stray '}' closed us early
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} without a matching \\end"
            )
        index = _close_environment(tokens, index, environment)
        rows.append(row)
        break
    # A final `\\` before `\end` ends the last row rather than starting an
    # empty one.
    if len(rows) > 1 and rows[-1] == [[]]:
        rows.pop()
    return rows, index


def _parse_begin(tokens: list, index: int, name: str, style: _Style,
                 stop: Stop) -> tuple:
    """The ``\\begin`` command handler: see `_parse_environment`."""
    return _parse_environment(tokens, index, style)


def _refuse_end(tokens: list, index: int, name: str, style: _Style,
                stop: Stop) -> NoReturn:
    raise UnsupportedLatexError("\\end without a matching \\begin")


def _refuse_ruled_line(tokens: list, index: int, name: str, style: _Style,
                       stop: Stop) -> NoReturn:
    """``\\hline`` and the other table rules."""
    raise UnsupportedLatexError(
        f"\\{name} is not supported: OMML matrices have no ruled lines"
    )


def _refuse_environment_only(tokens: list, index: int, name: str,
                             style: _Style, stop: Stop) -> NoReturn:
    """Plain-TeX ``\\matrix{...}`` and friends (`_ENVIRONMENT_ONLY`)."""
    raise UnsupportedLatexError(
        f"\\{name} works only as an environment here: write "
        f"\\begin{{{name}}} ... \\end{{{name}}}"
    )
