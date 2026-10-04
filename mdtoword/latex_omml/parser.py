"""The core recursive descent: token list in, OMML elements out.

`_parse_lines` drives a sequence -- line breaks, ``&`` alignment points,
infix commands, declarations and equation numbering act on the sequence
around them, so they are handled here -- and hands every atom to
`_parse_atom`, which dispatches commands through :mod:`.commands`.
`style` is the `_Style` the runs are written in.
"""

from __future__ import annotations

from typing import Optional

from . import arguments, commands
from .builders import (
    _equation_array,
    _infix_element,
    _mark_alignment,
    _run,
    _script,
    _styled_run,
)
from .colors import _colorize, _read_color
from .errors import UnsupportedLatexError
from .style import _FONT_SWITCHES, _PLAIN, _Style
from .symbols import _SPACING
from .tables import _INFIX, _NO_OP_SWITCHES, _NUMBERING
from .tokenizer import (
    _END_COMMAND,
    _RIGHT_COMMAND,
    _ROW_SEPARATOR,
    Stop,
    _read_raw_group,
    _skip_row_spacing,
    _skip_space,
)


def _stop_at_segment_end(stop: Stop) -> Stop:
    r"""`stop`, widened to also terminate on ``\\``, ``&`` and numbering.

    A construct whose operand runs to the end of the enclosing sequence --
    an n-ary operator's body, an infix command's denominator -- must not
    reach past either separator.  A line break plainly ends it, and so does
    an alignment point: in ``\sum_i a_i &= b`` the sum's operand is ``a_i``,
    with ``&`` starting the next aligned segment rather than being swept
    into the sum.  ``\tag`` and its relatives belong to the line, not to
    the operand, so they end it too and are handled by the line itself.
    """

    def combined(kind: str, value: str) -> bool:
        if kind == "amp" or (kind, value) == _ROW_SEPARATOR:
            return True
        if kind == "command" and value[1:] in _NUMBERING:
            return True
        return stop is not None and stop(kind, value)

    return combined


def _apply_alignments(lines: list, marks: list, alignment: str = "auto") -> None:
    r"""Turn each recorded ``&`` position into an OMML alignment point.

    With `alignment` "auto" -- a formula that is not inside an alignment
    environment -- a single line means there is no second line to align
    against, and a lone ``&`` in an ordinary formula is far more likely a
    literal ampersand that wanted escaping -- ``Tom & Jerry`` inside ``$…$``
    -- so that is refused rather than turned into an invisible marker.
    Inside ``aligned`` and its relatives ("explicit") the ``&`` is
    unambiguous, and on a single line it simply has nothing to align.
    """
    if len(lines) == 1 and marks[0]:
        if alignment == "explicit":
            return
        raise UnsupportedLatexError(
            "'&' is only meaningful inside a matrix environment or between "
            "the lines of a multi-line formula; write a literal '&' as '\\&'"
        )
    for line, positions in zip(lines, marks, strict=True):
        # Descending, so an inserted marker run cannot shift a position
        # that has not been applied yet.
        for position in sorted(positions, reverse=True):
            _mark_alignment(line, position)


def _parse_infix(tokens: list, index: int, numerator: list, stop: Stop,
                 style: _Style = _PLAIN, color: Optional[str] = None) -> tuple:
    r"""Consume an infix command and the denominator that follows it.

    `numerator` is whatever the enclosing sequence has parsed so far.  The
    denominator runs to the end of that sequence, so it takes the same
    `stop` -- widened to stop at a line break as well, since ``\\`` ends the
    fraction rather than being swallowed into its bottom half.  A colour
    or font declared before the command carries on into the denominator,
    as it does in TeX.
    """
    name = tokens[index][1][1:]
    denominator, index = _parse_lines(
        tokens, index + 1, _stop_at_segment_end(stop), style,
        allow_infix=False, color=color,
    )
    return [_infix_element(name, numerator, denominator[0])], index


def _read_numbering(tokens: list, index: int) -> tuple:
    r"""Consume ``\tag``, ``\label``, ``\nonumber`` or ``\notag``.

    Returns (tag_elements_or_None, index).  Only ``\tag`` produces
    anything: its text, in parentheses unless it is ``\tag*``, as an
    upright run set a ``\quad`` apart from the formula.
    """
    name = tokens[index][1][1:]
    index += 1
    if name in ("nonumber", "notag"):
        return None, index
    if name == "label":
        _, index = _read_raw_group(tokens, index, "label", verbatim=True)
        return None, index
    probe = _skip_space(tokens, index)
    starred = probe < len(tokens) and tokens[probe] == ("other", "*")
    if starred:
        index = probe + 1
    pieces, index = arguments._read_text_group(tokens, index, "tag*" if starred else "tag")
    opening, closing = ("", "") if starred else ("(", ")")
    tag: list = [_run(_SPACING["quad"])]
    if all(kind == "text" for kind, _ in pieces):
        text = "".join(value for _, value in pieces)
        tag.append(_run(f"{opening}{text}{closing}", upright=True))
        return tag, index
    if opening:
        tag.append(_run(opening, upright=True))
    tag.extend(arguments._text_elements(pieces, lambda text: _run(text, upright=True)))
    if closing:
        tag.append(_run(closing, upright=True))
    return tag, index


def _parse_lines(tokens: list, index: int, stop: Stop = None,
                 style: _Style = _PLAIN, allow_infix: bool = True,
                 alignment: str = "auto", color: Optional[str] = None,
                 context: str = "") -> tuple:
    r"""Parse tokens until `stop`, a closing brace, or the end, splitting on ``\\``.

    Returns (lines, index) with `index` left ON the terminating token and at
    least one line -- an empty sequence gives ``[[]]``.  A trailing ``\\``
    ends the last line rather than starting an empty one, matching how
    `_read_matrix_rows` treats the same token.

    An ``&`` marks an alignment point on the segment that follows it, so
    ``a &= b \\ c &= d`` lines its two ``=`` up.  A matrix consumes ``&`` as
    a cell break through `stop` long before it reaches here, so only the
    equation-array sense of the token is left by this point.  Without a
    ``\\`` there is nothing to align against, and a lone ``&`` is far more
    likely a literal ampersand that wanted escaping -- so that still fails
    loudly rather than becoming an invisible marker, unless `alignment` is
    "explicit" (inside ``aligned`` and friends).  "forbid" refuses every
    ``&``: ``gathered`` (named by `context`) centres its lines instead.

    Declarations act on the rest of the sequence: ``\color{red}``, ``\bf``
    and the like change `style`/`color` for what follows, up to the end of
    the group -- or of the line or ``&`` cell, which amsmath makes groups
    of their own.  ``\tag{...}`` is collected and placed at the end of its
    line; ``\label``, ``\nonumber`` and ``\notag`` produce nothing.

    `allow_infix` is cleared for the right-hand side of an infix command, so
    ``a \over b \over c`` is refused as ambiguous the way TeX itself refuses
    it, instead of silently picking one of the two readings.
    """
    lines: list = []
    marks: list = []
    tags: list = []
    elements: list = []
    positions: list = []
    tag = None
    current_style, current_color = style, color
    while index < len(tokens):
        kind, value = tokens[index]
        if kind == "space":
            index += 1
            continue
        if kind == "close":
            break
        if stop is not None and stop(kind, value):
            break
        if kind == "amp":
            if alignment == "forbid":
                raise UnsupportedLatexError(
                    f"'&' has no meaning inside \\begin{{{context}}}, which "
                    "centres its lines; use \\begin{aligned} to align them, "
                    "or write a literal '&' as '\\&'"
                )
            # Where the *next* element will land, so `a &= b` marks the `=`.
            positions.append(len(elements))
            current_style, current_color = style, color
            index += 1
            continue
        if (kind, value) == _ROW_SEPARATOR:
            lines.append(elements)
            marks.append(positions)
            tags.append(tag)
            elements = []
            positions = []
            tag = None
            current_style, current_color = style, color
            index = _skip_row_spacing(tokens, index + 1)
            continue
        if kind == "command":
            name = value[1:]
            if name in _INFIX:
                if not allow_infix:
                    raise UnsupportedLatexError(
                        f"Two infix commands in one group: \\{name} follows "
                        "another one; brace the halves, as in "
                        "{{a \\over b} \\over c}"
                    )
                elements, index = _parse_infix(
                    tokens, index, elements, stop, current_style, current_color)
                continue
            if name in _NUMBERING:
                line_tag, index = _read_numbering(tokens, index)
                if line_tag is not None:
                    if tag is not None:
                        raise UnsupportedLatexError(
                            "A line has two \\tag commands; give each line "
                            "one at most"
                        )
                    tag = line_tag
                continue
            if name in _NO_OP_SWITCHES:
                index += 1
                continue
            if name in _FONT_SWITCHES:
                current_style = _FONT_SWITCHES[name]
                index += 1
                continue
            if name == "color":
                current_color, index = _read_color(tokens, index + 1, "color")
                continue
        atom, index = _parse_atom(tokens, index, current_style, stop)
        sub, sup, index = arguments._read_scripts(tokens, index, current_style)
        if sub is not None or sup is not None:
            atom = [_script(atom, sub, sup)]
        if current_color is not None:
            _colorize(atom, current_color)
        elements.extend(atom)
    lines.append(elements)
    marks.append(positions)
    tags.append(tag)
    if len(lines) > 1 and not lines[-1] and tags[-1] is None:
        lines.pop()
        marks.pop()
        tags.pop()
    _apply_alignments(lines, marks, alignment)
    # After the alignment points, so a marker for an `&` at the very end of
    # a line lands before the tag rather than on it.
    for line, line_tag in zip(lines, tags, strict=True):
        if line_tag is not None:
            line.extend(line_tag)
    return lines, index


def _parse_sequence(tokens: list, index: int, stop: Stop = None,
                    style: _Style = _PLAIN) -> tuple:
    r"""Parse tokens into elements until `stop`, a closing brace, or the end.

    Returns (elements, index) with `index` left ON the terminating token.
    A sequence broken by ``\\`` becomes a single equation array holding one
    line each, so the caller still gets one flat element list.
    """
    lines, index = _parse_lines(tokens, index, stop, style)
    if len(lines) == 1:
        return lines[0], index
    return [_equation_array(lines)], index


def _parse_group(tokens: list, index: int, style: _Style = _PLAIN,
                 owner: str = "") -> tuple:
    """Read a `{...}` group, or exactly one atom if there is no brace.

    This is what the arguments of \\frac, \\sqrt, `^` and `_` all need.
    `owner` names the command whose argument this is, for error messages.

    TeX's own grouping rule treats an unbraced argument as exactly one
    token: `\\frac12x` means `\\frac{1}{2}x`, and `x^12` superscripts only
    the `1`.  Our tokenizer merges consecutive digits into a single
    `number` token (needed so `\\frac{12}{x}` and plain `12 + 3` read the
    whole literal), so here -- the unbraced path only -- a multi-digit
    number is split: its first character is consumed as the atom and the
    rest is pushed back onto the token stream as a new pending token.
    """
    index = _skip_space(tokens, index)
    if index >= len(tokens):
        if owner:
            raise UnsupportedLatexError(
                f"\\{owner} is missing an argument at the end of the formula"
            )
        raise UnsupportedLatexError("Formula ends where an argument was expected")
    if tokens[index][0] == "open":
        elements, index = _parse_sequence(tokens, index + 1, style=style)
        if index >= len(tokens) or tokens[index][0] != "close":
            raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
        return elements, index + 1
    kind, value = tokens[index]
    if owner and (kind == "close" or kind == "amp"
                  or (kind, value) in (_ROW_SEPARATOR, _END_COMMAND,
                                       _RIGHT_COMMAND)):
        raise UnsupportedLatexError(
            f"\\{owner} is missing an argument before {value!r}"
        )
    if kind == "number" and len(value) > 1:
        tokens[index] = ("number", value[0])
        tokens.insert(index + 1, ("number", value[1:]))
    return _parse_atom(tokens, index, style)


def _parse_atom(tokens: list, index: int, style: _Style = _PLAIN,
                stop: Stop = None) -> tuple:
    """Parse one atom: a group, a character, or a command. Returns (elements, index).

    `stop` is the terminator of the sequence this atom belongs to.  Only the
    n-ary operators need it: their operand runs to the end of the enclosing
    construct, so they must know where that end is.
    """
    kind, value = tokens[index]
    if kind == "open":
        return _parse_group(tokens, index, style)
    if kind == "command":
        return commands._parse_command(tokens, index, style, stop)
    if kind == "letter":
        return [_styled_run(value, style, "letter")], index + 1
    if kind in ("number", "other", "bracket"):
        if (kind, value) == ("other", "'"):
            # A prime: `f'` is TeX's `f^{\prime}`, drawn as ′, not as an
            # apostrophe.
            value = "′"
        elif (kind, value) == ("other", "~"):
            # A tie: an unbreakable space, never a visible tilde.
            return [_run(_SPACING[" "])], index + 1
        return [_styled_run(value, style, "number")], index + 1
    if kind in ("sub", "sup"):
        raise UnsupportedLatexError(
            f"'{value}' has nothing to attach to in the formula"
        )
    if kind == "amp":
        raise UnsupportedLatexError(
            "'&' is only meaningful inside a matrix environment or between "
            "the lines of a multi-line formula; write a literal '&' as '\\&'"
        )
    raise UnsupportedLatexError(f"Could not read {value!r} in the formula")
