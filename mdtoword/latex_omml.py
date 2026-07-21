"""Convert a LaTeX math string into Office Math Markup Language (OMML).

The result is a real Word equation that stays editable in Word's equation
editor, not a text fallback with dollar signs around it.

Only constructs listed in ``_parse_command`` -- plus the infix commands and
the ``\\`` line break, which ``_parse_lines`` has to handle itself because
they split the sequence around them -- are understood.  Anything else raises
:class:`UnsupportedLatexError` naming the offending construct, so the caller
can report it instead of silently emitting wrong output.
"""

from __future__ import annotations

import re
from typing import Any, Callable, Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class UnsupportedLatexError(ValueError):
    """Raised when a LaTeX construct has no OMML equivalent here."""


_SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ε", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "⋅", "ast": "∗", "star": "⋆",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥",
    "neq": "≠", "ne": "≠", "approx": "≈", "equiv": "≡",
    "sim": "∼", "simeq": "≃", "cong": "≅", "propto": "∝",
    "ll": "≪", "gg": "≫",
    "in": "∈", "notin": "∉", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "supseteq": "⊇", "cup": "∪", "cap": "∩",
    "emptyset": "∅", "varnothing": "∅", "setminus": "∖",
    "forall": "∀", "exists": "∃", "neg": "¬",
    "land": "∧", "lor": "∨",
    "rightarrow": "→", "to": "→", "leftarrow": "←",
    "leftrightarrow": "↔", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "Leftrightarrow": "⇔", "mapsto": "↦",
    "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮",
    "ddots": "⋱", "prime": "′", "degree": "°",
    "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ",
    "aleph": "ℵ", "angle": "∠", "perp": "⊥", "parallel": "∥",
    "circ": "∘", "bullet": "∙", "oplus": "⊕", "otimes": "⊗",
}

# Spacing commands become real spaces; Word does its own math spacing anyway.
_SPACING = {",": " ", ";": " ", ":": " ", "!": "", " ": " ",
            "quad": " ", "qquad": "  "}

# `\%` and friends: the backslash only escapes LaTeX's own syntax.
_ESCAPED = {"{": "{", "}": "}", "%": "%", "$": "$",
            "&": "&", "#": "#", "_": "_"}

_UPRIGHT_FUNCTIONS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "det", "dim", "ker",
    "deg", "gcd", "hom", "arg", "max", "min", "sup", "inf",
}

_FRACTIONS = {"frac", "dfrac", "tfrac"}
_UPRIGHT_TEXT = {"text", "textrm", "textnormal", "mathrm", "operatorname"}
_BOLD_UPRIGHT_STYLE = {"mathbf"}
_BOLD_ITALIC_STYLE = {"boldsymbol", "bm"}
_ITALIC_STYLE = {"mathit"}

# Big operators.  Each takes optional `_`/`^` limits and then swallows the
# rest of the enclosing group as its operand.
_NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐",
    "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
    "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁",
    "bigotimes": "⨂", "bigvee": "⋁", "bigwedge": "⋀",
}

# Integrals keep their limits beside the sign; every other big operator
# stacks them above and below.
_INTEGRAL_CHARACTERS = {"∫", "∬", "∭", "∮"}

# Operators whose subscript sits underneath rather than beside them.
_LIMIT_OPERATORS = {"lim": "lim", "limsup": "lim sup", "liminf": "lim inf"}

# Combining marks: each one composes with the character it follows.
_ACCENTS = {
    "hat": "̂", "widehat": "̂", "tilde": "̃",
    "widetilde": "̃", "bar": "̄", "vec": "⃗",
    "dot": "̇", "ddot": "̈", "acute": "́", "grave": "̀",
    "check": "̌", "breve": "̆",
}

_MATRIX_DELIMITERS = {
    "matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
    "cases": ("{", ""),
}

# What may follow `\left` and `\right`.  `.` is LaTeX's "no delimiter here",
# which OMML spells as an empty begChr/endChr.
_DELIMITER_CHARACTERS = {"(": "(", ")": ")", "[": "[", "]": "]",
                         "|": "|", "/": "/", ".": ""}
_DELIMITER_COMMANDS = {
    "{": "{", "}": "}", "|": "‖", "backslash": "\\",
    "lbrace": "{", "rbrace": "}", "langle": "⟨", "rangle": "⟩",
    "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈", "rceil": "⌉",
    "vert": "|", "Vert": "‖", "lvert": "|", "rvert": "|",
    "lVert": "‖", "rVert": "‖",
}

_ROW_SEPARATOR = ("command", "\\\\")
_END_COMMAND = ("command", "\\end")
_RIGHT_COMMAND = ("command", "\\right")

# Infix commands: each splits the group it appears in, taking everything to
# its left as the numerator and everything to its right as the denominator.
# That is why they cannot live in `_parse_command` with the prefix commands
# -- by the time it runs, the numerator has already been parsed and emitted.
_INFIX = frozenset({"over", "atop", "choose"})

# `\begin{array}` column letters.  OMML's `m:mcJc` spells the same three
# alignments; anything else in a column specification (`|` rules, `p{...}`
# paragraph columns, `@{...}` inserts) has no OMML equivalent and is
# refused rather than dropped.
_COLUMN_JUSTIFICATION = {"l": "left", "c": "center", "r": "right"}

# Plain-TeX spellings of environments this module supports only in their
# LaTeX `\begin{...}` / `\end{...}` form.  `\matrix{a & b}` takes its rows
# as a braced argument instead, which is a different construct with
# different bracing rules, so these fail loudly -- naming the environment
# form to use -- rather than being quietly treated as the environment.
_ENVIRONMENT_ONLY = frozenset({
    "matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix",
    "array", "cases",
})

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


# --------------------------------------------------------------------------
# OMML element constructors.  Build every element through these so the shapes
# stay consistent across this module and its callers.
# --------------------------------------------------------------------------


def _el(tag: str) -> Any:
    return OxmlElement(f"m:{tag}")


def _run(text: str, *, italic: bool = False, upright: bool = False,
         bold: bool = False) -> Any:
    """One <m:r>. Identifiers are italic, numbers and operators are not."""
    run = _el("r")
    if italic or upright or bold:
        properties = _el("rPr")
        # OOXML's CT_RPR makes <m:nor> and <m:sty> a *choice*: a run may
        # carry one or the other, never both. `\mathbf` is upright and bold
        # at once, so the two have to be reconciled rather than stacked --
        # emitting both is rejected by the ISO/IEC 29500-4 schema. <m:sty>
        # already encodes uprightness ("b" is bold upright, "bi" is bold
        # italic), so whenever a style value applies it says everything
        # <m:nor> would have, and <m:nor> is left for the unbolded literal
        # text of \text{...} and friends.
        if bold:
            sty = _el("sty")
            sty.set(qn("m:val"), "bi" if italic and not upright else "b")
            properties.append(sty)
        elif upright:
            nor = _el("nor")
            nor.set(qn("m:val"), "1")
            properties.append(nor)
        elif italic:
            sty = _el("sty")
            sty.set(qn("m:val"), "i")
            properties.append(sty)
        run.append(properties)
    text_element = _el("t")
    # Only declare xml:space when it changes anything: an unconditional
    # attribute would bloat every single run in the document.
    if text != text.strip():
        text_element.set(qn("xml:space"), "preserve")
    text_element.text = text
    run.append(text_element)
    return run


def _wrap(tag: str, children: list[Any]) -> Any:
    """A container element such as <m:num>, <m:den>, <m:e>, <m:sup>, <m:sub>."""
    element = _el(tag)
    for child in children:
        element.append(child)
    return element


def _fraction(numerator: list[Any], denominator: list[Any], *, no_bar: bool = False) -> Any:
    fraction = _el("f")
    if no_bar:
        properties = _el("fPr")
        fraction_type = _el("type")
        fraction_type.set(qn("m:val"), "noBar")
        properties.append(fraction_type)
        fraction.append(properties)
    fraction.append(_wrap("num", numerator))
    fraction.append(_wrap("den", denominator))
    return fraction


def _radical(degree: list[Any] | None, radicand: list[Any]) -> Any:
    radical = _el("rad")
    properties = _el("radPr")
    if degree is None:
        hide = _el("degHide")
        hide.set(qn("m:val"), "1")
        properties.append(hide)
    radical.append(properties)
    radical.append(_wrap("deg", degree or []))
    radical.append(_wrap("e", radicand))
    return radical


def _script(base: list[Any], sub: list[Any] | None, sup: list[Any] | None) -> Any:
    """<m:sSub>, <m:sSup> or <m:sSubSup> depending on which scripts are present."""
    if sub is not None and sup is not None:
        element = _el("sSubSup")
        element.append(_wrap("e", base))
        element.append(_wrap("sub", sub))
        element.append(_wrap("sup", sup))
        return element
    if sub is not None:
        element = _el("sSub")
        element.append(_wrap("e", base))
        element.append(_wrap("sub", sub))
        return element
    element = _el("sSup")
    element.append(_wrap("e", base))
    element.append(_wrap("sup", sup or []))
    return element


def _property(tag: str, name: str, value: str) -> Any:
    """A properties element holding one `m:val` child, e.g. <m:accPr><m:chr/>."""
    properties = _el(tag)
    child = _el(name)
    child.set(qn("m:val"), value)
    properties.append(child)
    return properties


def _nary(character: str, sub: list[Any] | None, sup: list[Any] | None,
          body: list[Any]) -> Any:
    """<m:nary>: a big operator with its limits and its operand."""
    nary = _el("nary")
    properties = _property("naryPr", "chr", character)
    limit_location = _el("limLoc")
    limit_location.set(
        qn("m:val"),
        "subSup" if character in _INTEGRAL_CHARACTERS else "undOvr",
    )
    properties.append(limit_location)
    # Without these Word draws an empty placeholder box where the missing
    # limit would go.
    for missing, tag in ((sub is None, "subHide"), (sup is None, "supHide")):
        if missing:
            hide = _el(tag)
            hide.set(qn("m:val"), "1")
            properties.append(hide)
    nary.append(properties)
    nary.append(_wrap("sub", sub or []))
    nary.append(_wrap("sup", sup or []))
    nary.append(_wrap("e", body))
    return nary


def _delimiter(begin: str, end: str, children: list[Any]) -> Any:
    """<m:d>: a fenced group. An empty `begin`/`end` means "no fence"."""
    delimiter = _el("d")
    properties = _el("dPr")
    for tag, value in (("begChr", begin), ("endChr", end)):
        child = _el(tag)
        child.set(qn("m:val"), value)
        properties.append(child)
    delimiter.append(properties)
    delimiter.append(_wrap("e", children))
    return delimiter


def _accent(character: str, base: list[Any]) -> Any:
    accent = _el("acc")
    accent.append(_property("accPr", "chr", character))
    accent.append(_wrap("e", base))
    return accent


def _overline(base: list[Any], *, position: str = "top") -> Any:
    """<m:bar>: a rule above (`top`) or below (`bot`) the base."""
    bar = _el("bar")
    bar.append(_property("barPr", "pos", position))
    bar.append(_wrap("e", base))
    return bar


def _limit_low(base: list[Any], limit: list[Any]) -> Any:
    element = _el("limLow")
    element.append(_wrap("e", base))
    element.append(_wrap("lim", limit))
    return element


def _matrix(rows: list[list[list[Any]]],
            alignments: list[str] | None = None) -> Any:
    """<m:m>: rows of cells. Short rows are padded so Word sees a rectangle.

    `alignments` is one OMML `m:mcJc` value per column, as `\\begin{array}`
    spells it; without it Word centres every column, which is what the
    matrix environments want.  A specification wider than the widest row
    still shapes the matrix, so `\\begin{array}{ccc} a & b \\\\ c & d` keeps
    the declared -- empty -- third column.
    """
    width = max((len(row) for row in rows), default=0)
    matrix = _el("m")
    if alignments:
        width = max(width, len(alignments))
        properties = _el("mPr")
        columns = _el("mcs")
        for justification in alignments:
            column_properties = _el("mcPr")
            count = _el("count")
            count.set(qn("m:val"), "1")
            column_properties.append(count)
            column_justification = _el("mcJc")
            column_justification.set(qn("m:val"), justification)
            column_properties.append(column_justification)
            column = _el("mc")
            column.append(column_properties)
            columns.append(column)
        properties.append(columns)
        # <m:mPr> is required to come first; Word rejects the part otherwise.
        matrix.append(properties)
    for row in rows:
        row_element = _el("mr")
        for column_index in range(width):
            row_element.append(
                _wrap("e", row[column_index] if column_index < len(row) else [])
            )
        matrix.append(row_element)
    return matrix


def _equation_array(lines: list[list[Any]]) -> Any:
    r"""<m:eqArr>: stacked lines, Word's rendering of ``\\`` outside a matrix."""
    array = _el("eqArr")
    for line in lines:
        array.append(_wrap("e", line))
    return array


def _mark_alignment(elements: list[Any], position: int) -> None:
    r"""Make the element at `position` an alignment point, as ``&`` asks.

    OMML spells an alignment point as ``<m:aln/>`` inside a run's
    ``<m:rPr>`` -- the run that *starts* the aligned segment carries it, so
    ``a &= b`` marks the ``=``.  ``<m:aln>`` is last in ``CT_RPR``'s
    sequence, and `_run` only ever writes ``<m:nor>`` or ``<m:sty>`` before
    it, so appending is always in schema order.

    Only a run can carry the marker.  When the segment starts with anything
    else -- a fraction, a matrix -- an empty run is inserted to hold it,
    which adds no glyph of its own.
    """
    if position < len(elements) and elements[position].tag == qn("m:r"):
        run = elements[position]
        properties = run.find(qn("m:rPr"))
        if properties is None:
            properties = _el("rPr")
            run.insert(0, properties)
        properties.append(_el("aln"))
        return
    marker = _el("r")
    properties = _el("rPr")
    properties.append(_el("aln"))
    marker.append(properties)
    marker_text = _el("t")
    # Explicitly empty rather than left unset, so the run serialises as
    # <m:t></m:t> -- the shape Word writes -- instead of a bare <m:t/>.
    marker_text.text = ""
    marker.append(marker_text)
    elements.insert(position, marker)


def _infix_element(name: str, numerator: list[Any],
                   denominator: list[Any]) -> Any:
    r"""Build the element for ``\over``, ``\atop`` or ``\choose``.

    ``\choose`` is exactly ``\binom`` written infix, so both go through the
    same barless-fraction-in-parentheses shape.
    """
    if name == "over":
        return _fraction(numerator, denominator)
    stack = _fraction(numerator, denominator, no_bar=True)
    if name == "atop":
        return stack
    return _delimiter("(", ")", [stack])


# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------


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


# --------------------------------------------------------------------------
# Parser.  Recursive descent over the token list, producing OMML elements.
# `style` is None, "bold" (upright, \mathbf), "bolditalic" (\boldsymbol / \bm)
# or "italic" (\mathit).
# --------------------------------------------------------------------------


def _stop_at_segment_end(stop: Stop) -> Stop:
    r"""`stop`, widened to also terminate on ``\\`` and ``&``.

    A construct whose operand runs to the end of the enclosing sequence --
    an n-ary operator's body, an infix command's denominator -- must not
    reach past either separator.  A line break plainly ends it, and so does
    an alignment point: in ``\sum_i a_i &= b`` the sum's operand is ``a_i``,
    with ``&`` starting the next aligned segment rather than being swept
    into the sum.
    """

    def combined(kind: str, value: str) -> bool:
        if kind == "amp" or (kind, value) == _ROW_SEPARATOR:
            return True
        return stop is not None and stop(kind, value)

    return combined


def _apply_alignments(lines: list, marks: list) -> None:
    r"""Turn each recorded ``&`` position into an OMML alignment point.

    A single line means there is no second line to align against, and a
    lone ``&`` in an ordinary formula is far more likely a literal
    ampersand that wanted escaping -- ``Tom & Jerry`` inside ``$…$`` -- so
    that is refused rather than turned into an invisible marker.
    """
    if len(lines) == 1 and marks[0]:
        raise UnsupportedLatexError(
            "'&' is only meaningful inside a matrix environment or between "
            "the lines of a multi-line formula; write a literal '&' as '\\&'"
        )
    for line, positions in zip(lines, marks):
        # Descending, so an inserted marker run cannot shift a position
        # that has not been applied yet.
        for position in sorted(positions, reverse=True):
            _mark_alignment(line, position)


def _parse_infix(tokens: list, index: int, numerator: list, stop: Stop,
                 style: Optional[str] = None) -> tuple:
    r"""Consume an infix command and the denominator that follows it.

    `numerator` is whatever the enclosing sequence has parsed so far.  The
    denominator runs to the end of that sequence, so it takes the same
    `stop` -- widened to stop at a line break as well, since ``\\`` ends the
    fraction rather than being swallowed into its bottom half.
    """
    name = tokens[index][1][1:]
    denominator, index = _parse_lines(
        tokens, index + 1, _stop_at_segment_end(stop), style, allow_infix=False,
    )
    return [_infix_element(name, numerator, denominator[0])], index


def _parse_lines(tokens: list, index: int, stop: Stop = None,
                 style: Optional[str] = None, allow_infix: bool = True) -> tuple:
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
    loudly rather than becoming an invisible marker.

    `allow_infix` is cleared for the right-hand side of an infix command, so
    ``a \over b \over c`` is refused as ambiguous the way TeX itself refuses
    it, instead of silently picking one of the two readings.
    """
    lines: list = []
    marks: list = []
    elements: list = []
    positions: list = []
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
            # Where the *next* element will land, so `a &= b` marks the `=`.
            positions.append(len(elements))
            index += 1
            continue
        if (kind, value) == _ROW_SEPARATOR:
            lines.append(elements)
            marks.append(positions)
            elements = []
            positions = []
            index += 1
            continue
        if kind == "command" and value[1:] in _INFIX:
            if not allow_infix:
                raise UnsupportedLatexError(
                    f"Two infix commands in one group: \\{value[1:]} follows "
                    "another one; brace the halves, as in "
                    "{{a \\over b} \\over c}"
                )
            elements, index = _parse_infix(tokens, index, elements, stop, style)
            continue
        atom, index = _parse_atom(tokens, index, style, stop)
        sub, sup, index = _read_scripts(tokens, index, style)
        if sub is None and sup is None:
            elements.extend(atom)
        else:
            elements.append(_script(atom, sub, sup))
    lines.append(elements)
    marks.append(positions)
    if len(lines) > 1 and not lines[-1]:
        lines.pop()
        marks.pop()
    _apply_alignments(lines, marks)
    return lines, index


def _parse_sequence(tokens: list, index: int, stop: Stop = None,
                    style: Optional[str] = None) -> tuple:
    r"""Parse tokens into elements until `stop`, a closing brace, or the end.

    Returns (elements, index) with `index` left ON the terminating token.
    A sequence broken by ``\\`` becomes a single equation array holding one
    line each, so the caller still gets one flat element list.
    """
    lines, index = _parse_lines(tokens, index, stop, style)
    if len(lines) == 1:
        return lines[0], index
    return [_equation_array(lines)], index


def _parse_group(tokens: list, index: int, style: Optional[str] = None) -> tuple:
    """Read a `{...}` group, or exactly one atom if there is no brace.

    This is what the arguments of \\frac, \\sqrt, `^` and `_` all need.

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
        raise UnsupportedLatexError("Formula ends where an argument was expected")
    if tokens[index][0] == "open":
        elements, index = _parse_sequence(tokens, index + 1, style=style)
        if index >= len(tokens) or tokens[index][0] != "close":
            raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
        return elements, index + 1
    kind, value = tokens[index]
    if kind == "number" and len(value) > 1:
        tokens[index] = ("number", value[0])
        tokens.insert(index + 1, ("number", value[1:]))
    return _parse_atom(tokens, index, style)


def _read_raw_group(tokens: list, index: int) -> tuple:
    """Read a `{...}` group as plain characters instead of as math.

    The tokenizer splits words letter by letter, so `\\begin{pmatrix}` and
    `\\text{если}` both need the raw values glued back together rather than
    parsed.  Returns (text, index_after_group).
    """
    index = _skip_space(tokens, index)
    if index >= len(tokens) or tokens[index][0] != "open":
        raise UnsupportedLatexError("Expected '{' after the command")
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
        if kind == "command":
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


def _read_bracket_argument(tokens: list, index: int,
                           style: Optional[str] = None) -> tuple:
    """Read an optional `[...]` argument, as used by `\\sqrt[3]{8}`.

    Returns (elements, index) or (None, index) when no bracket follows.
    """
    probe = _skip_space(tokens, index)
    if probe >= len(tokens) or tokens[probe] != ("bracket", "["):
        return None, index
    elements, probe = _parse_sequence(
        tokens, probe + 1, stop=lambda k, v: k == "bracket" and v == "]", style=style
    )
    if probe >= len(tokens) or tokens[probe] != ("bracket", "]"):
        raise UnsupportedLatexError("Unclosed '[' in an optional argument")
    return elements, probe + 1


def _read_scripts(tokens: list, index: int, style: Optional[str] = None) -> tuple:
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
            sub, index = _parse_group(tokens, probe + 1, style)
        elif kind == "sup":
            if sup is not None:
                raise UnsupportedLatexError("Two superscripts on one base")
            sup, index = _parse_group(tokens, probe + 1, style)
        else:
            return sub, sup, index


def _parse_atom(tokens: list, index: int, style: Optional[str] = None,
                stop: Stop = None) -> tuple:
    """Parse one atom: a group, a character, or a command. Returns (elements, index).

    `stop` is the terminator of the sequence this atom belongs to.  Only the
    n-ary operators need it: their operand runs to the end of the enclosing
    construct, so they must know where that end is.
    """
    kind, value = tokens[index]
    bold = style in ("bold", "bolditalic")
    if kind == "open":
        return _parse_group(tokens, index, style)
    if kind == "command":
        return _parse_command(tokens, index, style, stop)
    if kind == "letter":
        if style == "bold":
            return [_run(value, upright=True, bold=True)], index + 1
        return [_run(value, italic=True, bold=bold)], index + 1
    if kind in ("number", "other", "bracket"):
        if style == "italic":
            return [_run(value, italic=True)], index + 1
        return [_run(value, bold=bold)], index + 1
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


def _read_delimiter(tokens: list, index: int, command: str) -> tuple:
    """Read the fence character that follows `\\left` or `\\right`.

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


def _read_column_alignments(tokens: list, index: int) -> tuple:
    r"""Read `\begin{array}`'s ``{lcr}`` argument. Returns (alignments, index).

    Only ``l``, ``c`` and ``r`` survive: OMML's matrix has no vertical rule,
    no fixed-width paragraph column and no ``@{...}`` insert, so honouring
    ``{c|c}`` is impossible and dropping the rule would turn an augmented
    matrix into an ordinary one.  Both are refused instead.
    """
    missing = UnsupportedLatexError(
        "\\begin{array} needs a column specification, as in \\begin{array}{cc}"
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
                f"Column specification is not supported in \\begin{{array}}: "
                f"{character!r} (only 'l', 'c' and 'r' columns are)"
            )
        alignments.append(justification)
    if not alignments:
        raise missing
    return alignments, index


def _parse_environment(tokens: list, index: int,
                       style: Optional[str] = None) -> tuple:
    """Parse `\\begin{env} ... \\end{env}`. Returns (elements, index)."""
    environment, index = _read_raw_group(tokens, index)
    if environment == "array":
        alignments, index = _read_column_alignments(tokens, index)
        rows, index = _read_matrix_rows(tokens, index, environment, style)
        used = max((len(row) for row in rows), default=0)
        if used > len(alignments):
            raise UnsupportedLatexError(
                f"\\begin{{array}} declares {len(alignments)} columns but a "
                f"row uses {used}"
            )
        return [_matrix(rows, alignments)], index
    if environment not in _MATRIX_DELIMITERS:
        raise UnsupportedLatexError(
            f"LaTeX environment is not supported: \\begin{{{environment}}}"
        )
    rows, index = _read_matrix_rows(tokens, index, environment, style)
    begin, end = _MATRIX_DELIMITERS[environment]
    matrix = _matrix(rows)
    if begin or end:
        return [_delimiter(begin, end, [matrix])], index
    return [matrix], index


def _read_matrix_rows(tokens: list, index: int, environment: str,
                      style: Optional[str] = None) -> tuple:
    """Read cells split by `&` and rows split by `\\\\`, up to `\\end{env}`."""

    def stop(kind: str, value: str) -> bool:
        return kind == "amp" or (kind, value) in (_ROW_SEPARATOR, _END_COMMAND)

    rows: list = []
    row: list = []
    while True:
        cell, index = _parse_sequence(tokens, index, stop=stop, style=style)
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
            index += 1
            continue
        if (kind, value) != _END_COMMAND:  # a stray '}' closed us early
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} without a matching \\end"
            )
        closing, index = _read_raw_group(tokens, index + 1)
        if closing != environment:
            raise UnsupportedLatexError(
                f"\\begin{{{environment}}} is closed by \\end{{{closing}}}"
            )
        rows.append(row)
        break
    # A final `\\` before `\end` ends the last row rather than starting an
    # empty one.
    if len(rows) > 1 and rows[-1] == [[]]:
        rows.pop()
    return rows, index


def _parse_command(tokens: list, index: int, style: Optional[str] = None,
                   stop: Stop = None) -> tuple:
    """Parse one `\\command` and its arguments. Returns (elements, index)."""
    name = tokens[index][1][1:]
    index += 1
    bold = style in ("bold", "bolditalic")

    if name in _FRACTIONS:
        numerator, index = _parse_group(tokens, index, style)
        denominator, index = _parse_group(tokens, index, style)
        return [_fraction(numerator, denominator)], index

    if name == "sqrt":
        degree, index = _read_bracket_argument(tokens, index, style)
        if not degree:
            degree = None
        radicand, index = _parse_group(tokens, index, style)
        return [_radical(degree, radicand)], index

    if name in _UPRIGHT_TEXT:
        text, index = _read_raw_group(tokens, index)
        return [_run(text, upright=True, bold=bold)], index

    if name in _BOLD_UPRIGHT_STYLE:
        elements, index = _parse_group(tokens, index, "bold")
        return elements, index

    if name in _BOLD_ITALIC_STYLE:
        elements, index = _parse_group(tokens, index, "bolditalic")
        return elements, index

    if name in _ITALIC_STYLE:
        elements, index = _parse_group(tokens, index, "italic")
        return elements, index

    if name in _NARY:
        # The limits bind to the operator itself; everything after them, up
        # to the end of the enclosing construct -- or the next line break,
        # whichever comes first -- is the operand.
        sub, sup, index = _read_scripts(tokens, index, style)
        body, index = _parse_sequence(
            tokens, index, stop=_stop_at_segment_end(stop), style=style)
        return [_nary(_NARY[name], sub, sup, body)], index

    if name in _LIMIT_OPERATORS:
        base = [_run(_LIMIT_OPERATORS[name], upright=True, bold=bold)]
        sub, sup, index = _read_scripts(tokens, index, style)
        if sup is not None:
            raise UnsupportedLatexError(
                f"\\{name} takes a lower limit only, not a superscript"
            )
        if sub is None:
            return base, index
        return [_limit_low(base, sub)], index

    if name in _ACCENTS:
        base, index = _parse_group(tokens, index, style)
        return [_accent(_ACCENTS[name], base)], index

    if name in ("overline", "underline"):
        base, index = _parse_group(tokens, index, style)
        position = "top" if name == "overline" else "bot"
        return [_overline(base, position=position)], index

    if name == "binom":
        top, index = _parse_group(tokens, index, style)
        bottom, index = _parse_group(tokens, index, style)
        return [_delimiter("(", ")", [_fraction(top, bottom, no_bar=True)])], index

    if name == "substack":
        # One column, one line per `\\` -- the shape a stacked n-ary limit
        # such as `\sum_{\substack{i < j \\ i \in S}}` needs.
        probe = _skip_space(tokens, index)
        if probe >= len(tokens) or tokens[probe][0] != "open":
            raise UnsupportedLatexError(
                "\\substack needs a brace group, as in \\substack{a \\\\ b}"
            )
        lines, index = _parse_lines(tokens, probe + 1, style=style)
        if index >= len(tokens) or tokens[index][0] != "close":
            raise UnsupportedLatexError("Unbalanced braces: unclosed '{'")
        return [_matrix([[line] for line in lines])], index + 1

    if name == "left":
        begin, index = _read_delimiter(tokens, index, "left")
        children, index = _parse_sequence(
            tokens, index,
            stop=lambda kind, value: (kind, value) == _RIGHT_COMMAND,
            style=style,
        )
        if index >= len(tokens) or tokens[index] != _RIGHT_COMMAND:
            raise UnsupportedLatexError("\\left without a matching \\right")
        end, index = _read_delimiter(tokens, index + 1, "right")
        return [_delimiter(begin, end, children)], index

    if name == "right":
        raise UnsupportedLatexError("\\right without a matching \\left")

    if name == "begin":
        return _parse_environment(tokens, index, style)

    if name == "end":
        raise UnsupportedLatexError("\\end without a matching \\begin")

    # `_parse_lines` intercepts both of these wherever they can carry
    # meaning, so reaching them here means they turned up somewhere that
    # takes a single atom -- `x^\\`, `\frac\over x` -- where TeX has nothing
    # to attach them to either. Say which half is missing rather than
    # falling through to the generic "unsupported command" below.
    if name == "\\":
        raise UnsupportedLatexError(
            "A line break \\\\ has nothing to break here"
        )

    if name in _INFIX:
        raise UnsupportedLatexError(
            f"\\{name} needs an expression on both sides of it"
        )

    # Checked before the remaining fallback tables (_SYMBOLS, _SPACING,
    # _ESCAPED, _UPRIGHT_FUNCTIONS) -- not just "the symbol table" -- so a
    # construct this module handles only in another form never silently
    # degrades into something that merely looks plausible. Everything above
    # this point is a construct branch for something already implemented;
    # _ENVIRONMENT_ONLY only needs to stay disjoint from the tables it
    # precedes.
    if name in _ENVIRONMENT_ONLY:
        raise UnsupportedLatexError(
            f"\\{name} works only as an environment here: write "
            f"\\begin{{{name}}} ... \\end{{{name}}}"
        )

    if name in _SYMBOLS:
        return [_run(_SYMBOLS[name], italic=style == "bolditalic", bold=bold)], index

    if name in _SPACING:
        spacing = _SPACING[name]
        if not spacing:
            return [], index
        return [_run(spacing)], index

    if name in _ESCAPED:
        return [_run(_ESCAPED[name], bold=bold)], index

    if name in _UPRIGHT_FUNCTIONS:
        return [_run(name, upright=True, bold=bold)], index

    raise UnsupportedLatexError(f"Unsupported LaTeX command: \\{name}")


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------


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
