"""Convert a LaTeX math string into Office Math Markup Language (OMML).

The result is a real Word equation that stays editable in Word's equation
editor, not a text fallback with dollar signs around it.

Only constructs listed in ``_parse_command`` are understood.  Anything else
raises :class:`UnsupportedLatexError` naming the offending construct, so the
caller can report it instead of silently emitting wrong output.
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

# Recognised, but deliberately not implemented here.  A later task adds these;
# until then they must fail loudly rather than be dropped or mis-rendered.
_NOT_YET = {
    "sum": "big operator", "prod": "big operator", "coprod": "big operator",
    "bigcup": "big operator", "bigcap": "big operator",
    "int": "integral", "iint": "integral", "iiint": "integral",
    "oint": "integral",
    "lim": "limit", "limsup": "limit", "liminf": "limit",
    "left": "delimiter", "right": "delimiter",
    "begin": "environment", "end": "environment",
    "matrix": "matrix", "pmatrix": "matrix", "bmatrix": "matrix",
    "vmatrix": "matrix", "Vmatrix": "matrix", "array": "matrix",
    "cases": "matrix", "substack": "matrix",
    "binom": "binomial", "choose": "binomial",
    "over": "fraction", "atop": "fraction",
    "hat": "accent", "widehat": "accent", "bar": "accent", "vec": "accent",
    "tilde": "accent", "widetilde": "accent", "dot": "accent",
    "ddot": "accent", "overline": "accent", "underline": "accent",
    "\\": "line break",
}

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
        if upright:
            nor = _el("nor")
            nor.set(qn("m:val"), "1")
            properties.append(nor)
        if italic and not upright:
            sty = _el("sty")
            sty.set(qn("m:val"), "bi" if bold else "i")
            properties.append(sty)
        elif bold:
            sty = _el("sty")
            sty.set(qn("m:val"), "b")
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


def _parse_sequence(tokens: list, index: int, stop: Stop = None,
                    style: Optional[str] = None) -> tuple:
    """Parse tokens into elements until `stop`, a closing brace, or the end.

    Returns (elements, index) with `index` left ON the terminating token.
    """
    elements: list = []
    while index < len(tokens):
        kind, value = tokens[index]
        if kind == "space":
            index += 1
            continue
        if kind == "close":
            break
        if stop is not None and stop(kind, value):
            break
        atom, index = _parse_atom(tokens, index, style)
        sub, sup, index = _read_scripts(tokens, index, style)
        if sub is None and sup is None:
            elements.extend(atom)
        else:
            elements.append(_script(atom, sub, sup))
    return elements, index


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


def _parse_atom(tokens: list, index: int, style: Optional[str] = None) -> tuple:
    """Parse one atom: a group, a character, or a command. Returns (elements, index)."""
    kind, value = tokens[index]
    bold = style in ("bold", "bolditalic")
    if kind == "open":
        return _parse_group(tokens, index, style)
    if kind == "command":
        return _parse_command(tokens, index, style)
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
            "'&' is only meaningful inside a matrix environment"
        )
    raise UnsupportedLatexError(f"Could not read {value!r} in the formula")


def _parse_command(tokens: list, index: int, style: Optional[str] = None) -> tuple:
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

    # Checked before the symbol table so a Task 3 construct never silently
    # degrades into something that merely looks plausible.
    if name in _NOT_YET:
        raise UnsupportedLatexError(
            f"LaTeX {_NOT_YET[name]} is not supported yet: \\{name}"
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
