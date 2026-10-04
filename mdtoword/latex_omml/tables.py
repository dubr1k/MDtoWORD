"""Classification tables: which commands and environments belong to
which construct.

Each table names a family of commands -- big operators, accents, math
alphabets, colours, matrix and line environments... -- and, where it
matters, the value the construct is built from.  The command dispatch
table in :mod:`.commands` is assembled from these sets.
"""

from __future__ import annotations

_UPRIGHT_FUNCTIONS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "coth", "log", "ln", "lg", "exp", "det", "dim",
    "ker", "deg", "gcd", "hom", "arg", "Pr",
}

_FRACTIONS = {"frac", "dfrac", "tfrac", "cfrac"}
_BINOMIALS = {"binom", "dbinom", "tbinom"}

# Text-mode commands: their argument is literal text, set as Word "normal
# text" (`<m:nor>`).  The value says how that text is styled on top: bold
# and italic go to the run's Word formatting (`<w:b/>`, `<w:i/>`), because
# `<m:nor>` and `<m:sty>` are mutually exclusive in OMML; the sans-serif and
# monospace faces have no Word-formatting spelling short of naming a font,
# so those use the math alphabet (`<m:scr>`) instead.
_TEXT_COMMANDS = {
    "text": None, "textrm": None, "textnormal": None, "textup": None,
    "mbox": None, "hbox": None,
    "textbf": "bold", "textit": "italic", "emph": "italic",
    "textsf": "sans-serif", "texttt": "monospace",
}

# Math alphabets carried by `<m:scr>` -- the exact values ECMA-376's
# ST_Script defines, and what Word's own MathML import writes.
_SCRIPT_ALPHABETS = {
    "mathbb": "double-struck", "Bbb": "double-struck",
    "mathbbm": "double-struck",
    "mathcal": "script", "mathscr": "script",
    "mathfrak": "fraktur",
    "mathsf": "sans-serif",
    "mathtt": "monospace",
}
_BOLD_ITALIC_STYLE = {"boldsymbol", "bm", "pmb"}

# Big operators.  Each takes optional `_`/`^` limits and then swallows the
# rest of the enclosing group as its operand.
_NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐",
    "int": "∫", "iint": "∬", "iiint": "∭", "iiiint": "⨌",
    "oint": "∮", "oiint": "∯", "oiiint": "∰",
    "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁",
    "bigotimes": "⨂", "bigodot": "⨀", "biguplus": "⨄", "bigsqcup": "⨆",
    "bigvee": "⋁", "bigwedge": "⋀",
}

# Integrals keep their limits beside the sign; every other big operator
# stacks them above and below.
_INTEGRAL_CHARACTERS = {"∫", "∬", "∭", "⨌", "∮", "∯", "∰"}

# Operators whose subscript sits underneath rather than beside them -- the
# same list Word's own LaTeX import treats as limit-style.
_LIMIT_OPERATORS = {
    "lim": "lim", "limsup": "lim sup", "liminf": "lim inf",
    "max": "max", "min": "min", "sup": "sup", "inf": "inf",
    "injlim": "inj lim", "projlim": "proj lim",
}

# Combining marks: each one composes with the character it follows.
_ACCENTS = {
    "hat": "̂", "widehat": "̂", "tilde": "̃",
    "widetilde": "̃", "bar": "̄", "vec": "⃗",
    "dot": "̇", "ddot": "̈", "dddot": "⃛", "ddddot": "⃜",
    "acute": "́", "grave": "̀", "check": "̌", "breve": "̆",
    "mathring": "̊",
}

# Stretchy characters drawn over or under their argument: <m:groupChr>.
# The flag says whether `^`/`_` after the construct become limits stacked
# above/below it -- `\underbrace{a+b}_{n}` -- as they do in TeX, where the
# braces are `\mathop`s with `\limits`.
_GROUP_CHARACTERS = {
    "overbrace": ("⏞", "top", True), "underbrace": ("⏟", "bot", True),
    "overparen": ("⏜", "top", True), "underparen": ("⏝", "bot", True),
    "overbracket": ("⎴", "top", True), "underbracket": ("⎵", "bot", True),
    "overrightarrow": ("→", "top", False), "overleftarrow": ("←", "top", False),
    "overleftrightarrow": ("↔", "top", False),
    "underrightarrow": ("→", "bot", False), "underleftarrow": ("←", "bot", False),
    "underleftrightarrow": ("↔", "bot", False),
}

# `\xrightarrow[below]{above}` and its siblings: an arrow that stretches to
# fit the text written on it.
_EXTENSIBLE_ARROWS = {
    "xrightarrow": "→", "xleftarrow": "←", "xleftrightarrow": "↔",
    "xRightarrow": "⇒", "xLeftarrow": "⇐", "xLeftrightarrow": "⇔",
    "xhookrightarrow": "↪", "xhookleftarrow": "↩", "xmapsto": "↦",
    "xrightharpoonup": "⇀", "xrightharpoondown": "⇁",
    "xleftharpoonup": "↼", "xleftharpoondown": "↽",
    "xrightleftharpoons": "⇌", "xleftrightharpoons": "⇋",
    "xlongequal": "=",
}

# `\cancel` strikes bottom-left to top-right, `\bcancel` the other way.
_CANCELS = {
    "cancel": ("strikeBLTR",), "bcancel": ("strikeTLBR",),
    "xcancel": ("strikeBLTR", "strikeTLBR"),
}

# <m:phant> flags: (show, zero width, zero ascent, zero descent).
_PHANTOMS = {
    "phantom": (False, False, False, False),
    "hphantom": (False, False, True, True),
    "vphantom": (False, True, False, False),
}

# `\big(` and its siblings: a fixed-size delimiter, emitted as a plain
# character since OMML has no "one size larger" fence.
_BIG_DELIMITERS = frozenset(
    size + side
    for size in ("big", "Big", "bigg", "Bigg")
    for side in ("", "l", "r", "m")
)

# Spacing-class wrappers: Word does its own spacing, so only the content
# matters.
_MATH_CLASSES = frozenset({
    "mathrel", "mathbin", "mathord", "mathpunct", "mathopen", "mathclose",
    "mathinner",
})

# Declarations that change the style of everything after them, up to the
# end of the enclosing group, line or cell.
_NO_OP_SWITCHES = frozenset({
    "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle",
    "tiny", "scriptsize", "footnotesize", "small", "normalsize",
    "large", "Large", "LARGE", "huge", "Huge",
})

# Equation numbering: `\tag` adds a number at the end of its line, the
# others only affect numbering, which Word does not do inside an equation.
_NUMBERING = frozenset({"tag", "label", "nonumber", "notag"})

# Named colours, with the values MathJax and KaTeX -- which render the
# Markdown these formulas usually come from -- give them (CSS colours).
_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000",
    "green": "008000", "blue": "0000FF", "cyan": "00FFFF",
    "magenta": "FF00FF", "yellow": "FFFF00", "gray": "808080",
    "grey": "808080", "darkgray": "A9A9A9", "lightgray": "D3D3D3",
    "orange": "FFA500", "purple": "800080", "brown": "A52A2A",
    "lime": "00FF00", "olive": "808000", "teal": "008080",
    "navy": "000080", "maroon": "800000", "pink": "FFC0CB",
    "violet": "EE82EE",
}

_MATRIX_DELIMITERS = {
    "matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
    "smallmatrix": ("", ""), "psmallmatrix": ("(", ")"),
    "bsmallmatrix": ("[", "]"), "Bsmallmatrix": ("{", "}"),
    "vsmallmatrix": ("|", "|"), "Vsmallmatrix": ("‖", "‖"),
}

# `cases` and its mathtools relatives: left-aligned columns behind a brace.
# `rcases` puts the brace on the right.
_CASES_DELIMITERS = {
    "cases": ("{", ""), "dcases": ("{", ""),
    "rcases": ("", "}"), "drcases": ("", "}"),
}

# Multi-line environments that become an <m:eqArr>.  "align" ones turn
# `&` into alignment points; "gather" ones centre their lines and have no
# `&` at all; "equation" ones are a single formula, where `&` keeps the
# ordinary multi-line rule.
_LINE_ENVIRONMENTS = {
    "aligned": "align", "split": "align", "alignedat": "align",
    "align": "align", "align*": "align", "flalign": "align",
    "flalign*": "align", "alignat": "align", "alignat*": "align",
    "eqnarray": "align", "eqnarray*": "align",
    "gathered": "gather", "gather": "gather", "gather*": "gather",
    "multline": "gather", "multline*": "gather",
    "equation": "equation", "equation*": "equation",
    "displaymath": "equation",
}
# These take the number of `&`-column pairs as an argument, which only
# matters to TeX's own layout.
_COLUMN_COUNT_ENVIRONMENTS = frozenset({"alignedat", "alignat", "alignat*"})
# These take an optional `[t]`/`[b]`/`[c]` vertical position.
_POSITIONED_ENVIRONMENTS = frozenset({"aligned", "alignedat", "gathered"})


# Infix commands: each splits the group it appears in, taking everything to
# its left as the numerator and everything to its right as the denominator.
# That is why they cannot live in `_parse_command` with the prefix commands
# -- by the time it runs, the numerator has already been parsed and emitted.
_INFIX = frozenset({"over", "atop", "choose", "brace", "brack"})

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
