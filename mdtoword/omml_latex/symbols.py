"""Character tables: Unicode characters, operators, fences and accents to LaTeX."""

from __future__ import annotations

# Characters that map to a LaTeX command. Where the forward converter maps
# several commands to one character (``\le``/``\leq``) the canonical one is
# used; the round trip only needs the character to come back.
_SYMBOLS: dict[str, str] = {
    # Greek, lowercase. The variant pairs follow the glyphs LaTeX draws:
    # \epsilon is the lunate U+03F5, \varepsilon the open U+03B5; \phi is
    # the stroked U+03D5, \varphi the loopy U+03C6.
    "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "varepsilon",
    "ϵ": "epsilon", "ζ": "zeta", "η": "eta", "θ": "theta", "ϑ": "vartheta",
    "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "µ": "mu", "ν": "nu",
    "ξ": "xi", "π": "pi", "ϖ": "varpi", "ρ": "rho", "ϱ": "varrho",
    "σ": "sigma", "ς": "varsigma", "τ": "tau", "υ": "upsilon", "ϕ": "phi",
    "φ": "varphi", "χ": "chi", "ψ": "psi", "ω": "omega",
    # Greek, uppercase (only those that differ from Latin capitals).
    "Γ": "Gamma", "Δ": "Delta", "∆": "Delta", "Θ": "Theta", "Λ": "Lambda",
    "Ξ": "Xi", "Π": "Pi", "Σ": "Sigma", "Υ": "Upsilon", "Φ": "Phi",
    "Ψ": "Psi", "Ω": "Omega",
    # Operators and relations.
    "∞": "infty", "∂": "partial", "∇": "nabla", "×": "times", "÷": "div",
    "±": "pm", "∓": "mp", "⋅": "cdot", "·": "cdot", "∗": "ast", "⋆": "star",
    "≤": "leq", "≥": "geq", "≠": "neq", "≈": "approx", "≡": "equiv",
    "∼": "sim", "≃": "simeq", "≅": "cong", "∝": "propto", "≪": "ll",
    "≫": "gg", "∈": "in", "∉": "notin", "∋": "ni", "⊂": "subset",
    "⊆": "subseteq", "⊃": "supset", "⊇": "supseteq", "∪": "cup", "∩": "cap",
    "∅": "emptyset", "∖": "setminus", "∀": "forall", "∃": "exists",
    "∄": "nexists", "¬": "neg", "∧": "land", "∨": "lor", "→": "to",
    "←": "leftarrow", "↔": "leftrightarrow", "⇒": "Rightarrow",
    "⇐": "Leftarrow", "⇔": "Leftrightarrow", "↦": "mapsto",
    "⟶": "longrightarrow", "⟵": "longleftarrow", "⟹": "Longrightarrow",
    "⟸": "Longleftarrow", "⟺": "Longleftrightarrow", "↑": "uparrow",
    "↓": "downarrow", "⇑": "Uparrow", "⇓": "Downarrow",
    "…": "ldots", "⋯": "cdots", "⋮": "vdots", "⋱": "ddots", "°": "degree",
    "ℏ": "hbar", "ℓ": "ell", "ℜ": "Re", "ℑ": "Im", "ℵ": "aleph", "℘": "wp",
    "∠": "angle", "⊥": "perp", "∥": "parallel", "∘": "circ", "∙": "bullet",
    "⊕": "oplus", "⊗": "otimes", "⊙": "odot", "⊖": "ominus", "≺": "prec",
    "≻": "succ", "⪯": "preceq", "⪰": "succeq", "⊢": "vdash", "⊨": "models",
    "∣": "mid", "∤": "nmid", "†": "dagger", "‡": "ddagger", "∴": "therefore",
    "∵": "because", "⊤": "top", "□": "square", "△": "triangle",
    "≲": "lesssim", "≳": "gtrsim", "≐": "doteq", "≜": "triangleq",
    "ı": "imath", "ȷ": "jmath", "√": "surd",
    # Big operators met as plain characters rather than as m:nary.
    "∑": "sum", "∏": "prod", "∐": "coprod", "∫": "int", "∬": "iint",
    "∭": "iiint", "∮": "oint", "⋃": "bigcup", "⋂": "bigcap",
}

# Characters with a fixed LaTeX spelling that is not a bare command.
_LITERALS: dict[str, str] = {
    "{": r"\{", "}": r"\}", "#": r"\#", "%": r"\%", "&": r"\&", "$": r"\$",
    "_": r"\_", "\\": r"\backslash", "^": r"\hat{}", "~": r"\sim",
    "−": "-", "′": "'", "″": "''", "‴": "'''", "∶": ":",
    # Invisible operators Word inserts between a function name and its
    # argument or between juxtaposed factors.
    "⁡": "", "⁢": "", "⁣": "", "⁤": "", "​": "",
}

# Text-mode escapes, for the inside of \text{...}.
_TEXT_ESCAPES: dict[str, str] = {
    "{": r"\{", "}": r"\}", "#": r"\#", "%": r"\%", "&": r"\&", "$": r"\$",
    "_": r"\_", "\\": r"\textbackslash{}", "^": r"\^{}", "~": r"\~{}",
}

# Upright function names the forward converter knows as commands. Others are
# written as \operatorname{...} -- valid LaTeX that also converts back.
_FUNCTIONS = frozenset({
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "det", "dim", "ker",
    "deg", "gcd", "hom", "arg", "max", "min", "sup", "inf", "lim",
})
_MULTIWORD_FUNCTIONS = {"lim sup": "limsup", "lim inf": "liminf"}

_NARY: dict[str, str] = {
    "∑": "sum", "∏": "prod", "∐": "coprod", "∫": "int", "∬": "iint",
    "∭": "iiint", "∮": "oint", "∯": "oiint", "∰": "oiiint", "⋃": "bigcup",
    "⋂": "bigcap", "⨁": "bigoplus", "⨂": "bigotimes", "⨀": "bigodot",
    "⋁": "bigvee", "⋀": "bigwedge", "⨄": "biguplus", "⨆": "bigsqcup",
}

# Combining marks (what the forward converter writes) and the spacing forms
# Word sometimes uses instead.
_ACCENTS: dict[str, str] = {
    "̂": "hat", "̃": "tilde", "̄": "bar", "̅": "bar",
    "⃗": "vec", "̇": "dot", "̈": "ddot", "⃛": "dddot",
    "̌": "check", "̆": "breve", "́": "acute", "̀": "grave",
    "⃖": "overleftarrow", "⃡": "overleftrightarrow",
    "^": "hat", "ˆ": "hat", "~": "tilde", "˜": "tilde", "¯": "bar",
    "→": "vec", "˙": "dot", "¨": "ddot", "ˇ": "check", "˘": "breve",
    "´": "acute", "`": "grave",
}

_INTEGRALS = frozenset("∫∬∭∮∯∰")

# Fence characters of m:d, as \left/\right understand them.
_DELIMITERS: dict[str, str] = {
    "": ".", "(": "(", ")": ")", "[": "[", "]": "]", "{": r"\{", "}": r"\}",
    "|": "|", "‖": r"\|", "⟨": r"\langle", "〈": r"\langle", "⟩": r"\rangle",
    "〉": r"\rangle", "⌊": r"\lfloor", "⌋": r"\rfloor", "⌈": r"\lceil",
    "⌉": r"\rceil", "/": "/", "\\": r"\backslash",
}

# Delimiters wrapping a lone m:m that make it one of the matrix environments.
_MATRIX_ENVIRONMENTS: dict[tuple[str, str], str] = {
    ("(", ")"): "pmatrix", ("[", "]"): "bmatrix", ("{", "}"): "Bmatrix",
    ("|", "|"): "vmatrix", ("‖", "‖"): "Vmatrix", ("{", ""): "cases",
}

_COLUMN_LETTERS = {"left": "l", "center": "c", "right": "r"}

# Mathematical alphanumeric styles, from the Unicode character name.
_ALPHANUMERIC_STYLES = (
    ("BOLD ITALIC", "boldsymbol"), ("BOLD SCRIPT", "mathcal"),
    ("BOLD FRAKTUR", "mathfrak"), ("SANS-SERIF", "mathsf"),
    ("DOUBLE-STRUCK", "mathbb"), ("MONOSPACE", "mathtt"), ("FRAKTUR", "mathfrak"),
    ("BLACK-LETTER", "mathfrak"), ("SCRIPT", "mathcal"), ("BOLD", "mathbf"),
    ("ITALIC", None), ("PLANCK", None),
)

# Relations get a space on either side: LaTeX ignores it, readers do not.
_RELATIONS = frozenset("=<>≤≥≠≈≡∼≃≅∝≪≫∈∉∋⊂⊆⊃⊇→←↔⇒⇐⇔↦⟶⟵⟹⟸⟺∣")

_SPACE_RUNS = {" ": r"\,", "  ": r"\qquad"}
