"""Character tables: what each symbol, spacing, escape and delimiter
command stands for.

Pure data with no imports, shared by the tokenizer-level readers, the
text-mode readers and the command handlers.
"""

from __future__ import annotations

_SYMBOLS = {
    # Greek.  `\epsilon`/`\phi` are the lunate/straight forms and the
    # `var` spellings the open/curly ones -- the glyphs LaTeX, MathJax and
    # KaTeX all draw, and the code points Word's own LaTeX import uses.
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ϵ", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "varkappa": "ϰ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ",
    "sigma": "σ", "varsigma": "ς", "tau": "τ",
    "upsilon": "υ", "phi": "ϕ", "varphi": "φ", "chi": "χ",
    "psi": "ψ", "omega": "ω", "digamma": "ϝ",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    # Letter-like symbols.
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "hbar": "ℏ", "hslash": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ",
    "aleph": "ℵ", "beth": "ℶ", "gimel": "ℷ", "daleth": "ℸ",
    "wp": "℘", "mho": "℧", "Finv": "Ⅎ", "Game": "⅁", "eth": "ð",
    "imath": "ı", "jmath": "ȷ", "complement": "∁",
    # Binary operators.
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "⋅", "ast": "∗", "star": "⋆",
    "circ": "∘", "bullet": "∙", "oplus": "⊕", "otimes": "⊗",
    "ominus": "⊖", "oslash": "⊘", "odot": "⊙", "circledast": "⊛",
    "circledcirc": "⊚", "circleddash": "⊝",
    "boxplus": "⊞", "boxminus": "⊟", "boxtimes": "⊠", "boxdot": "⊡",
    "cup": "∪", "cap": "∩", "uplus": "⊎", "sqcup": "⊔", "sqcap": "⊓",
    "Cup": "⋓", "Cap": "⋒",
    "vee": "∨", "wedge": "∧", "land": "∧", "lor": "∨",
    "setminus": "∖", "smallsetminus": "∖", "wr": "≀", "amalg": "⨿",
    "dagger": "†", "ddagger": "‡", "dag": "†", "ddag": "‡",
    "diamond": "⋄", "bigtriangleup": "△", "bigtriangledown": "▽",
    "triangleleft": "◁", "triangleright": "▷",
    "lhd": "⊲", "rhd": "⊳", "unlhd": "⊴", "unrhd": "⊵",
    "barwedge": "⊼", "veebar": "⊻", "doublebarwedge": "⩞", "dotplus": "∔",
    "ltimes": "⋉", "rtimes": "⋊", "leftthreetimes": "⋋",
    "rightthreetimes": "⋌", "curlyvee": "⋎", "curlywedge": "⋏",
    "intercal": "⊺", "divideontimes": "⋇",
    # Relations.
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "lt": "<", "gt": ">",
    "leqq": "≦", "geqq": "≧", "leqslant": "⩽", "geqslant": "⩾",
    "lesssim": "≲", "gtrsim": "≳", "lessapprox": "⪅", "gtrapprox": "⪆",
    "lessgtr": "≶", "gtrless": "≷", "lll": "⋘", "ggg": "⋙",
    "neq": "≠", "ne": "≠", "approx": "≈", "approxeq": "≊", "equiv": "≡",
    "sim": "∼", "simeq": "≃", "backsim": "∽", "cong": "≅",
    "propto": "∝", "varpropto": "∝", "asymp": "≍", "bowtie": "⋈",
    "doteq": "≐", "doteqdot": "≑", "triangleq": "≜", "coloneqq": "≔",
    "eqqcolon": "≕", "circeq": "≗", "eqcirc": "≖", "bumpeq": "≏",
    "Bumpeq": "≎", "risingdotseq": "≓", "fallingdotseq": "≒",
    "ll": "≪", "gg": "≫",
    "prec": "≺", "succ": "≻", "preceq": "⪯", "succeq": "⪰",
    "precsim": "≾", "succsim": "≿",
    "in": "∈", "notin": "∉", "ni": "∋", "owns": "∋",
    "subset": "⊂", "subseteq": "⊆", "supset": "⊃", "supseteq": "⊇",
    "subsetneq": "⊊", "supsetneq": "⊋", "Subset": "⋐", "Supset": "⋑",
    "sqsubset": "⊏", "sqsupset": "⊐", "sqsubseteq": "⊑", "sqsupseteq": "⊒",
    "vdash": "⊢", "dashv": "⊣", "models": "⊨", "vDash": "⊨",
    "Vdash": "⊩", "Vvdash": "⊪",
    "mid": "∣", "parallel": "∥", "perp": "⊥",
    "smile": "⌣", "frown": "⌢", "between": "≬", "pitchfork": "⋔",
    "vartriangleleft": "⊲", "vartriangleright": "⊳",
    "trianglelefteq": "⊴", "trianglerighteq": "⊵",
    # Negated relations.
    "nmid": "∤", "nparallel": "∦", "nleq": "≰", "ngeq": "≱",
    "nless": "≮", "ngtr": "≯", "nsubseteq": "⊈", "nsupseteq": "⊉",
    "ncong": "≇", "nsim": "≁", "nprec": "⊀", "nsucc": "⊁",
    "npreceq": "⋠", "nsucceq": "⋡", "nvdash": "⊬", "nvDash": "⊭",
    "nVdash": "⊮", "nVDash": "⊯", "nexists": "∄",
    "ntriangleleft": "⋪", "ntriangleright": "⋫",
    "ntrianglelefteq": "⋬", "ntrianglerighteq": "⋭",
    # Logic.
    "forall": "∀", "exists": "∃", "neg": "¬", "lnot": "¬",
    "top": "⊤", "bot": "⊥", "therefore": "∴", "because": "∵",
    "implies": "⟹", "impliedby": "⟸", "iff": "⟺",
    # Sets.
    "emptyset": "∅", "varnothing": "∅",
    # Arrows.
    "rightarrow": "→", "to": "→", "leftarrow": "←", "gets": "←",
    "leftrightarrow": "↔", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "Leftrightarrow": "⇔", "mapsto": "↦", "longmapsto": "⟼",
    "longrightarrow": "⟶", "longleftarrow": "⟵",
    "longleftrightarrow": "⟷", "Longrightarrow": "⟹",
    "Longleftarrow": "⟸", "Longleftrightarrow": "⟺",
    "uparrow": "↑", "downarrow": "↓", "updownarrow": "↕",
    "Uparrow": "⇑", "Downarrow": "⇓", "Updownarrow": "⇕",
    "nearrow": "↗", "searrow": "↘", "nwarrow": "↖", "swarrow": "↙",
    "hookrightarrow": "↪", "hookleftarrow": "↩",
    "rightharpoonup": "⇀", "rightharpoondown": "⇁",
    "leftharpoonup": "↼", "leftharpoondown": "↽",
    "rightleftharpoons": "⇌", "leftrightharpoons": "⇋",
    "upharpoonleft": "↿", "upharpoonright": "↾", "restriction": "↾",
    "downharpoonleft": "⇃", "downharpoonright": "⇂",
    "rightrightarrows": "⇉", "leftleftarrows": "⇇",
    "rightleftarrows": "⇄", "leftrightarrows": "⇆",
    "upuparrows": "⇈", "downdownarrows": "⇊",
    "twoheadrightarrow": "↠", "twoheadleftarrow": "↞",
    "rightarrowtail": "↣", "leftarrowtail": "↢",
    "rightsquigarrow": "⇝", "leadsto": "⇝", "leftrightsquigarrow": "↭",
    "circlearrowleft": "↺", "circlearrowright": "↻",
    "curvearrowleft": "↶", "curvearrowright": "↷",
    "dashrightarrow": "⇢", "dashleftarrow": "⇠",
    "looparrowright": "↬", "looparrowleft": "↫",
    "Lsh": "↰", "Rsh": "↱", "Lleftarrow": "⇚", "Rrightarrow": "⇛",
    "multimap": "⊸",
    "nleftarrow": "↚", "nrightarrow": "↛", "nleftrightarrow": "↮",
    "nLeftarrow": "⇍", "nRightarrow": "⇏", "nLeftrightarrow": "⇎",
    # Dots.
    "ldots": "…", "dots": "…", "dotsc": "…", "dotso": "…",
    "cdots": "⋯", "dotsb": "⋯", "dotsm": "⋯", "dotsi": "⋯",
    "vdots": "⋮", "ddots": "⋱", "iddots": "⋰",
    "ldotp": ".", "cdotp": "⋅", "colon": ":",
    # Delimiters written outside `\left`...`\right`: plain characters.
    "|": "‖", "vert": "|", "Vert": "‖", "lvert": "|", "rvert": "|",
    "lVert": "‖", "rVert": "‖", "langle": "⟨", "rangle": "⟩",
    "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈", "rceil": "⌉",
    "lbrace": "{", "rbrace": "}", "lbrack": "[", "rbrack": "]",
    "llbracket": "⟦", "rrbracket": "⟧", "backslash": "\\",
    "ulcorner": "⌜", "urcorner": "⌝", "llcorner": "⌞", "lrcorner": "⌟",
    # Miscellaneous.
    "prime": "′", "degree": "°", "angle": "∠", "measuredangle": "∡",
    "sphericalangle": "∢", "surd": "√",
    "square": "□", "Box": "□", "blacksquare": "■",
    "triangle": "△", "vartriangle": "△", "triangledown": "▽",
    "blacktriangle": "▴", "blacktriangledown": "▾",
    "blacktriangleleft": "◀", "blacktriangleright": "▶",
    "Diamond": "◇", "lozenge": "◊", "blacklozenge": "⧫",
    "bigcirc": "◯", "bigstar": "★", "checkmark": "✓", "maltese": "✠",
    "flat": "♭", "natural": "♮", "sharp": "♯",
    "clubsuit": "♣", "diamondsuit": "♢", "heartsuit": "♡", "spadesuit": "♠",
    "circledR": "®", "copyright": "©", "pounds": "£", "yen": "¥",
    "euro": "€", "S": "§", "P": "¶", "And": "&",
}

# Variant capital Greek: the same letters as `\Gamma`..., set in math italic.
_ITALIC_SYMBOLS = {
    "varGamma": "Γ", "varDelta": "Δ", "varTheta": "Θ", "varLambda": "Λ",
    "varXi": "Ξ", "varPi": "Π", "varSigma": "Σ", "varUpsilon": "Υ",
    "varPhi": "Φ", "varPsi": "Ψ", "varOmega": "Ω",
}

# Spacing commands become real spaces; Word does its own math spacing anyway.
# Negative spaces have no OMML equivalent and become nothing.
_SPACING = {",": " ", ";": " ", ":": " ", ">": " ", "!": "", " ": " ",
            "quad": " ", "qquad": "  ", "enspace": " ", "space": " ",
            "thinspace": " ", "medspace": " ", "thickspace": " ",
            "negthinspace": "", "negmedspace": "", "negthickspace": ""}

# `\%` and friends: the backslash only escapes LaTeX's own syntax.
_ESCAPED = {"{": "{", "}": "}", "%": "%", "$": "$",
            "&": "&", "#": "#", "_": "_"}

# Commands that stand for a plain character inside `\text{...}` too.
_TEXT_SYMBOLS = {
    "ldots": "…", "dots": "…", "textbackslash": "\\",
    "textasciitilde": "~", "textasciicircum": "^", "textunderscore": "_",
    "S": "§", "P": "¶", "dag": "†", "ddag": "‡", "copyright": "©",
    "pounds": "£", "euro": "€", "yen": "¥", "textdegree": "°",
}


# What may follow `\left` and `\right`.  `.` is LaTeX's "no delimiter here",
# which OMML spells as an empty begChr/endChr.
_DELIMITER_CHARACTERS = {"(": "(", ")": ")", "[": "[", "]": "]",
                         "|": "|", "/": "/", ".": "", "<": "⟨", ">": "⟩"}
_DELIMITER_COMMANDS = {
    "{": "{", "}": "}", "|": "‖", "backslash": "\\",
    "lbrace": "{", "rbrace": "}", "langle": "⟨", "rangle": "⟩",
    "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈", "rceil": "⌉",
    "vert": "|", "Vert": "‖", "lvert": "|", "rvert": "|",
    "lVert": "‖", "rVert": "‖", "lbrack": "[", "rbrack": "]",
    "llbracket": "⟦", "rrbracket": "⟧",
    "uparrow": "↑", "downarrow": "↓", "updownarrow": "↕",
    "Uparrow": "⇑", "Downarrow": "⇓", "Updownarrow": "⇕",
}
