"""Constants shared by the renderer's mixins.

Regular expressions that recognise Markdown, LaTeX and HTML constructs,
size caps, fonts, shading colours, borders and indents, the localized
GitHub alert titles and the HTML tag tables the raw-HTML handling consults.
Nothing in this module holds state.
"""

from __future__ import annotations

import re

from docx.shared import Cm, Pt, RGBColor

_TASK_PREFIX = re.compile(r"^\[([ xX])\]\s+")
_CYRILLIC = re.compile(r"[Ѐ-ӿ]")
# Commands whose argument is literal text, e.g. "\text{путь}". Their content
# is stripped before the bare-Cyrillic prose guard below runs, so a formula
# that writes a Russian word the correct way -- inside \text{...} -- is not
# mistaken for prose; Cyrillic left over *outside* one of these still warns.
_TEXT_COMMAND = re.compile(
    r"\\(?:text|mathrm|textrm|textnormal|operatorname)\{[^{}]*\}"
)
_BLACK = RGBColor(0, 0, 0)

# Cap on how much of a remote image response we will read into memory. Applies
# regardless of ``allow_remote_images`` -- a large or hostile URL should not be
# able to exhaust memory just because fetching is permitted.
_MAX_REMOTE_IMAGE_BYTES = 20 * 1024 * 1024
# Local images are read whole too; the same ceiling keeps one absurd file
# from taking the process down.
_MAX_LOCAL_IMAGE_BYTES = 50 * 1024 * 1024

# ``latex_omml`` parses these environments itself, so they are passed through
# with their ``\begin``/``\end`` wrapper intact.
_MATRIX_ENVIRONMENTS = frozenset(
    {"matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix"}
)
# The amsmath plugin hands us the environment complete with its wrapper, e.g.
# "\begin{align}\na &= b \\\\\nc &= d\n\end{align}".
_AMSMATH_WRAPPER = re.compile(
    r"^\\begin\{(?P<environment>[A-Za-z]+)\*?\}"
    r"(?P<body>.*)"
    r"\\end\{(?P=environment)\*?\}$",
    re.DOTALL,
)
# ``alignat`` and ``flalign`` add a column-count argument right after the
# opening tag, e.g. "\begin{alignat}{2}". Every other environment keeps a
# leading brace group as part of its body -- ``\begin{equation}{\bf x}...``
# is ordinary physics LaTeX, not an argument, and must not be eaten.
_COLUMN_ARGUMENT_ENVIRONMENTS = frozenset({"alignat", "flalign"})
_COLUMN_ARGUMENT = re.compile(r"^\{[^{}]*\}")
_LINE_BREAK = re.compile(r"\\\\")
# An unescaped "&" is amsmath column alignment; "\&" is a literal ampersand.
_ALIGNMENT_MARKER = re.compile(r"(?<!\\)&")

# Multiplier applied to the user's chosen body size for each heading level,
# plus whether that level is bold/italic. Word's default template only
# defines a size for Heading 1-2 and leaves 3-9 to inherit Normal's size
# verbatim, and only defines bold for 1-4 -- both of which leave everything
# from H3 down indistinguishable from body text. Every level here is forced
# bold; level 6 is additionally italic so it stays visually distinct from
# level 5, which shares its multiplier.
_HEADING_SCALE: dict[int, tuple[float, bool, bool]] = {
    1: (1.5, True, False),
    2: (1.35, True, False),
    3: (1.2, True, False),
    4: (1.1, True, False),
    5: (1.0, True, False),
    6: (1.0, True, True),
}
# A single-dollar fragment with none of these is unlikely to be a real
# formula: no LaTeX command, no super/subscript, no digit, no operator.
_MATH_INDICATOR = re.compile(r"[\\^_0-9+\-*/=<>]")

_THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme")

_CODE_FONT = "Courier New"
_INLINE_CODE_SHADING = "F2F2F2"
_CODE_BLOCK_SHADING = "F5F5F5"
_CODE_BLOCK_BORDER = ("single", 4, 4, "BFBFBF")
_QUOTE_STEP = Pt(18)
_QUOTE_BORDER = ("single", 12, 8, "A6A6A6")
_ALERT_BORDER = ("single", 24, 8, "7F7F7F")
_ALERT_SHADING = "F2F2F2"
_DEFINITION_INDENT = Cm(1)
_GOST_FIRST_LINE = Cm(1.25)
# (top, right, bottom, left) in millimetres.
_GOST_MARGINS = (20.0, 15.0, 20.0, 30.0)
_GOST_USER_MARGINS = (15.0, 15.0, 15.0, 30.0)
_DEFAULT_MARGINS = (25.4, 25.4, 25.4, 25.4)

_ALERT_MARKER = re.compile(r"^\[!(note|tip|important|warning|caution)\][ \t]*", re.IGNORECASE)
_ALERT_TITLES = {
    "note": ("Note", "Примечание"),
    "tip": ("Tip", "Совет"),
    "important": ("Important", "Важно"),
    "warning": ("Warning", "Внимание"),
    "caution": ("Caution", "Осторожно"),
}
_STARRED_TAG = re.compile(r"(?<!\\)\\tag\*\s*\{")
_TOC_MARKER = re.compile(r"^\s*(?:\[TOC\]|\[\[_TOC_\]\]|\$\{toc\})\s*$", re.IGNORECASE)
_TABLE_CAPTION = re.compile(r"^(?:(?:Table|Таблица)\s*[:.]|:)\s+", re.IGNORECASE)

# Code points XML 1.0 does not allow anywhere in a document (C0 controls
# other than tab/newline/carriage return, lone surrogates, U+FFFE/U+FFFF).
_XML_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
# Word's limit on the number of table columns.
_MAX_TABLE_COLUMNS = 63
# python-docx refuses core-property values longer than this.
_MAX_PROPERTY_LENGTH = 255
# A BCP 47-shaped language tag: "ru", "ru-RU", "zh-Hant-TW".
_LANGUAGE_TAG = re.compile(r"^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$")
# Children of w:compat that the schema orders before w:doNotExpandShiftReturn.
_COMPAT_BEFORE_SHIFT_RETURN = (
    "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading",
    "spaceForUL", "noColumnBalance", "balanceSingleByteDoubleByteWidth",
    "noExtraLineSpacing", "doNotLeaveBackslashAlone", "ulTrailSpace",
)
_HTML_COMMENT = re.compile(r"^<!--.*?-->$", re.DOTALL)
_HTML_TAG = re.compile(
    r"^<(?P<close>/)?(?P<name>[A-Za-z][A-Za-z0-9-]*)(?P<attrs>[^>]*?)(?P<self>/)?\s*>$",
    re.DOTALL,
)
_HTML_ATTR = re.compile(
    r"([A-Za-z_:][-A-Za-z0-9_:.]*)\s*(?:=\s*(\"[^\"]*\"|'[^']*'|[^\s\"'=<>`]+))?"
)
_HTML_PIECES = re.compile(r"(<!--.*?-->|</?[A-Za-z][^>]*>)", re.DOTALL)
# Inline tags that map onto run formatting.
_HTML_FORMAT = {
    "b": "bold", "strong": "bold",
    "i": "italic", "em": "italic", "var": "italic", "cite": "italic", "dfn": "italic",
    "u": "underline", "ins": "underline",
    "s": "strike", "del": "strike", "strike": "strike",
    "sub": "sub", "sup": "sup", "mark": "mark",
    "code": "code", "kbd": "code", "tt": "code", "samp": "code",
}
# Wrappers whose removal loses nothing but presentation: unwrapped silently.
_HTML_TRANSPARENT = frozenset(
    {"span", "small", "big", "abbr", "font", "nobr", "wbr", "bdi", "bdo", "time", "data", "q"}
)
# Block-level tags inside an HTML block: each one starts a new paragraph.
_HTML_BLOCK_BREAKS = frozenset(
    {
        "p", "div", "details", "summary", "section", "article", "header", "footer",
        "main", "aside", "nav", "figure", "figcaption", "blockquote", "pre", "center",
        "ul", "ol", "li", "dl", "dt", "dd", "table", "thead", "tbody", "tfoot", "tr",
        "h1", "h2", "h3", "h4", "h5", "h6", "hr", "address",
    }
)
# Tags whose *content* is not document text at all.
_HTML_DROP_CONTENT = frozenset({"script", "style", "template", "noscript", "iframe", "object"})
