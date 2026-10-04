"""Pure helpers and the small records the renderer keeps.

Module-level functions that need no renderer state -- equation numbers,
``\\tag`` splitting, style fonts, remote/SVG image targets, HTML attributes,
the plain text of inline tokens -- plus ``_HtmlImage`` (an ``<img>`` tag
dressed as an image token), ``_ListState`` (one open list) and ``_Cell``
(one buffered table cell).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
import html
from typing import Any

from docx.oxml.ns import qn
from docx.styles.style import ParagraphStyle

from .. import latex_omml
from .constants import _HTML_ATTR, _STARRED_TAG, _THEME_FONT_ATTRS, _XML_INVALID

def _equation_number(label: str | None) -> str | None:
    """``1`` -> ``(1)``: how a ``$$ … $$ (1)`` label is displayed."""
    if label is None or not label.strip():
        return None
    return f"({label.strip()})"


def _split_tag(latex: str) -> tuple[str, str | None]:
    r"""Take ``\tag{..}`` (and ``\label``/``\nonumber``) out of a display formula.

    Returns the formula without them and the number as it is displayed:
    ``\tag{3}`` gives ``(3)``, the starred ``\tag*{A}`` gives a bare ``A``.
    """
    split = getattr(latex_omml, "split_equation_tag", None)
    if split is None or not latex.strip():
        return latex.strip(), None
    starred = _STARRED_TAG.search(latex) is not None
    remaining, tag = split(latex)
    if tag is None:
        return remaining.strip(), None
    if starred:
        return remaining.strip(), tag.strip() or None
    return remaining.strip(), _equation_number(tag)


def _plain_formatting() -> dict[str, bool]:
    return {
        "bold": False, "italic": False, "strike": False, "code": False,
        "sub": False, "sup": False, "mark": False, "underline": False,
    }


def _set_style_font(style: ParagraphStyle, font_name: str) -> None:
    """Set a style's font and clear any theme attributes overriding it.

    Word's built-in style definitions -- most visibly ``Heading 1``-``9``,
    whose ``w:rFonts`` point at the document theme's
    ``majorHAnsi``/``majorEastAsia``/``majorBidi`` fonts -- carry ``*Theme``
    attributes alongside the explicit ``ascii``/``hAnsi`` pair that
    ``style.font.name = ...`` writes. In OOXML a ``*Theme`` attribute takes
    precedence over its explicit sibling, so Word keeps rendering the style
    in the theme's font regardless of what was just set; python-docx's
    readback of ``style.font.name`` hides this because it only ever reports
    the explicit value it wrote, never checking whether a theme attribute
    overrides it. This mirrors the ``w:themeColor`` fix already applied to
    heading colours, except the colour setter clears the whole ``<w:color>``
    element while the font-name setter leaves ``w:rFonts`` otherwise
    untouched -- so the theme attributes have to be stripped here by hand.
    The lowercase ``cstheme`` spelling is what Word's own template actually
    writes; ``csTheme`` is stripped too since other OOXML producers vary.

    ``eastAsia`` and ``cs`` are set explicitly too, so text Word would
    otherwise route to the east-Asian or complex-script slot -- which can
    include some Cyrillic runs -- also honours the chosen font instead of
    falling back to whatever those slots would otherwise resolve to.
    """
    style.font.name = font_name
    run_properties = style.element.get_or_add_rPr()
    fonts = run_properties.get_or_add_rFonts()
    for attribute in _THEME_FONT_ATTRS:
        key = qn(f"w:{attribute}")
        if fonts.get(key) is not None:
            del fonts.attrib[key]
    fonts.set(qn("w:eastAsia"), font_name)
    fonts.set(qn("w:cs"), font_name)


def _is_remote_target(target: str) -> bool:
    """Does this image target reach off the local filesystem?

    Beyond the obvious ``http(s)`` URLs this covers UNC and protocol-relative
    paths: on Windows ``//host/share/x.png`` is a UNC reference, and merely
    asking ``Path.is_file()`` about it makes the SMB redirector connect out
    and authenticate -- an outbound request and an NTLM credential leak from
    what looks like a plain filesystem check.
    """
    lowered = target.lower()
    return lowered.startswith(("http://", "https://")) or target.startswith(("//", "\\\\"))


def _looks_like_svg(target: str, data: bytes) -> bool:
    if target.lower().split("?", 1)[0].endswith(".svg"):
        return True
    head = data[:2048].lstrip().lower()
    return head.startswith(b"<svg") or (head.startswith(b"<?xml") and b"<svg" in head)


def _html_attributes(raw: str) -> dict[str, str]:
    attributes: dict[str, str] = {}
    for name, value in _HTML_ATTR.findall(raw):
        if value[:1] in {'"', "'"}:
            value = value[1:-1]
        attributes[name.lower()] = _XML_INVALID.sub("", html.unescape(value))
    return attributes


def _inline_plain_text(children: Sequence[Any]) -> str:
    """Flatten inline tokens to the text a reader sees (no markup)."""
    parts: list[str] = []
    for token in children:
        if token.type in {"text", "code_inline"}:
            parts.append(token.content)
        elif token.type in {"softbreak", "hardbreak"}:
            parts.append(" ")
        elif token.type in {"math_inline", "math_inline_double"}:
            parts.append(token.content)
        elif token.type == "image":
            parts.append(token.content)
    return "".join(parts)


class _HtmlImage:
    """Adapter giving an ``<img>`` tag the interface of an image token."""

    def __init__(self, attributes: dict[str, str]) -> None:
        self._attributes = attributes
        self.content = attributes.get("alt", "")

    def attrGet(self, name: str) -> str | None:
        return self._attributes.get(name)


@dataclass
class _ListState:
    ordered: bool
    num_id: int
    level: int
    has_paragraph: bool = False


@dataclass
class _Cell:
    children: list[Any]
    content: str
    alignment: str | None
    header: bool
    line: int | None
