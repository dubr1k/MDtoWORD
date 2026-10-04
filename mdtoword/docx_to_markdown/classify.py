"""Paragraph classification and run formatting.

What a paragraph *is* -- heading, list item, quote, call-out, code line,
thematic break -- and how a run is formatted are both decided from the
direct properties first, then the style chain. Paragraph results are cached
for the whole document.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .code import _CODE_CHAR_STYLE, _is_monospace, is_code_paragraph
from .numbering import _Level, _Numbering
from .styles import _Style, _Styles
from .wordml import (
    O_HR,
    V_RECT,
    W_ASCII,
    W_B,
    W_BOTTOM,
    W_CAPS,
    W_DSTRIKE,
    W_HANSI,
    W_HIGHLIGHT,
    W_I,
    W_ILVL,
    W_IND,
    W_JC,
    W_LEFT,
    W_NUMID,
    W_NUMPR,
    W_OUTLINELVL,
    W_PBDR,
    W_PPR,
    W_PSTYLE,
    W_RFONTS,
    W_RPR,
    W_RSTYLE,
    W_START,
    W_STRIKE,
    W_U,
    W_VAL,
    W_VANISH,
    W_VERTALIGN,
    _int,
    _on,
    _plain_text,
    _q,
    has_content,
)

_HEADING_STYLE = re.compile(r"heading\s*([1-9])")
_QUOTE_STYLES = {"quote": 1, "intense quote": 1, "block text": 1}
_QUOTE_LEVEL_STYLE = re.compile(r"quote\s*([1-9])")


@dataclass
class _ParaInfo:
    chain: list[_Style]
    heading: int = 0
    title: bool = False
    subtitle: bool = False
    level: _Level | None = None
    quote: int = 0
    code: bool = False
    rule: bool = False
    indent: int = 0
    centered: bool = False
    callout: bool = False
    quote_bar: bool = False
    toc_heading: bool = False
    direct_indent: int | None = None


@dataclass
class _RunFormat:
    attrs: frozenset[str]
    code: bool
    hidden: bool
    caps: bool


class _Classifier:
    """Classify paragraphs and resolve run formatting for one document."""

    def __init__(self, styles: _Styles, numbering: _Numbering) -> None:
        self.styles = styles
        self.numbering = numbering
        # Keyed by the element itself, not id(): holding the proxy keeps lxml
        # from recycling it, so an id can never alias another paragraph.
        self._info_cache: dict[Any, _ParaInfo] = {}

    # -- paragraph classification ----------------------------------------------

    def paragraph_info(self, paragraph: Any) -> _ParaInfo:
        cached = self._info_cache.get(paragraph)
        if cached is not None:
            return cached
        ppr = paragraph.find(W_PPR)
        style_element = ppr.find(W_PSTYLE) if ppr is not None else None
        style_id = (style_element.get(W_VAL) if style_element is not None
                    else self.styles.default_paragraph)
        chain = self.styles.chain(style_id)
        info = _ParaInfo(chain=chain)
        name = chain[0].name if chain else ""

        def inherited(tag: str) -> Any:
            if ppr is not None and ppr.find(tag) is not None:
                return ppr.find(tag)
            for style in chain:
                if style.ppr is not None and style.ppr.find(tag) is not None:
                    return style.ppr.find(tag)
            return None

        heading = _HEADING_STYLE.fullmatch(name)
        if heading:
            info.heading = int(heading.group(1))
        elif name == "title":
            info.heading, info.title = 1, True
        elif name == "subtitle":
            info.subtitle = True
        elif name == "toc heading":
            info.toc_heading = True
        elif not name.startswith("toc"):
            outline = inherited(W_OUTLINELVL)
            if outline is not None and 0 <= _int(outline.get(W_VAL), 9) <= 8:
                info.heading = _int(outline.get(W_VAL)) + 1

        quote = _QUOTE_LEVEL_STYLE.fullmatch(name)
        if name in _QUOTE_STYLES:
            info.quote = _QUOTE_STYLES[name]
        elif quote:
            info.quote = int(quote.group(1))

        info.level = self._list_level(ppr, chain)
        if info.level is not None and info.subtitle:
            info.level = None

        indent_element = ppr.find(W_IND) if ppr is not None else None
        indent_value = None
        if indent_element is not None:
            indent_value = indent_element.get(W_LEFT) or indent_element.get(W_START)
        if indent_value is not None:
            info.indent = info.direct_indent = _int(indent_value)
        elif info.level is not None and info.level.indent is not None:
            info.indent = info.level.indent
        else:
            for style in chain:
                ind = style.ppr.find(W_IND) if style.ppr is not None else None
                value = (ind.get(W_LEFT) or ind.get(W_START)) if ind is not None else None
                if value is not None:
                    info.indent = _int(value)
                    break

        justification = inherited(W_JC)
        info.centered = justification is not None and justification.get(W_VAL) == "center"

        borders = ppr.find(W_PBDR) if ppr is not None else None
        bottom = borders.find(W_BOTTOM) if borders is not None else None
        # A shaded paragraph with a bar on its left is a call-out box -- how
        # GitHub-style alerts ("> [!NOTE]") look in Word.
        shading = ppr.find(_q("w:shd")) if ppr is not None else None
        left = borders.find(_q("w:left")) if borders is not None else None
        left_bar = left is not None and left.get(W_VAL) not in ("nil", "none")
        info.callout = (
            left_bar and shading is not None
            and (shading.get(_q("w:fill")) or "auto").lower() not in ("auto", "ffffff")
        )
        # A bare bar on the left marks a block quote whatever the paragraph's
        # style: a list item or a code line inside "> ..." keeps its own.
        info.quote_bar = left_bar and not info.callout
        has_rule = bottom is not None and bottom.get(W_VAL) not in ("nil", "none")
        if not has_rule:
            has_rule = any(rect.get(O_HR) in ("t", "true") for rect in paragraph.iter(V_RECT))
        if has_rule and not _plain_text(paragraph).strip() and not has_content(paragraph):
            info.rule = True

        if not (info.heading or info.level is not None or info.quote or info.title
                or info.subtitle or info.callout or info.toc_heading):
            info.code = is_code_paragraph(self, paragraph, chain, style_id)
        self._info_cache[paragraph] = info
        return info

    def _list_level(self, ppr: Any, chain: list[_Style]) -> _Level | None:
        num_id = ilvl = None
        num_pr = ppr.find(W_NUMPR) if ppr is not None else None
        if num_pr is not None:
            num_element, level_element = num_pr.find(W_NUMID), num_pr.find(W_ILVL)
            num_id = num_element.get(W_VAL) if num_element is not None else None
            ilvl = level_element.get(W_VAL) if level_element is not None else None
        for style in chain:
            if num_id is not None and ilvl is not None:
                break
            style_num = style.ppr.find(W_NUMPR) if style.ppr is not None else None
            if style_num is None:
                continue
            if num_id is None and style_num.find(W_NUMID) is not None:
                num_id = style_num.find(W_NUMID).get(W_VAL)
            if ilvl is None and style_num.find(W_ILVL) is not None:
                ilvl = style_num.find(W_ILVL).get(W_VAL)
        if not num_id or num_id == "0":
            return None
        level = self.numbering.level(num_id, _int(ilvl))
        if level is None or level.number_format == "none":
            return None
        return level

    # -- run formatting ----------------------------------------------------------

    def run_format(self, run: Any, *, in_link: bool) -> _RunFormat:
        rpr = run.find(W_RPR)
        style_element = rpr.find(W_RSTYLE) if rpr is not None else None
        chain = self.styles.chain(style_element.get(W_VAL) if style_element is not None else None)
        sources = ([rpr] if rpr is not None else []) + [s.rpr for s in chain if s.rpr is not None]

        def prop(tag: str) -> Any:
            for source in sources:
                found = source.find(tag)
                if found is not None:
                    return found
            return None

        attrs: set[str] = set()
        if _on(prop(W_B)):
            attrs.add("bold")
        if _on(prop(W_I)):
            attrs.add("italic")
        if _on(prop(W_STRIKE)) or _on(prop(W_DSTRIKE)):
            attrs.add("strike")
        hyperlink_style = any(style.name == "hyperlink" for style in chain)
        underline = prop(W_U)
        if (underline is not None and underline.get(W_VAL, "single") != "none"
                and not in_link and not hyperlink_style):
            attrs.add("underline")
        highlight = prop(W_HIGHLIGHT)
        if highlight is not None and highlight.get(W_VAL) not in (None, "none"):
            attrs.add("highlight")
        vertical = prop(W_VERTALIGN)
        if vertical is not None:
            if vertical.get(W_VAL) == "superscript":
                attrs.add("sup")
            elif vertical.get(W_VAL) == "subscript":
                attrs.add("sub")
        font = None
        for source in sources:
            fonts = source.find(W_RFONTS)
            if fonts is not None and (fonts.get(W_ASCII) or fonts.get(W_HANSI)):
                font = fonts.get(W_ASCII) or fonts.get(W_HANSI)
                break
        code = _is_monospace(font) or any(_CODE_CHAR_STYLE.search(s.name) for s in chain)
        return _RunFormat(frozenset(attrs), code, _on(prop(W_VANISH)), _on(prop(W_CAPS)))
