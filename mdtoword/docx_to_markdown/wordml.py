"""WordprocessingML names and the element helpers every module shares.

Qualified tag and attribute names (``W_P`` is ``{w-namespace}p``), OOXML
value parsing, the visible text, runs and non-text content of an element,
and access to the package parts (styles, numbering, notes) related to the
main document.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from docx.oxml import parse_xml

# --------------------------------------------------------------------------
# XML names
# --------------------------------------------------------------------------

_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}


def _q(name: str) -> str:
    prefix, local = name.split(":")
    return f"{{{_NS[prefix]}}}{local}"


W_P, W_R, W_T, W_TBL = _q("w:p"), _q("w:r"), _q("w:t"), _q("w:tbl")
W_PPR, W_RPR, W_PSTYLE, W_RSTYLE = _q("w:pPr"), _q("w:rPr"), _q("w:pStyle"), _q("w:rStyle")
W_VAL, W_TYPE, W_ID, W_NAME = _q("w:val"), _q("w:type"), _q("w:id"), _q("w:name")
W_TR, W_TC, W_TCPR, W_TRPR = _q("w:tr"), _q("w:tc"), _q("w:tcPr"), _q("w:trPr")
W_GRIDSPAN, W_VMERGE, W_HMERGE = _q("w:gridSpan"), _q("w:vMerge"), _q("w:hMerge")
W_GRIDBEFORE, W_GRIDAFTER = _q("w:gridBefore"), _q("w:gridAfter")
W_HYPERLINK, W_ANCHOR, W_TOOLTIP = _q("w:hyperlink"), _q("w:anchor"), _q("w:tooltip")
W_FLDSIMPLE, W_INSTR, W_FLDCHAR = _q("w:fldSimple"), _q("w:instr"), _q("w:fldChar")
W_FLDCHARTYPE, W_INSTRTEXT = _q("w:fldCharType"), _q("w:instrText")
W_TAB, W_PTAB, W_BR, W_CR = _q("w:tab"), _q("w:ptab"), _q("w:br"), _q("w:cr")
W_NOBREAKHYPHEN, W_SYM, W_CHAR, W_FONT = (
    _q("w:noBreakHyphen"), _q("w:sym"), _q("w:char"), _q("w:font"))
W_DRAWING, W_PICT, W_OBJECT = _q("w:drawing"), _q("w:pict"), _q("w:object")
W_FOOTNOTEREFERENCE, W_ENDNOTEREFERENCE = _q("w:footnoteReference"), _q("w:endnoteReference")
W_FOOTNOTE, W_ENDNOTE = _q("w:footnote"), _q("w:endnote")
W_SDT, W_SDTCONTENT, W_SDTPR = _q("w:sdt"), _q("w:sdtContent"), _q("w:sdtPr")
W_DOCPARTOBJ, W_DOCPARTGALLERY = _q("w:docPartObj"), _q("w:docPartGallery")
W_INS, W_DEL, W_MOVETO, W_MOVEFROM = _q("w:ins"), _q("w:del"), _q("w:moveTo"), _q("w:moveFrom")
W_SMARTTAG, W_CUSTOMXML, W_DIR, W_BDO = (
    _q("w:smartTag"), _q("w:customXml"), _q("w:dir"), _q("w:bdo"))
W_RUBY, W_RUBYBASE = _q("w:ruby"), _q("w:rubyBase")
W_BOOKMARKSTART, W_TXBXCONTENT = _q("w:bookmarkStart"), _q("w:txbxContent")
W_B, W_I, W_STRIKE, W_DSTRIKE, W_U = _q("w:b"), _q("w:i"), _q("w:strike"), _q("w:dstrike"), _q("w:u")
W_HIGHLIGHT, W_VERTALIGN, W_VANISH, W_CAPS = (
    _q("w:highlight"), _q("w:vertAlign"), _q("w:vanish"), _q("w:caps"))
W_RFONTS, W_ASCII, W_HANSI = _q("w:rFonts"), _q("w:ascii"), _q("w:hAnsi")
W_NUMPR, W_NUMID, W_ILVL = _q("w:numPr"), _q("w:numId"), _q("w:ilvl")
W_OUTLINELVL, W_IND, W_LEFT, W_START = _q("w:outlineLvl"), _q("w:ind"), _q("w:left"), _q("w:start")
W_JC, W_PBDR, W_BOTTOM = _q("w:jc"), _q("w:pBdr"), _q("w:bottom")
W_STYLE, W_STYLEID, W_BASEDON, W_DEFAULT = (
    _q("w:style"), _q("w:styleId"), _q("w:basedOn"), _q("w:default"))
W_ABSTRACTNUM, W_ABSTRACTNUMID, W_NUM, W_LVL = (
    _q("w:abstractNum"), _q("w:abstractNumId"), _q("w:num"), _q("w:lvl"))
W_LVLOVERRIDE, W_STARTOVERRIDE, W_NUMFMT, W_LVLTEXT = (
    _q("w:lvlOverride"), _q("w:startOverride"), _q("w:numFmt"), _q("w:lvlText"))
W_NUMSTYLELINK = _q("w:numStyleLink")
M_OMATH, M_OMATHPARA = _q("m:oMath"), _q("m:oMathPara")
R_ID, R_EMBED, R_LINK = _q("r:id"), _q("r:embed"), _q("r:link")
WP_DOCPR, A_BLIP, ASVG_SVGBLIP = _q("wp:docPr"), _q("a:blip"), _q("asvg:svgBlip")
V_IMAGEDATA, V_RECT, O_HR, O_TITLE = _q("v:imagedata"), _q("v:rect"), _q("o:hr"), _q("o:title")
MC_ALTERNATECONTENT, MC_CHOICE, MC_FALLBACK = (
    _q("mc:AlternateContent"), _q("mc:Choice"), _q("mc:Fallback"))

_TRANSPARENT_INLINE = frozenset({W_INS, W_MOVETO, W_SMARTTAG, W_CUSTOMXML, W_DIR, W_BDO})
_SKIPPED_TEXT = frozenset({W_DEL, W_MOVEFROM})

# --------------------------------------------------------------------------
# Values and element content
# --------------------------------------------------------------------------


def _on(element: Any) -> bool:
    """OOXML on/off property: present and not explicitly false."""
    if element is None:
        return False
    value = element.get(W_VAL)
    return value is None or value.lower() not in ("0", "false", "off", "none")


def _int(value: str | None, default: int = 0) -> int:
    try:
        return int(value) if value is not None else default
    except ValueError:
        return default


def _plain_text(element: Any) -> str:
    """The visible text of a paragraph, for heading slugs."""
    parts: list[str] = []
    for child in element:
        tag = child.tag
        if tag in _SKIPPED_TEXT or tag in (M_OMATH, M_OMATHPARA):
            continue
        if tag == W_T:
            parts.append(child.text or "")
        elif tag in (W_TAB, W_PTAB):
            parts.append(" ")
        elif isinstance(tag, str):
            parts.append(_plain_text(child))
    return "".join(parts)


def text_runs(element: Any) -> Iterable[Any]:
    """The runs of an element in order, deleted text and paragraph properties skipped."""
    for child in element:
        tag = child.tag
        if tag == W_R:
            yield child
        elif tag in _SKIPPED_TEXT or tag == W_PPR or not isinstance(tag, str):
            continue
        else:
            yield from text_runs(child)


def has_content(paragraph: Any) -> bool:
    """Pictures or equations: non-text content that makes a paragraph real."""
    for tag in (W_DRAWING, M_OMATH, W_OBJECT):
        if next(paragraph.iter(tag), None) is not None:
            return True
    for pict in paragraph.iter(W_PICT):
        if next(pict.iter(V_IMAGEDATA), None) is not None:
            return True
    return False


# --------------------------------------------------------------------------
# Package parts
# --------------------------------------------------------------------------


def related_part(part: Any, reltype: str) -> Any:
    for relationship in part.rels.values():
        if relationship.reltype == reltype and not relationship.is_external:
            return relationship.target_part
    return None


def element_of(part: Any) -> Any:
    element = getattr(part, "element", None)
    if element is not None:
        return element
    return parse_xml(part.blob)


def part_element(main_part: Any, reltype: str) -> Any:
    """The root element of the part related to ``main_part`` by ``reltype``, if any."""
    part = related_part(main_part, reltype)
    return element_of(part) if part is not None else None
