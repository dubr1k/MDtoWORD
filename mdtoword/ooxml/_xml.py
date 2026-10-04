"""Schema-ordered element insertion and small XML utilities shared by the package.

``w:pPr``, ``w:rPr``, ``w:trPr``, ``w:settings`` and friends are ordered
sequences in the ECMA-376 schema, and Word reports a file as corrupt when
their children are out of order. Every helper of the package inserts
children through :func:`_insert_ordered` with one of the sequences below.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from typing import Any

from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.styles.style import BaseStyle
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def _tags(*nsptags: str) -> tuple[str, ...]:
    """Clark-notation tag names for prefixed tags such as ``"w:pStyle"``."""
    return tuple(qn(tag) for tag in nsptags)


def _w(*local_names: str) -> tuple[str, ...]:
    return _tags(*(f"w:{name}" for name in local_names))


# CT_PPr (paragraph properties), also valid for a style's w:pPr.
_PPR_SEQUENCE = _w(
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
)
# CT_RPr (run properties).
_RPR_SEQUENCE = _w(
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
    "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
    "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
)
# CT_TrPr (table row properties).
_TRPR_SEQUENCE = _w(
    "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
    "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del",
    "trPrChange",
)
# CT_Settings (settings.xml root).
_SETTINGS_SEQUENCE = _tags(
    "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation",
    "w:removeDateAndTime", "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape",
    "w:printPostScriptOverText", "w:printFractionalCharacterWidth", "w:printFormsData",
    "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts",
    "w:saveFormsData", "w:mirrorMargins", "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader", "w:bordersDoNotSurroundFooter", "w:gutterAtTop",
    "w:hideSpellingErrors", "w:hideGrammaticalErrors", "w:activeWritingStyle",
    "w:proofState", "w:formsDesign", "w:attachedTemplate", "w:linkStyles",
    "w:stylePaneFormatFilter", "w:stylePaneSortMethod", "w:documentType",
    "w:mailMerge", "w:revisionView", "w:trackRevisions", "w:doNotTrackMoves",
    "w:doNotTrackFormatting", "w:documentProtection", "w:autoFormatOverride",
    "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop", "w:autoHyphenation",
    "w:consecutiveHyphenLimit", "w:hyphenationZone", "w:doNotHyphenateCaps",
    "w:showEnvelope", "w:summaryLength", "w:clickAndTypeStyle",
    "w:defaultTableStyle", "w:evenAndOddHeaders", "w:bookFoldRevPrinting",
    "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", "w:noPunctuationKerning",
    "w:characterSpacingControl", "w:printTwoOnOne", "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter", "w:noLineBreaksBefore", "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema", "w:saveInvalidXml", "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly", "w:useXSLTWhenSaving", "w:saveThroughXslt",
    "w:showXMLTags", "w:alwaysMergeEmptyNamespace", "w:updateFields",
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang",
    "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
    "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol",
    "w:listSeparator",
)
_STYLES_SEQUENCE = _w("docDefaults", "latentStyles", "style")
_DOC_DEFAULTS_SEQUENCE = _w("rPrDefault", "pPrDefault")
_PBDR_SEQUENCE = _w("top", "left", "bottom", "right", "between", "bar")


def _insert_ordered(parent: Any, child: Any, sequence: Sequence[str]) -> Any:
    """Insert ``child`` into ``parent`` at the position ``sequence`` dictates.

    ``sequence`` lists the schema's child tags (Clark notation) in order.
    ``child`` goes before the first existing successor; with no successor it
    goes right after the last known predecessor, which keeps it ahead of
    trailing extension elements such as ``w14:docId`` in settings.xml.
    """
    position = sequence.index(child.tag)
    successors = frozenset(sequence[position + 1 :])
    predecessors = frozenset(sequence[: position + 1])
    last_predecessor = None
    for existing in parent.iterchildren():
        if existing.tag in successors:
            existing.addprevious(child)
            return child
        if existing.tag in predecessors:
            last_predecessor = existing
    if last_predecessor is not None:
        last_predecessor.addnext(child)
    else:
        parent.insert(0, child)
    return child


def _replace_ordered(parent: Any, child: Any, sequence: Sequence[str]) -> Any:
    """Remove any existing ``child.tag`` children, then insert ``child`` in order."""
    for existing in parent.findall(child.tag):
        parent.remove(existing)
    return _insert_ordered(parent, child, sequence)


def _get_or_insert(parent: Any, nsptag: str, sequence: Sequence[str]) -> Any:
    """Return the ``nsptag`` child of ``parent``, creating it in schema order."""
    existing = parent.find(qn(nsptag))
    if existing is not None:
        return existing
    return _insert_ordered(parent, OxmlElement(nsptag), sequence)


def _ppr_of(target: Any) -> Any:
    """``w:pPr`` of a Paragraph, a paragraph style, or a raw w:p / w:style."""
    if isinstance(target, Paragraph):
        return target._p.get_or_add_pPr()
    if isinstance(target, BaseStyle):
        return target.element.get_or_add_pPr()
    if hasattr(target, "get_or_add_pPr"):
        return target.get_or_add_pPr()
    raise TypeError(f"expected a Paragraph or a paragraph style, got {type(target)!r}")


def _rpr_of(target: Any) -> Any:
    """``w:rPr`` of a Run, a character/paragraph style, or a raw w:r / w:style."""
    if isinstance(target, Run):
        return target._r.get_or_add_rPr()
    if isinstance(target, BaseStyle):
        return target.element.get_or_add_rPr()
    if hasattr(target, "get_or_add_rPr"):
        return target.get_or_add_rPr()
    raise TypeError(f"expected a Run or a style, got {type(target)!r}")


_HEX_COLOR = re.compile(r"^#?([0-9A-Fa-f]{6})$")


def _hex_color(value: str) -> str:
    """Normalize ``"#a0a0a0"``/``"A0A0A0"``/``"auto"`` to Word's ``ST_HexColor``."""
    if value.lower() == "auto":
        return "auto"
    match = _HEX_COLOR.match(value.strip())
    if match is None:
        raise ValueError(f"expected a 6-digit hex color, got {value!r}")
    return match.group(1).upper()


def _is_element(node: Any) -> bool:
    """True for real elements (comments and processing instructions excluded)."""
    return isinstance(node.tag, str)


def _xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")
    )


def _free_partname(package: Any, preferred: str, template: str) -> PackURI:
    taken = {str(part.partname) for part in package.iter_parts()}
    if preferred not in taken:
        return PackURI(preferred)
    return package.next_partname(template)


def _max_int_attribute(elements: Iterable[str], default: int) -> int:
    values = [int(value) for value in elements if value.lstrip("-").isdigit()]
    return max(values) if values else default
