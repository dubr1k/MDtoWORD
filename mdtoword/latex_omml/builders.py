"""OMML element constructors.

Build every element through these so the shapes stay consistent across
this package and its callers.
"""

from __future__ import annotations

from typing import Any, Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .style import _Style
from .tables import _INTEGRAL_CHARACTERS


def _el(tag: str) -> Any:
    return OxmlElement(f"m:{tag}")


def _on(tag: str, value: str = "1") -> Any:
    """A properties child holding one `m:val`, e.g. <m:degHide m:val="1"/>."""
    element = _el(tag)
    element.set(qn("m:val"), value)
    return element


def _run(text: str, *, italic: bool = False, upright: bool = False,
         bold: bool = False, plain: bool = False,
         script: Optional[str] = None) -> Any:
    """One <m:r>. Identifiers are italic, numbers and operators are not.

    `upright` makes the run Word "normal text" (`<m:nor>`), the shape of
    ``\\text{...}``.  `script` sets a math alphabet (`<m:scr>`), and
    `plain` asks for an explicit upright math style (`<m:sty m:val="p">`)
    -- upright *math*, as in ``\\mathrm{d}``, rather than normal text.
    """
    run = _el("r")
    if italic or upright or bold or plain or script is not None:
        properties = _el("rPr")
        # OOXML's CT_RPR makes <m:nor> and <m:sty> a *choice*: a run may
        # carry one or the other, never both. `\mathbf` is upright and bold
        # at once, so the two have to be reconciled rather than stacked --
        # emitting both is rejected by the ISO/IEC 29500-4 schema. <m:sty>
        # already encodes uprightness ("b" is bold upright, "bi" is bold
        # italic), so whenever a style value applies it says everything
        # <m:nor> would have, and <m:nor> is left for the unbolded literal
        # text of \text{...} and friends.
        if script is not None or plain:
            # A math alphabet always states its style too: without one,
            # Word would italicise the letters, and most alphabets
            # (double-struck, script, fraktur) have no italic form at all.
            if script is not None:
                properties.append(_on("scr", script))
            if bold:
                value = "bi" if italic else "b"
            else:
                value = "i" if italic else "p"
            properties.append(_on("sty", value))
        elif bold:
            properties.append(_on("sty", "bi" if italic and not upright else "b"))
        elif upright:
            properties.append(_on("nor"))
        elif italic:
            properties.append(_on("sty", "i"))
        run.append(properties)
    text_element = _el("t")
    # Only declare xml:space when it changes anything: an unconditional
    # attribute would bloat every single run in the document.
    if text != text.strip():
        text_element.set(qn("xml:space"), "preserve")
    text_element.text = text
    run.append(text_element)
    return run


def _styled_run(text: str, style: _Style, kind: str) -> Any:
    """A run for `text` written in the alphabet `style` selects.

    `kind` is "letter" for a Latin letter (italic by default), "number" for
    digits, punctuation and other plain characters (upright by default),
    and "symbol" for the character a command such as ``\\alpha`` or
    ``\\le`` stands for (left to Word's own defaults).
    """
    bold = style.bold
    if style.script is not None:
        return _run(text, script=style.script, bold=bold)
    if style.face == "upright":
        # Every run, not only Latin letters: a Cyrillic or Greek letter
        # typed directly inside `\mathrm{...}` must not be italicised.
        return _run(text, plain=True)
    if kind == "letter":
        if style.face == "bold":
            return _run(text, upright=True, bold=True)
        return _run(text, italic=True, bold=bold)
    if kind == "number":
        if style.face == "italic":
            return _run(text, italic=True)
        return _run(text, bold=bold)
    return _run(text, italic=style.face == "bolditalic", bold=bold)


def _word_format(run: Any, *, bold: bool = False, italic: bool = False) -> Any:
    """Add Word run formatting (`<w:rPr>`) to an <m:r>.

    It sits after the run's <m:rPr> and before its <m:t>, as CT_R orders
    them; <w:b> comes before <w:i>, as CT_RPr orders those.
    """
    properties = OxmlElement("w:rPr")
    if bold:
        properties.append(OxmlElement("w:b"))
    if italic:
        properties.append(OxmlElement("w:i"))
    position = 1 if len(run) and run[0].tag == qn("m:rPr") else 0
    run.insert(position, properties)
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
        properties.append(_on("type", "noBar"))
        fraction.append(properties)
    fraction.append(_wrap("num", numerator))
    fraction.append(_wrap("den", denominator))
    return fraction


def _radical(degree: list[Any] | None, radicand: list[Any]) -> Any:
    radical = _el("rad")
    properties = _el("radPr")
    if degree is None:
        properties.append(_on("degHide"))
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
    properties.append(_on(name, value))
    return properties


def _nary(character: str, sub: list[Any] | None, sup: list[Any] | None,
          body: list[Any], limit_location: Optional[str] = None) -> Any:
    """<m:nary>: a big operator with its limits and its operand.

    `limit_location` overrides where the limits go -- "undOvr" (stacked,
    `\\limits`) or "subSup" (beside, `\\nolimits`); by default integrals
    keep them beside the sign and every other operator stacks them.
    """
    nary = _el("nary")
    properties = _property("naryPr", "chr", character)
    if limit_location is None:
        limit_location = (
            "subSup" if character in _INTEGRAL_CHARACTERS else "undOvr")
    properties.append(_on("limLoc", limit_location))
    # Without these Word draws an empty placeholder box where the missing
    # limit would go.
    for missing, tag in ((sub is None, "subHide"), (sup is None, "supHide")):
        if missing:
            properties.append(_on(tag))
    nary.append(properties)
    nary.append(_wrap("sub", sub or []))
    nary.append(_wrap("sup", sup or []))
    nary.append(_wrap("e", body))
    return nary


def _delimiter(begin: str, end: str, *segments: list[Any],
               separator: Optional[str] = None) -> Any:
    """<m:d>: a fenced group. An empty `begin`/`end` means "no fence".

    Several `segments` are split by `separator` -- the ``\\middle``
    character, which Word stretches together with the fences.
    """
    delimiter = _el("d")
    properties = _el("dPr")
    properties.append(_on("begChr", begin))
    if separator is not None:
        properties.append(_on("sepChr", separator))
    properties.append(_on("endChr", end))
    delimiter.append(properties)
    for segment in segments or ([],):
        delimiter.append(_wrap("e", segment))
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


def _limit_upp(base: list[Any], limit: list[Any]) -> Any:
    element = _el("limUpp")
    element.append(_wrap("e", base))
    element.append(_wrap("lim", limit))
    return element


def _limits(base: list[Any], sub: list[Any] | None,
            sup: list[Any] | None) -> list[Any]:
    """`base` with `sub` stacked under it and `sup` over it, as TeX's
    ``\\limits`` places them: <m:limLow> inside <m:limUpp>."""
    result = base
    if sub is not None:
        result = [_limit_low(result, sub)]
    if sup is not None:
        result = [_limit_upp(result, sup)]
    return result


def _group_character(base: list[Any], character: str, position: str,
                     vertical: str) -> Any:
    """<m:groupChr>: a stretchy `character` above (`top`) or below (`bot`)
    the base.  `vertical` (`m:vertJc`) says which part sits on the
    baseline: ``\\overbrace`` keeps its base there (`bot`), an extensible
    arrow keeps the arrow there."""
    group = _el("groupChr")
    properties = _el("groupChrPr")
    properties.append(_on("chr", character))
    properties.append(_on("pos", position))
    properties.append(_on("vertJc", vertical))
    group.append(properties)
    group.append(_wrap("e", base))
    return group


def _border_box(base: list[Any], *, hide: bool = False,
                strikes: tuple = ()) -> Any:
    """<m:borderBox>: a frame (``\\boxed``) or, with the frame hidden and
    diagonal strikes on, a crossed-out expression (``\\cancel``)."""
    box = _el("borderBox")
    if hide or strikes:
        properties = _el("borderBoxPr")
        # CT_BorderBoxPr's sequence order: the four hides, then the strikes.
        if hide:
            for tag in ("hideTop", "hideBot", "hideLeft", "hideRight"):
                properties.append(_on(tag))
        for tag in ("strikeBLTR", "strikeTLBR"):
            if tag in strikes:
                properties.append(_on(tag))
        box.append(properties)
    box.append(_wrap("e", base))
    return box


def _phantom(base: list[Any], *, show: bool, zero_width: bool = False,
             zero_ascent: bool = False, zero_descent: bool = False) -> Any:
    """<m:phant>: `base` taking up its space invisibly (``\\phantom``) or
    visibly with some dimensions zeroed (``\\smash``)."""
    phantom = _el("phant")
    properties = _el("phantPr")
    # CT_PhantPr's sequence order: show, zeroWid, zeroAsc, zeroDesc.
    if not show:
        properties.append(_on("show", "0"))
    for flag, tag in ((zero_width, "zeroWid"), (zero_ascent, "zeroAsc"),
                      (zero_descent, "zeroDesc")):
        if flag:
            properties.append(_on(tag))
    phantom.append(properties)
    phantom.append(_wrap("e", base))
    return phantom


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
            column_properties.append(_on("count", "1"))
            column_properties.append(_on("mcJc", justification))
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
    sequence, and `_run` only ever writes ``<m:nor>`` or ``<m:scr>`` /
    ``<m:sty>`` before it, so appending is always in schema order.

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
    r"""Build the element for ``\over``, ``\atop``, ``\choose``,
    ``\brace`` or ``\brack``.

    ``\choose`` is exactly ``\binom`` written infix, so both go through the
    same barless-fraction-in-parentheses shape; ``\brace`` and ``\brack``
    are the same stack in braces and brackets.
    """
    if name == "over":
        return _fraction(numerator, denominator)
    stack = _fraction(numerator, denominator, no_bar=True)
    if name == "atop":
        return stack
    begin, end = {"choose": ("(", ")"), "brace": ("{", "}"),
                  "brack": ("[", "]")}[name]
    return _delimiter(begin, end, [stack])
