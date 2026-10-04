"""Colour: reading a ``\\color`` / ``\\textcolor`` specification and
applying the colour to OMML that has already been built."""

from __future__ import annotations

from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .builders import _el
from .errors import UnsupportedLatexError
from .tables import _COLORS
from .tokenizer import _read_raw_bracket, _read_raw_group

# Structures whose own glyphs (fraction bar, radical sign, fences, ...) take
# their formatting from the <m:ctrlPr> at the end of their properties.
_CONTROL_PROPERTIES = {
    "acc": "accPr", "bar": "barPr", "borderBox": "borderBoxPr",
    "box": "boxPr", "d": "dPr", "eqArr": "eqArrPr", "f": "fPr",
    "func": "funcPr", "groupChr": "groupChrPr", "limLow": "limLowPr",
    "limUpp": "limUppPr", "m": "mPr", "nary": "naryPr", "phant": "phantPr",
    "rad": "radPr", "sPre": "sPrePr", "sSub": "sSubPr",
    "sSubSup": "sSubSupPr", "sSup": "sSupPr",
}
_CONTROL_TAGS = {qn(f"m:{tag}"): properties
                 for tag, properties in _CONTROL_PROPERTIES.items()}


def _color_properties(color: str) -> Any:
    properties = OxmlElement("w:rPr")
    color_element = OxmlElement("w:color")
    color_element.set(qn("w:val"), color)
    properties.append(color_element)
    return properties


def _colorize(elements: list[Any], color: str) -> None:
    r"""Colour every glyph in `elements` that has no colour of its own yet.

    Runs get ``<w:color>`` in their Word formatting; structures -- whose
    bar, radical sign or fences are not runs -- get it through the
    ``<m:ctrlPr>`` that closes their properties.  Anything already coloured
    was coloured by a nested ``\color``, which wins, so it is left alone.
    """
    for root in elements:
        for node in list(root.iter()):
            if node.tag == qn("m:r"):
                properties = node.find(qn("w:rPr"))
                if properties is None:
                    position = 1 if len(node) and node[0].tag == qn("m:rPr") else 0
                    node.insert(position, _color_properties(color))
                elif properties.find(qn("w:color")) is None:
                    color_element = OxmlElement("w:color")
                    color_element.set(qn("w:val"), color)
                    properties.append(color_element)
                continue
            tag = _CONTROL_TAGS.get(node.tag)
            if tag is None:
                continue
            properties = node.find(qn(f"m:{tag}"))
            if properties is None:
                properties = _el(tag)
                node.insert(0, properties)
            if properties.find(qn("m:ctrlPr")) is None:
                control = _el("ctrlPr")
                control.append(_color_properties(color))
                properties.append(control)


def _read_color(tokens: list, index: int, owner: str) -> tuple:
    r"""Read a colour argument -- ``{red}``, ``{#1E90FF}``,
    ``[HTML]{1E90FF}``, ``[rgb]{1,0,0}``, ``[RGB]{255,0,0}`` or
    ``[gray]{0.5}``.  Returns (RRGGBB, index)."""
    model, index = _read_raw_bracket(tokens, index)
    specification, index = _read_raw_group(tokens, index, owner)
    specification = specification.strip()
    if model is None:
        if specification in _COLORS:
            return _COLORS[specification], index
        digits = specification[1:] if specification.startswith("#") else ""
        if len(digits) == 3:
            digits = "".join(digit * 2 for digit in digits)
        if len(digits) == 6 and all(d in "0123456789abcdefABCDEF" for d in digits):
            return digits.upper(), index
        raise UnsupportedLatexError(
            f"Unknown colour in \\{owner}: {specification!r} (named "
            f"colours: {', '.join(sorted(_COLORS))}; or use #RRGGBB or "
            "[HTML]{RRGGBB})"
        )
    model = model.strip()
    try:
        if model == "HTML" and len(specification) == 6:
            int(specification, 16)
            return specification.upper(), index
        if model in ("rgb", "RGB"):
            parts = [part.strip() for part in specification.split(",")]
            if len(parts) == 3:
                scale = 255.0 if model == "RGB" else 1.0
                values = [float(part) / scale for part in parts]
                if all(0.0 <= part <= 1.0 for part in values):
                    return "".join(
                        f"{round(part * 255):02X}" for part in values), index
        if model == "gray":
            level = float(specification)
            if 0.0 <= level <= 1.0:
                return f"{round(level * 255):02X}" * 3, index
    except ValueError:
        pass
    raise UnsupportedLatexError(
        f"Colour is not supported in \\{owner}: [{model}]{{{specification}}}"
    )
