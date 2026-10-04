"""Handlers for the composite OMML elements that map to one LaTeX construct each.

Fractions, radicals, scripts, n-ary operators, accents, bars, functions,
limits, group characters, boxes and phantoms. Every handler takes the
running :class:`~.converter._Converter` and the element and returns
``(latex, kind)``.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .names import _is_on, _m
from .symbols import _ACCENTS, _FUNCTIONS, _INTEGRALS, _MULTIWORD_FUNCTIONS, _NARY
from .text import _SIMPLE_ATOM, _braced_base, _join, _text_latex

if TYPE_CHECKING:
    from .converter import _Converter


def fraction(converter: _Converter, element: Any) -> tuple[str, str]:
    kind = converter.prop_value(element, "fPr", "type", "bar")
    numerator = converter.seq(element.find(_m("num")))
    denominator = converter.seq(element.find(_m("den")))
    if kind == "noBar":
        return f"{{{numerator} \\atop {denominator}}}", "atom"
    if kind in ("lin", "skw"):
        return f"{{{numerator}}}/{{{denominator}}}", "atom"
    return f"\\frac{{{numerator}}}{{{denominator}}}", "atom"


def radical(converter: _Converter, element: Any) -> tuple[str, str]:
    radicand = converter.seq(element.find(_m("e")))
    degree = ""
    if not _is_on(converter.prop(element, "radPr", "degHide")):
        degree = converter.seq(element.find(_m("deg")))
    if degree:
        return f"\\sqrt[{degree}]{{{radicand}}}", "atom"
    return f"\\sqrt{{{radicand}}}", "atom"


# -- scripts ----------------------------------------------------------------


def superscript(converter: _Converter, element: Any) -> tuple[str, str]:
    base = _braced_base(converter.seq(element.find(_m("e"))))
    return f"{base}^{{{converter.seq(element.find(_m('sup')))}}}", "atom"


def subscript(converter: _Converter, element: Any) -> tuple[str, str]:
    base = _braced_base(converter.seq(element.find(_m("e"))))
    return f"{base}_{{{converter.seq(element.find(_m('sub')))}}}", "atom"


def sub_superscript(converter: _Converter, element: Any) -> tuple[str, str]:
    base = _braced_base(converter.seq(element.find(_m("e"))))
    sub = converter.seq(element.find(_m("sub")))
    sup = converter.seq(element.find(_m("sup")))
    return f"{base}_{{{sub}}}^{{{sup}}}", "atom"


def pre_script(converter: _Converter, element: Any) -> tuple[str, str]:
    sub = converter.seq(element.find(_m("sub")))
    sup = converter.seq(element.find(_m("sup")))
    base = _braced_base(converter.seq(element.find(_m("e"))))
    return f"{{}}_{{{sub}}}^{{{sup}}}{base}", "atom"


# -- big operators ------------------------------------------------------------


def _limit(converter: _Converter, container: Any) -> str:
    """A limit of a big operator; a one-column stack becomes \\substack."""
    if container is None:
        return ""
    children = converter.children(container)
    if (len(children) == 1 and children[0].tag == _m("m")
            and children[0].find(_m("mPr")) is None):
        rows = children[0].findall(_m("mr"))
        if all(len(row.findall(_m("e"))) == 1 for row in rows):
            lines = [converter.seq(row.find(_m("e"))) for row in rows]
            return r"\substack{" + r" \\ ".join(lines) + "}"
    return converter.seq(container)


def nary(converter: _Converter, element: Any) -> tuple[str, str]:
    character = converter.prop_value(element, "naryPr", "chr", "∫") or "∫"
    command = _NARY.get(character)
    head = f"\\{command}" if command else r"\mathop{" + _text_latex(character) + "}"
    # Integrals put their limits beside the sign, every other operator
    # above and below; say so only when the document differs.
    location = converter.prop_value(element, "naryPr", "limLoc")
    integral = character in _INTEGRALS
    if location == "undOvr" and integral:
        head += r"\limits"
    elif location == "subSup" and not integral:
        head += r"\nolimits"
    if not _is_on(converter.prop(element, "naryPr", "subHide")):
        sub = _limit(converter, element.find(_m("sub")))
        if sub:
            head += f"_{{{sub}}}"
    if not _is_on(converter.prop(element, "naryPr", "supHide")):
        sup = _limit(converter, element.find(_m("sup")))
        if sup:
            head += f"^{{{sup}}}"
    body = converter.seq(element.find(_m("e")))
    return (f"{head} {body}" if body else head), "nary"


# -- accents, bars and functions ------------------------------------------------


def accent(converter: _Converter, element: Any) -> tuple[str, str]:
    character = converter.prop_value(element, "accPr", "chr", "̂") or "̂"
    base = converter.seq(element.find(_m("e")))
    command = _ACCENTS.get(character)
    if command:
        return f"\\{command}{{{base}}}", "atom"
    return f"\\overset{{{_text_latex(character)}}}{{{base}}}", "atom"


def bar(converter: _Converter, element: Any) -> tuple[str, str]:
    position = converter.prop_value(element, "barPr", "pos", "bot")
    base = converter.seq(element.find(_m("e")))
    command = "overline" if position == "top" else "underline"
    return f"\\{command}{{{base}}}", "atom"


def function(converter: _Converter, element: Any) -> tuple[str, str]:
    name_el = element.find(_m("fName"))
    previous = converter.in_function_name
    if name_el is not None and all(
        child.tag == _m("r") for child in converter.children(name_el)
    ):
        converter.in_function_name = True
    try:
        name = converter.seq(name_el)
    finally:
        converter.in_function_name = previous
    argument = converter.seq(element.find(_m("e")))
    if not argument:
        return name, "atom"
    if _SIMPLE_ATOM.fullmatch(argument) and argument[0].isalnum():
        return f"{name} {argument}", "atom"
    if _SIMPLE_ATOM.fullmatch(argument) or argument.startswith(r"\left"):
        return _join([name, argument]), "atom"
    return f"{name}{{{argument}}}", "atom"


# -- limits and group characters ------------------------------------------------


def _group_character(converter: _Converter, element: Any) -> tuple[str, str] | None:
    """(character, position) of a lone m:groupChr inside a container."""
    children = converter.children(element) if element is not None else []
    if len(children) == 1 and children[0].tag == _m("groupChr"):
        group = children[0]
        return (converter.prop_value(group, "groupChrPr", "chr", "⏟") or "⏟",
                converter.prop_value(group, "groupChrPr", "pos", "bot") or "bot")
    return None


def limit_lower(converter: _Converter, element: Any) -> tuple[str, str]:
    base_el = element.find(_m("e"))
    limit = converter.seq(element.find(_m("lim")))
    group = _group_character(converter, base_el)
    base = converter.seq(base_el)
    if group is not None and group[0] == "⏟":
        return f"{base}_{{{limit}}}", "atom"
    if re.fullmatch(r"\\[A-Za-z]+", base) and base[1:] in (
            _FUNCTIONS | set(_MULTIWORD_FUNCTIONS.values())):
        return f"{base}_{{{limit}}}", "atom"
    return f"\\underset{{{limit}}}{{{base}}}", "atom"


def limit_upper(converter: _Converter, element: Any) -> tuple[str, str]:
    base_el = element.find(_m("e"))
    limit = converter.seq(element.find(_m("lim")))
    group = _group_character(converter, base_el)
    base = converter.seq(base_el)
    if group is not None and group[0] == "⏞":
        return f"{base}^{{{limit}}}", "atom"
    return f"\\overset{{{limit}}}{{{base}}}", "atom"


def group_character(converter: _Converter, element: Any) -> tuple[str, str]:
    character = converter.prop_value(element, "groupChrPr", "chr", "⏟") or "⏟"
    position = converter.prop_value(element, "groupChrPr", "pos", "bot") or "bot"
    base = converter.seq(element.find(_m("e")))
    if character == "⏟":
        return f"\\underbrace{{{base}}}", "atom"
    if character == "⏞":
        return f"\\overbrace{{{base}}}", "atom"
    if character in ("→", "⟶"):
        command = "xrightarrow" if position == "bot" else "overrightarrow"
        return f"\\{command}{{{base}}}", "atom"
    if character in ("←", "⟵"):
        command = "xleftarrow" if position == "bot" else "overleftarrow"
        return f"\\{command}{{{base}}}", "atom"
    command = "overset" if position == "top" else "underset"
    return f"\\{command}{{{_text_latex(character)}}}{{{base}}}", "atom"


# -- boxes ----------------------------------------------------------------------


def border_box(converter: _Converter, element: Any) -> tuple[str, str]:
    return f"\\boxed{{{converter.seq(element.find(_m('e')))}}}", "atom"


def box(converter: _Converter, element: Any) -> tuple[str, str]:
    return "{" + converter.seq(element.find(_m("e"))) + "}", "atom"


def phantom(converter: _Converter, element: Any) -> tuple[str, str]:
    base = converter.seq(element.find(_m("e")))
    show = converter.prop(element, "phantPr", "show")
    if show is not None and _is_on(show):
        return base, "atom"
    return f"\\phantom{{{base}}}", "atom"
