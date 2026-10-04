"""Equations: OMML to LaTeX with warnings counted, and display-equation paragraphs."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from ..omml_latex import equation_environment, equation_rows, omml_to_latex
from .wordml import (
    M_OMATH,
    M_OMATHPARA,
    W_BOOKMARKSTART,
    W_DRAWING,
    W_ENDNOTEREFERENCE,
    W_FLDCHAR,
    W_FOOTNOTEREFERENCE,
    W_OBJECT,
    W_PICT,
    W_PPR,
    W_R,
    W_T,
    W_TAB,
    _q,
)

if TYPE_CHECKING:
    from .classify import _ParaInfo
    from .diagnostics import _Diagnostics


def latex(diagnostics: _Diagnostics, element: Any) -> str:
    """One formula as LaTeX, its unconverted elements counted as warnings."""
    found: list[str] = []
    result = omml_to_latex(element, found)
    diagnostics.formula_warnings(found)
    return result


def latex_rows(diagnostics: _Diagnostics, element: Any) -> list[str] | None:
    """The rows of a formula that is one equation array, else ``None``."""
    found: list[str] = []
    rows = equation_rows(element, found)
    diagnostics.formula_warnings(found)
    return rows


def display_math(diagnostics: _Diagnostics, element: Any, info: _ParaInfo) -> list[str] | None:
    """Blocks for a paragraph that is one display equation, else ``None``."""
    maths: list[Any] = []
    other: list[str] = []
    for child in element:
        tag = child.tag
        if tag in (M_OMATH, M_OMATHPARA):
            maths.append(child)
        elif tag == W_R:
            if any(grandchild.tag in (W_DRAWING, W_PICT, W_OBJECT, W_FOOTNOTEREFERENCE,
                                      W_ENDNOTEREFERENCE, W_FLDCHAR)
                   for grandchild in child):
                return None
            other.append("".join(
                (t.text or "") if t.tag == W_T else " "
                for t in child if t.tag in (W_T, W_TAB)
            ))
        elif tag in (W_PPR, W_BOOKMARKSTART, _q("w:bookmarkEnd"), _q("w:proofErr")):
            continue
        elif isinstance(tag, str):
            return None
    if len(maths) != 1:
        return None
    rest = "".join(other).strip()
    label_match = re.fullmatch(r"\((.+)\)", rest)
    if rest and label_match is None:
        return None
    math = maths[0]
    if math.tag == M_OMATH and not info.centered and label_match is None:
        return None
    label = label_match.group(1).strip() if label_match else ""
    formulas = [math] if math.tag == M_OMATH else math.findall(M_OMATH)
    blocks = []
    for index, formula in enumerate(formulas):
        rows = latex_rows(diagnostics, formula)
        if rows is not None and len(rows) > 1:
            environment = equation_environment(rows)
            text = (f"\\begin{{{environment}}}\n" + " \\\\\n".join(rows)
                    + f"\n\\end{{{environment}}}")
        else:
            text = latex(diagnostics, formula)
        if not text:
            continue
        suffix = f" ({label})" if label and index == len(formulas) - 1 else ""
        blocks.append(f"$$\n{text}\n$${suffix}")
    return blocks
