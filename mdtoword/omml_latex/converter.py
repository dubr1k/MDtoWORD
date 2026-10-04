"""The OMML walker: traversal, text runs, warnings and the element dispatch table.

Each OMML element is converted by a handler ``(converter, element) ->
(latex, kind)`` looked up in :data:`_HANDLERS`; the composite elements live
in :mod:`.structures` and :mod:`.matrices`, the containers and text runs
here. This module also holds the public entry points.
"""

from __future__ import annotations

import re
from typing import Any

from ..errors import ConversionWarning
from .matrices import delimiter, equation_array, equation_array_rows, matrix
from .names import _M, _SKIPPED_W, _TRANSPARENT_W, _VAL, _is_on, _local, _m, _w
from .structures import (
    accent,
    bar,
    border_box,
    box,
    fraction,
    function,
    group_character,
    limit_lower,
    limit_upper,
    nary,
    phantom,
    pre_script,
    radical,
    sub_superscript,
    subscript,
    superscript,
)
from .symbols import _SPACE_RUNS
from .text import _escape_text_mode, _function_command, _join, _text_latex

_LATIN_WORD = re.compile(r"[A-Za-z][A-Za-z0-9]*")


class _Converter:
    """One conversion: holds the warnings raised while walking the tree."""

    def __init__(self) -> None:
        self.partial: dict[str, int] = {}
        # Set while converting the name of an m:func, whose runs are then
        # function names (\sin, \operatorname{...}) rather than variables.
        self.in_function_name = False

    # -- traversal --------------------------------------------------------

    def children(self, element: Any) -> list[Any]:
        """Content children of an OMML container, transparent wrappers opened."""
        result: list[Any] = []
        for child in element:
            tag = child.tag
            if not isinstance(tag, str):  # comments, processing instructions
                continue
            if tag in _SKIPPED_W:
                continue
            if tag in _TRANSPARENT_W:
                result.extend(self.children(child))
                continue
            if tag == _w("sdt"):
                content = child.find(_w("sdtContent"))
                if content is not None:
                    result.extend(self.children(content))
                continue
            if tag.startswith(f"{{{_M}}}"):
                local = _local(tag)
                if local.endswith("Pr") or local in ("ctrlPr", "argPr"):
                    continue
            result.append(child)
        return result

    def seq(self, container: Any) -> str:
        """LaTeX for the content of an ``m:e``-like container."""
        if container is None:
            return ""
        pieces: list[tuple[str, str]] = [self.node(child) for child in self.children(container)]
        out: list[str] = []
        for index, (latex, kind) in enumerate(pieces):
            # An n-ary operator's operand runs to the end of its group when
            # parsed back; brace it so what follows stays outside.
            if kind == "nary" and any(text for text, _ in pieces[index + 1:]):
                latex = "{" + latex + "}"
            out.append(latex)
        return _join(out).strip()

    def node(self, element: Any) -> tuple[str, str]:
        """(LaTeX, kind) for one element; kind is ``nary`` or ``atom``."""
        tag = element.tag
        handler = _HANDLERS.get(tag)
        if handler is not None:
            return handler(self, element)
        if tag == _w("r"):
            return self.word_run(element), "atom"
        return self.unknown(element), "atom"

    def unknown(self, element: Any) -> str:
        name = _local(element.tag) or "?"
        self.partial[name] = self.partial.get(name, 0) + 1
        text = "".join(t.text or "" for t in element.iter(_m("t")))
        return _text_latex(text)

    # -- properties ---------------------------------------------------------

    def prop(self, element: Any, properties: str, name: str) -> Any:
        """The ``name`` child of the element's ``properties`` (``m:fPr`` ...), if any."""
        holder = element.find(_m(properties))
        return holder.find(_m(name)) if holder is not None else None

    def prop_value(self, element: Any, properties: str, name: str,
                   default: str | None = None) -> str | None:
        """The ``m:val`` of :meth:`prop`, ``default`` when the property is absent."""
        child = self.prop(element, properties, name)
        if child is None:
            return default
        return child.get(_VAL, "")

    # -- runs ---------------------------------------------------------------

    def run(self, element: Any) -> tuple[str, str]:
        text = "".join(t.text or "" for t in element.findall(_m("t")))
        properties = element.find(_m("rPr"))
        normal = style = script = None
        aligned = line_break = False
        if properties is not None:
            normal = _is_on(properties.find(_m("nor"))) or None
            sty = properties.find(_m("sty"))
            style = sty.get(_VAL) if sty is not None else None
            scr = properties.find(_m("scr"))
            script = scr.get(_VAL) if scr is not None else None
            aligned = properties.find(_m("aln")) is not None
            line_break = properties.find(_m("brk")) is not None
        body = self.run_text(text, normal=bool(normal), style=style, script=script)
        prefix = ""
        if line_break:
            prefix += r"\\ "
        if aligned:
            prefix += "&"
        return prefix + body, "atom"

    def run_text(self, text: str, *, normal: bool, style: str | None,
                 script: str | None) -> str:
        if not text:
            return ""
        if self.in_function_name:
            command = _function_command(text)
            if command:
                return command
            return r"\operatorname{" + _escape_text_mode(text.strip()) + "}"
        if normal:
            return _function_command(text) or (r"\text{" + _escape_text_mode(text) + "}")
        if not text.strip():
            return _SPACE_RUNS.get(text, r"\ " * len(text))
        if script in ("double-struck", "script", "fraktur", "sans-serif", "monospace"):
            command = {"double-struck": "mathbb", "script": "mathcal",
                       "fraktur": "mathfrak", "sans-serif": "mathsf",
                       "monospace": "mathtt"}[script]
            return f"\\{command}{{{_text_latex(text)}}}"
        if style == "p" or script == "roman":
            command = _function_command(text)
            if command:
                return command
            if _LATIN_WORD.fullmatch(text) and not text.isdigit():
                return r"\mathrm{" + text + "}"
            return _text_latex(text)
        if style == "b":
            return r"\mathbf{" + _text_latex(text) + "}"
        if style == "bi":
            return r"\boldsymbol{" + _text_latex(text) + "}"
        if style == "i" and not any(ch.isalpha() for ch in text):
            return r"\mathit{" + _text_latex(text) + "}"
        return _text_latex(text)

    def word_run(self, element: Any) -> str:
        """A WordprocessingML run inside a formula: ordinary text."""
        text = "".join(t.text or "" for t in element.iter(_w("t")))
        if not text:
            return ""
        return r"\text{" + _escape_text_mode(text) + "}"

    # -- containers -----------------------------------------------------------

    def math(self, element: Any) -> tuple[str, str]:
        return self.seq(element), "atom"

    def math_paragraph(self, element: Any) -> tuple[str, str]:
        lines = [self.seq(math) for math in element.findall(_m("oMath"))]
        return r" \\ ".join(line for line in lines if line), "atom"

    # -- warnings -----------------------------------------------------------

    def warnings(self) -> list[ConversionWarning]:
        if not self.partial:
            return []
        names = ", ".join(
            f"m:{name}" + (f" (x{count})" if count > 1 else "")
            for name, count in sorted(self.partial.items())
        )
        return [ConversionWarning(
            f"Equation elements not understood, kept as plain text: {names}",
            code="formula_partial",
        )]


_HANDLERS = {
    _m("r"): _Converter.run,
    _m("f"): fraction,
    _m("rad"): radical,
    _m("sSup"): superscript,
    _m("sSub"): subscript,
    _m("sSubSup"): sub_superscript,
    _m("sPre"): pre_script,
    _m("nary"): nary,
    _m("d"): delimiter,
    _m("m"): matrix,
    _m("eqArr"): equation_array,
    _m("acc"): accent,
    _m("bar"): bar,
    _m("func"): function,
    _m("limLow"): limit_lower,
    _m("limUpp"): limit_upper,
    _m("groupChr"): group_character,
    _m("borderBox"): border_box,
    _m("box"): box,
    _m("phant"): phantom,
    _m("oMath"): _Converter.math,
    _m("oMathPara"): _Converter.math_paragraph,
}


def _clean(latex: str) -> str:
    return re.sub(r"[ \t]{2,}", " ", latex).strip()


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------


def omml_to_latex(element: Any, warnings: list[str] | None = None) -> str:
    """Convert an OMML element -- usually ``m:oMath`` or ``m:oMathPara`` -- to LaTeX.

    The result has no surrounding ``$``. Elements that cannot be converted
    keep their text; when ``warnings`` is given, a ``formula_partial``
    :class:`~mdtoword.errors.ConversionWarning` naming them is appended.
    """
    converter = _Converter()
    latex, _ = converter.node(element)
    if warnings is not None:
        warnings.extend(converter.warnings())
    return _clean(latex)


def equation_rows(element: Any, warnings: list[str] | None = None) -> list[str] | None:
    """Rows of a formula that is exactly one equation array, else ``None``.

    A display equation consisting of one ``m:eqArr`` is what both Word and the
    forward converter make of a multi-line formula. Its rows -- ``&`` marking
    each ``m:aln`` alignment point -- can then be laid out one per line inside
    ``aligned``/``gathered`` (see :func:`equation_environment`), which the
    forward converter turns back into the same equation array.
    """
    converter = _Converter()
    target = element
    if element.tag == _m("oMathPara"):
        maths = element.findall(_m("oMath"))
        if len(maths) != 1:
            return None
        target = maths[0]
    if target.tag != _m("oMath"):
        return None
    children = converter.children(target)
    if len(children) != 1 or children[0].tag != _m("eqArr"):
        return None
    rows = [_clean(row) for row in equation_array_rows(converter, children[0])]
    if warnings is not None:
        warnings.extend(converter.warnings())
    return rows
