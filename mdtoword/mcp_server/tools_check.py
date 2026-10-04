"""Инструмент check_latex: станет ли формула настоящим уравнением Word."""

from __future__ import annotations

import re
from typing import Annotated

import anyio
from mcp.types import ToolAnnotations
from mdit_py_plugins.amsmath import ENVIRONMENTS as _AMSMATH_ENVIRONMENTS
from pydantic import Field

from .. import latex_omml
from ..converters import ConversionError
from .factory import formula_converter
from .models import LatexCheckReport, LatexCheckResult
from .server import mcp

# (открывающий, закрывающий, строчная ли формула). Порядок важен: «$$»
# проверяется раньше «$».
_MATH_DELIMITERS = (
    ("$$", "$$", False),
    ("\\[", "\\]", False),
    ("\\(", "\\)", True),
    ("$", "$", True),
)
_ENVIRONMENT_BLOCK = re.compile(
    r"^\\begin\{(?P<name>[A-Za-z]+)\*?\}.*\\end\{(?P=name)\*?\}$", re.DOTALL
)


def _strip_math_delimiters(formula: str) -> tuple[str, bool]:
    """Снять ограничители; вернуть (LaTeX, строчная ли формула).

    Без ограничителей формула считается выключной. Пробелы внутри строчной
    формулы сохраняются: для ``$…$`` они значимы при разборе.
    """
    text = formula.strip()
    for opening, closing, inline in _MATH_DELIMITERS:
        if (
            len(text) >= len(opening) + len(closing)
            and text.startswith(opening)
            and text.endswith(closing)
        ):
            inner = text[len(opening) : len(text) - len(closing)]
            return (inner if inline and inner.strip() else inner.strip()), inline
    return text, False


def _formula_as_markdown(latex: str) -> str:
    """Записать формулу так, как её записал бы автор документа.

    Окружение amsmath верхнего уровня (``align``, ``gather``…) — отдельным
    блоком: его разворачивает рендерер, а не ``latex_omml``. Всё остальное —
    display-формулой в ``$$``.
    """
    match = _ENVIRONMENT_BLOCK.match(latex)
    if match and match.group("name") in _AMSMATH_ENVIRONMENTS:
        return latex + "\n"
    return f"$$\n{latex}\n$$\n"


def _render_verdict(formula: str, markdown: str, fallback_error: str) -> LatexCheckResult:
    """Отрендерить крошечный документ и вынести вердикт по формуле в нём.

    Годна, если рендер прошёл без предупреждений и в документе действительно
    появилось уравнение; иначе ``error`` — текст предупреждений рендерера.
    """
    converter = formula_converter()
    try:
        # _render — внутри пакета: нужен сам документ, чтобы убедиться, что
        # уравнение в нём есть, а не просто нет предупреждений.
        document, warnings = converter._render(markdown, None)
    except ConversionError as error:
        return LatexCheckResult(formula=formula, ok=False, error=str(error))
    if warnings:
        return LatexCheckResult(
            formula=formula, ok=False, error="; ".join(str(w) for w in warnings)
        )
    if not document.element.body.xpath(".//m:oMath"):
        return LatexCheckResult(formula=formula, ok=False, error=fallback_error)
    return LatexCheckResult(formula=formula, ok=True)


def _check_formula(formula: str) -> LatexCheckResult:
    """Проверить одну формулу так, как её обработает markdown_to_word.

    Строчная формула (``$…$``, ``\\(…\\)``) всегда идёт через рендерер в
    абзаце из одной формулы: у строчной математики есть эвристика «это проза,
    а не формула» (кириллица, слова без математических знаков), и такой
    ``$…$`` остаётся текстом, хотя ``latex_to_omml`` его бы принял.

    Выключная — сначала быстрым путём через ``latex_to_omml`` (с отделённым
    ``\\tag``, если ``latex_omml`` уже умеет его отделять). Если он отказал,
    вердикт выносит рендерер: окружения amsmath, ``\\tag`` и прочее, что
    рендерер разбирает сам, иначе давали бы ложный отказ.
    """
    latex, inline = _strip_math_delimiters(formula)
    if not latex:
        return LatexCheckResult(formula=formula, ok=False, error="Empty formula")
    if inline:
        return _render_verdict(
            formula,
            f"${latex}$\n",
            "Not recognised as inline math: $…$ must not start or end with a "
            "space and must not contain an unescaped $",
        )

    split_tag = getattr(latex_omml, "split_equation_tag", None)
    try:
        body = split_tag(latex)[0] if split_tag is not None else latex
        latex_omml.latex_to_omml(body)
    except Exception as error:  # noqa: BLE001 — подробности даст рендерер
        direct_error = str(error) or type(error).__name__
    else:
        return LatexCheckResult(formula=formula, ok=True)
    return _render_verdict(formula, _formula_as_markdown(latex), direct_error)


_CHECK_LATEX_TITLE = "Check LaTeX formulas"


@mcp.tool(
    title=_CHECK_LATEX_TITLE,
    annotations=ToolAnnotations(
        title=_CHECK_LATEX_TITLE,
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def check_latex(
    formulas: Annotated[
        list[str],
        Field(
            min_length=1,
            description=(
                "LaTeX formulas, one per item, as written in the document: $…$ or "
                "\\(…\\) is checked as inline math; $$…$$, \\[…\\] or bare as display math"
            ),
        ),
    ],
) -> LatexCheckReport:
    """Check that LaTeX formulas will become native Word equations.

    Use it before writing a document (or while fixing one) to make sure
    every formula converts to a real, editable Word equation (OMML) instead
    of being kept as verbatim text with a warning. Pass each formula exactly
    as it appears in the document. `$…$` or `\\(…\\)` is judged as INLINE
    math, including the renderer's prose check: an inline formula with
    Cyrillic words or with only plain words (`$x = длина$`, `$PATH and
    HOME$`) stays text, so it is reported as not ok. `$$…$$`, `\\[…\\]` or
    no delimiters is judged as display math. Whole amsmath environments
    (`\\begin{align}…\\end{align}` and the like) and `\\tag{n}` are checked
    the same way `markdown_to_word` treats them.

    Each result says `ok`, and if not, `error` names the construct that
    failed (e.g. an unsupported command) so you can rewrite it with a
    supported equivalent and check again. Reads and writes no files.
    """

    def check_all() -> LatexCheckReport:
        results = [_check_formula(formula) for formula in formulas]
        return LatexCheckReport(
            results=results, all_ok=all(result.ok for result in results)
        )

    return await anyio.to_thread.run_sync(check_all)
