# LaTeX Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** LaTeX-формулы из Markdown попадают в DOCX настоящими уравнениями Word (OMML), которые можно открыть в редакторе формул и редактировать, а не строкой с долларами; всё, что конвертер не умеет, попадает дословно и с предупреждением, а не молча портится.

**Architecture:** Новый самостоятельный модуль `latex_omml.py` — токенизатор → парсер в AST → эмиттер OMML, без зависимости от `python-docx` кроме `OxmlElement`/`qn`. `gfm_docx_renderer.py` включает плагины `dollarmath` и `amsmath`, получает изолированные math-токены и отдаёт их содержимое в `latex_omml`. Неподдерживаемая конструкция поднимает `UnsupportedLatexError`, рендерер ловит её, пишет формулу моноширинным текстом дословно и добавляет предупреждение в `self.warnings` — ровно тот же паттерн, что уже используется для картинок.

**Tech Stack:** Python 3.12, python-docx 1.1.2, markdown-it-py 4.0.0, mdit-py-plugins 0.5.0 (`dollarmath`, `amsmath` — уже установлены), unittest.

## Global Constraints

- Интерпретатор: **только** `/opt/anaconda3/envs/mdtoword/bin/python`. Системный `python3` падает — две копии Qt.
- Тесты (pytest в env нет, `unittest discover` не работает — нет `tests/__init__.py`):
  ```bash
  cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter tests.test_latex_omml
  ```
  Модуль `tests.test_latex_omml` появляется в Task 2 — до этого его в команде нет.
- **Новых зависимостей не добавлять.** `MML2OMML.XSL` принадлежит Microsoft Office, распространять его нельзя; готовой Python-библиотеки LaTeX→OMML не существует (`mathml2omml` — JavaScript, Spire.Doc — коммерческая). Поэтому конвертер пишем сами.
- GUI (`md_to_word_converter.py`, `gui_theme.py`) не трогать — эта работа отревьюена и закрыта.
- Ничего не портить молча: любая неподдерживаемая конструкция обязана дать и дословный текст, и предупреждение.
- Коммит после каждой задачи, стиль `feat:` / `fix:` / `chore:` / `test:`.

## Подтверждённые факты (замеры 2026-07-20)

| Факт | Доказательство |
|---|---|
| Формулы сейчас портятся | `$a\,b$` → `$a,b$`; `$50\%$` → `$50%$`; `$\{x\}$` → `${x}$`; **`$a*b*c$` → `$abc$`** — звёздочки съедены как курсив |
| Нативных формул нет вообще | в сгенерированном DOCX `m:oMath` встречается 0 раз |
| Плагины решают проблему порчи | с `dollarmath`+`amsmath` токены приходят как `math_inline 'a\,b'`, `math_inline 'a*b*c'`, `math_inline '50\%'`, `math_block '\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}'`, `amsmath '\begin{equation}...\end{equation}'` — содержимое байт-в-байт |
| Плагины уже установлены | `mdit_py_plugins` 0.5.0 содержит `dollarmath`, `amsmath`, `texmath` |
| Ручной OMML работает | собранная вручную дробь сохраняется и переоткрывается: `m:oMath`=1, `m:f`=1, `m:num`/`m:den`=1/1; namespace `m` уже зарегистрирован в `docx.oxml.ns.nsmap` |
| Внутри code fence всё цело | ```` ```latex ```` переносится дословно, включая `\\` |

---

### Task 1: Формулы перестают портиться

Первый шаг закрывает потерю данных независимо от того, насколько далеко зайдёт конвертер OMML. Формулы пока остаются текстом, но текстом верным.

**Files:**
- Modify: `gfm_docx_renderer.py` (импорты, сборка парсера в `render`, `_render_inline`, `_render_block`)
- Test: `tests/test_gfm_docx_renderer.py`

**Interfaces:**
- Produces: метод `GfmDocxRenderer._render_math_literal(latex: str, display: bool) -> None` — пишет формулу моноширинным текстом. Task 4 заменит его тело на попытку OMML с откатом на этот же вывод.

- [ ] **Step 1: Write the failing tests**

```python
    def test_inline_math_survives_verbatim(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            r"Формула $a\,b$ и $a*b*c$ и $50\%$ в строке."
        )
        text = document.paragraphs[0].text
        self.assertIn(r"a\,b", text)
        self.assertIn("a*b*c", text)
        self.assertIn(r"50\%", text)

    def test_display_math_survives_verbatim(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "$$\n\\int_0^\\infty e^{-x^2}\\,dx = \\frac{\\sqrt{\\pi}}{2}\n$$\n"
        )
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn(r"\,dx", text)
        self.assertIn(r"\frac{\sqrt{\pi}}{2}", text)

    def test_amsmath_environment_survives_verbatim(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "\\begin{equation}\n    F = G\\frac{m_1 m_2}{r^2}\n\\end{equation}\n"
        )
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn(r"G\frac{m_1 m_2}{r^2}", text)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gfm_docx_renderer -v`
Expected: FAIL — в тексте будет `a,b`, `abc`, `50%`, потому что backslash-escape и emphasis съедают символы.

- [ ] **Step 3: Write the implementation**

Импорты в `gfm_docx_renderer.py`:

```python
from mdit_py_plugins.amsmath import amsmath_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
```

Сборка парсера в `render`:

```python
        parser = (
            MarkdownIt("js-default", {"breaks": True, "html": False, "linkify": True})
            .enable("linkify")
            .use(footnote_plugin)
            .use(dollarmath_plugin)
            .use(amsmath_plugin)
        )
```

В `_render_inline`, в цикле по `children`, добавить ветку перед `else`:

```python
            elif token_type in {"math_inline", "math_inline_double"}:
                self._render_math_literal(token.content, display=False)
```

В `_render_block`, рядом с обработкой `fence`, добавить:

```python
        if token_type in {"math_block", "math_block_label", "amsmath"}:
            self._render_math_literal(token.content, display=True)
            return
```

Новый метод рядом с `_render_code_block`:

```python
    def _render_math_literal(self, latex: str, display: bool) -> None:
        """Write a formula as verbatim monospace text, preserving every character."""
        text = latex.strip("\n")
        if display:
            paragraph = self.document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run(text)
        else:
            if self._paragraph is None:
                self._paragraph = self._new_paragraph()
            run = self._paragraph.add_run(text)
        run.font.name = "Courier New"
        run.font.size = Pt(10)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: три новых теста прибавились к текущему счёту, всё зелёное.

- [ ] **Step 6: Commit**

```bash
git add gfm_docx_renderer.py tests/test_gfm_docx_renderer.py
git commit -m "fix: stop Markdown from corrupting LaTeX formulas"
```

---

### Task 2: `latex_omml.py` — токенизатор, парсер и базовый эмиттер

**Files:**
- Create: `latex_omml.py`
- Test: `tests/test_latex_omml.py`

**Interfaces:**
- Produces:
  - `class UnsupportedLatexError(ValueError)` — поднимается на всём, что конвертер не умеет; сообщение содержит саму непонятую конструкцию.
  - `def latex_to_omml(latex: str) -> Any` — возвращает элемент `<m:oMath>` со всем содержимым формулы. Поднимает `UnsupportedLatexError`.
  - `def omml_children(latex: str) -> list[Any]` — тот же разбор, но возвращает список дочерних элементов без обёртки; нужен Task 3 для матриц и Task 4 для `oMathPara`.

**Требуемое покрытие этой задачи:** числа, идентификаторы, операторы, группировка `{}`, `^`/`_` (включая одновременные), `\frac` (и `\dfrac`/`\tfrac`), `\sqrt` с необязательной степенью `\sqrt[n]{}`, таблица символов (греческие буквы и распространённые операторы), `\text`/`\mathrm`/`\mathbf`/`\mathit`, пробельные команды.

- [ ] **Step 1: Write the failing tests**

Создать `tests/test_latex_omml.py`:

```python
import unittest

from latex_omml import UnsupportedLatexError, latex_to_omml


def xml_of(latex: str) -> str:
    return latex_to_omml(latex).xml


class LatexToOmmlTests(unittest.TestCase):
    def test_plain_run_marks_identifiers_italic_and_numbers_upright(self):
        xml = xml_of("x + 2")
        self.assertIn("<m:t>x</m:t>", xml)
        self.assertIn("<m:t>2</m:t>", xml)
        self.assertIn('<m:sty m:val="i"/>', xml)

    def test_fraction_builds_num_and_den(self):
        xml = xml_of(r"\frac{a+b}{2}")
        self.assertIn("<m:f>", xml)
        self.assertIn("<m:num>", xml)
        self.assertIn("<m:den>", xml)
        self.assertIn("<m:t>a</m:t>", xml)
        self.assertIn("<m:t>2</m:t>", xml)

    def test_superscript_and_subscript(self):
        self.assertIn("<m:sSup>", xml_of("x^2"))
        self.assertIn("<m:sSub>", xml_of("a_1"))
        self.assertIn("<m:sSubSup>", xml_of("x_i^2"))

    def test_square_root_with_and_without_degree(self):
        plain = xml_of(r"\sqrt{2}")
        self.assertIn("<m:rad>", plain)
        self.assertIn('<m:degHide m:val="1"/>', plain)
        cubic = xml_of(r"\sqrt[3]{8}")
        self.assertIn("<m:deg>", cubic)
        self.assertNotIn('<m:degHide m:val="1"/>', cubic)

    def test_greek_and_operators_map_to_unicode(self):
        self.assertIn("<m:t>α</m:t>", xml_of(r"\alpha"))
        self.assertIn("<m:t>∞</m:t>", xml_of(r"\infty"))
        self.assertIn("<m:t>≤</m:t>", xml_of(r"\leq"))
        self.assertIn("<m:t>×</m:t>", xml_of(r"\times"))

    def test_text_command_is_upright(self):
        xml = xml_of(r"\text{если}")
        self.assertIn("<m:t>если</m:t>", xml)
        self.assertIn('<m:nor m:val="1"/>', xml)

    def test_unsupported_command_raises_with_the_offending_token(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\qedsymbol{x}")
        self.assertIn("qedsymbol", str(caught.exception))

    def test_unbalanced_braces_raise(self):
        with self.assertRaises(UnsupportedLatexError):
            latex_to_omml(r"\frac{a}{b")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_latex_omml -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'latex_omml'`.

- [ ] **Step 3: Write the implementation**

Создать `latex_omml.py`. Ниже — полный код служебных частей, которые легко сделать неправильно; тело парсера пишется по спецификации под ними.

Шапка модуля, элементы OMML и таблица символов:

```python
from __future__ import annotations

import re
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class UnsupportedLatexError(ValueError):
    """Raised when a LaTeX construct has no OMML equivalent here."""


_SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ε", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "⋅", "ast": "∗", "star": "⋆",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥",
    "neq": "≠", "ne": "≠", "approx": "≈", "equiv": "≡",
    "sim": "∼", "simeq": "≃", "cong": "≅", "propto": "∝",
    "ll": "≪", "gg": "≫",
    "in": "∈", "notin": "∉", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "supseteq": "⊇", "cup": "∪", "cap": "∩",
    "emptyset": "∅", "varnothing": "∅", "setminus": "∖",
    "forall": "∀", "exists": "∃", "neg": "¬",
    "land": "∧", "lor": "∨",
    "rightarrow": "→", "to": "→", "leftarrow": "←",
    "leftrightarrow": "↔", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "Leftrightarrow": "⇔", "mapsto": "↦",
    "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮",
    "ddots": "⋱", "prime": "′", "degree": "°",
    "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ",
    "aleph": "ℵ", "angle": "∠", "perp": "⊥", "parallel": "∥",
    "circ": "∘", "bullet": "∙", "oplus": "⊕", "otimes": "⊗",
}

# Spacing commands become real spaces; Word does its own math spacing anyway.
_SPACING = {",": " ", ";": " ", ":": " ", "!": "", " ": " ",
            "quad": " ", "qquad": "  "}

_UPRIGHT_FUNCTIONS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "det", "dim", "ker",
    "deg", "gcd", "hom", "arg", "max", "min", "sup", "inf",
}

_TOKEN_RE = re.compile(
    r"""
    (?P<command>\\[A-Za-z]+ | \\.)     # \frac, \alpha, \\, \{, \,
  | (?P<number>[0-9]+(?:\.[0-9]+)?)
  | (?P<letter>[A-Za-z])
  | (?P<open>\{) | (?P<close>\})
  | (?P<sup>\^) | (?P<sub>_)
  | (?P<bracket>\[|\])
  | (?P<amp>&)
  | (?P<space>\s+)
  | (?P<other>[^\s])
    """,
    re.VERBOSE,
)
```

Конструкторы OMML — использовать только их, чтобы форма элементов была единообразной:

```python
def _el(tag: str) -> Any:
    return OxmlElement(f"m:{tag}")


def _run(text: str, *, italic: bool = False, upright: bool = False, bold: bool = False) -> Any:
    """One <m:r>. Identifiers are italic, numbers and operators are not."""
    run = _el("r")
    if italic or upright or bold:
        properties = _el("rPr")
        if upright:
            nor = _el("nor")
            nor.set(qn("m:val"), "1")
            properties.append(nor)
        if italic and not upright:
            sty = _el("sty")
            sty.set(qn("m:val"), "bi" if bold else "i")
            properties.append(sty)
        elif bold:
            sty = _el("sty")
            sty.set(qn("m:val"), "b")
            properties.append(sty)
        run.append(properties)
    text_element = _el("t")
    text_element.set(qn("xml:space"), "preserve")
    text_element.text = text
    run.append(text_element)
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
        fraction_type = _el("type")
        fraction_type.set(qn("m:val"), "noBar")
        properties.append(fraction_type)
        fraction.append(properties)
    fraction.append(_wrap("num", numerator))
    fraction.append(_wrap("den", denominator))
    return fraction


def _radical(degree: list[Any] | None, radicand: list[Any]) -> Any:
    radical = _el("rad")
    properties = _el("radPr")
    if degree is None:
        hide = _el("degHide")
        hide.set(qn("m:val"), "1")
        properties.append(hide)
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
```

Спецификация парсера — реализовать рекурсивный спуск, который поверх токенов из `_TOKEN_RE` строит список OMML-элементов:

- `_tokenize(latex)` возвращает список пар `(kind, value)`, отбрасывая `space` (но не `\ ` из `_SPACING`). На незакрытой `{` или лишней `}` поднимать `UnsupportedLatexError`.
- `_parse_group(tokens, index)` читает либо `{...}` целиком, либо ровно один следующий атом. Возвращает `(list[element], next_index)`. Это ровно та функция, что нужна аргументам `\frac`, `\sqrt`, `^`, `_`.
- Основной цикл собирает список элементов. Встретив `^` или `_`, он **снимает последний уже собранный элемент** и делает его базой через `_script`. Два подряд идущих скрипта над одной базой дают `sSubSup` — учесть порядок `x_i^2` и `x^2_i`.
- `number` → `_run(value)`; `letter` → `_run(value, italic=True)`; `other` (операторы, скобки) → `_run(value)`.
- `command`:
  - `\frac`, `\dfrac`, `\tfrac` → два `_parse_group`, затем `_fraction`
  - `\sqrt` → если следующий токен `[`, прочитать степень до `]`, затем группу → `_radical`
  - имя в `_SYMBOLS` → `_run(symbol)`
  - имя в `_SPACING` (после `\`) → `_run(spacing)`
  - имя в `_UPRIGHT_FUNCTIONS` → `_run(name, upright=True)`
  - `\text`, `\mathrm`, `\operatorname` → группа, но её содержимое собирается **как сырой текст**, а не как формула: `_run(raw, upright=True)`
  - `\mathbf` → группа, каждый идентификатор жирным; `\mathit` → курсивом
  - `\left`, `\right`, `\begin`, `\end`, `\sum`, `\prod`, `\int`, `\lim` и прочее → **пока** `UnsupportedLatexError` (Task 3 их добавит)
  - неизвестная команда → `UnsupportedLatexError(f"Unsupported LaTeX command: \\{name}")`

Публичные функции:

```python
def omml_children(latex: str) -> list[Any]:
    """Parse a LaTeX math string into a list of OMML elements."""
    tokens = _tokenize(latex)
    elements, index = _parse_sequence(tokens, 0, stop=None)
    if index != len(tokens):
        raise UnsupportedLatexError(f"Could not parse the whole formula: {latex!r}")
    return elements


def latex_to_omml(latex: str) -> Any:
    """Parse a LaTeX math string into a single <m:oMath> element."""
    math = _el("oMath")
    for child in omml_children(latex):
        math.append(child)
    return math
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_latex_omml -v`
Expected: PASS, все 8 тестов.

- [ ] **Step 5: Run full suite**

Добавить `tests.test_latex_omml` в команду полного прогона. Expected: всё зелёное.

- [ ] **Step 6: Commit**

```bash
git add latex_omml.py tests/test_latex_omml.py
git commit -m "feat: convert core LaTeX math to OMML"
```

---

### Task 3: Крупные конструкции — суммы, пределы, скобки, акценты, матрицы

**Files:**
- Modify: `latex_omml.py`
- Test: `tests/test_latex_omml.py`

**Interfaces:**
- Consumes: `_el`, `_run`, `_wrap`, `_fraction`, `_radical`, `_script`, `_parse_group`, `_parse_sequence`, `UnsupportedLatexError`, `omml_children` — всё из Task 2.
- Produces: расширенный набор команд в том же `_parse_sequence`; никаких новых публичных функций.

- [ ] **Step 1: Write the failing tests**

```python
    def test_nary_sum_carries_limits(self):
        xml = xml_of(r"\sum_{i=1}^{n} i")
        self.assertIn("<m:nary>", xml)
        self.assertIn('<m:chr m:val="∑"/>', xml)
        self.assertIn("<m:sub>", xml)
        self.assertIn("<m:sup>", xml)

    def test_integral_uses_its_own_character(self):
        xml = xml_of(r"\int_0^1 x\,dx")
        self.assertIn("<m:nary>", xml)
        self.assertIn('<m:chr m:val="∫"/>', xml)

    def test_limit_uses_lim_low(self):
        xml = xml_of(r"\lim_{x \to 0} f(x)")
        self.assertIn("<m:limLow>", xml)
        self.assertIn("<m:t>lim</m:t>", xml)

    def test_left_right_delimiters_build_a_delimiter_object(self):
        xml = xml_of(r"\left( \frac{a}{b} \right)")
        self.assertIn("<m:d>", xml)
        self.assertIn('<m:begChr m:val="("/>', xml)
        self.assertIn('<m:endChr m:val=")"/>', xml)

    def test_accents_and_overline(self):
        self.assertIn("<m:acc>", xml_of(r"\hat{x}"))
        self.assertIn("<m:bar>", xml_of(r"\overline{AB}"))
        self.assertIn('<m:chr m:val="⃗"/>', xml_of(r"\vec{v}"))

    def test_binomial_is_a_barless_fraction_in_parentheses(self):
        xml = xml_of(r"\binom{n}{k}")
        self.assertIn('<m:type m:val="noBar"/>', xml)
        self.assertIn("<m:d>", xml)

    def test_pmatrix_builds_a_matrix_in_parentheses(self):
        xml = xml_of(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}")
        self.assertIn("<m:m>", xml)
        self.assertEqual(xml.count("<m:mr>"), 2)
        self.assertIn('<m:begChr m:val="("/>', xml)

    def test_cases_environment(self):
        xml = xml_of(r"\begin{cases} x & x > 0 \\ -x & x \leq 0 \end{cases}")
        self.assertIn("<m:m>", xml)
        self.assertIn('<m:begChr m:val="{"/>', xml)

    def test_still_rejects_genuinely_unknown_commands(self):
        with self.assertRaises(UnsupportedLatexError):
            latex_to_omml(r"\qedsymbol")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_latex_omml -v`
Expected: FAIL — Task 2 поднимает `UnsupportedLatexError` на `\sum`, `\int`, `\lim`, `\left`, `\hat`, `\binom`, `\begin`.

- [ ] **Step 3: Write the implementation**

Добавить конструкторы:

```python
_NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐",
    "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
    "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁",
    "bigotimes": "⨂", "bigvee": "⋁", "bigwedge": "⋀",
}

_ACCENTS = {
    "hat": "̂", "widehat": "̂", "tilde": "̃",
    "widetilde": "̃", "bar": "̄", "vec": "⃗",
    "dot": "̇", "ddot": "̈", "acute": "́", "grave": "̀",
    "check": "̌", "breve": "̆",
}

_MATRIX_DELIMITERS = {
    "matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
    "cases": ("{", ""),
}


def _nary(character: str, sub: list[Any] | None, sup: list[Any] | None, body: list[Any]) -> Any:
    nary = _el("nary")
    properties = _el("naryPr")
    chr_element = _el("chr")
    chr_element.set(qn("m:val"), character)
    properties.append(chr_element)
    limit_location = _el("limLoc")
    limit_location.set(qn("m:val"), "undOvr" if character != "∫" else "subSup")
    properties.append(limit_location)
    if sub is None:
        hide = _el("subHide")
        hide.set(qn("m:val"), "1")
        properties.append(hide)
    if sup is None:
        hide = _el("supHide")
        hide.set(qn("m:val"), "1")
        properties.append(hide)
    nary.append(properties)
    nary.append(_wrap("sub", sub or []))
    nary.append(_wrap("sup", sup or []))
    nary.append(_wrap("e", body))
    return nary


def _delimiter(begin: str, end: str, children: list[Any]) -> Any:
    delimiter = _el("d")
    properties = _el("dPr")
    begin_element = _el("begChr")
    begin_element.set(qn("m:val"), begin)
    properties.append(begin_element)
    end_element = _el("endChr")
    end_element.set(qn("m:val"), end)
    properties.append(end_element)
    delimiter.append(properties)
    delimiter.append(_wrap("e", children))
    return delimiter


def _accent(character: str, base: list[Any]) -> Any:
    accent = _el("acc")
    properties = _el("accPr")
    chr_element = _el("chr")
    chr_element.set(qn("m:val"), character)
    properties.append(chr_element)
    accent.append(properties)
    accent.append(_wrap("e", base))
    return accent


def _overline(base: list[Any]) -> Any:
    bar = _el("bar")
    properties = _el("barPr")
    position = _el("pos")
    position.set(qn("m:val"), "top")
    properties.append(position)
    bar.append(properties)
    bar.append(_wrap("e", base))
    return bar


def _limit_low(base: list[Any], limit: list[Any]) -> Any:
    element = _el("limLow")
    element.append(_wrap("e", base))
    element.append(_wrap("lim", limit))
    return element


def _matrix(rows: list[list[list[Any]]]) -> Any:
    matrix = _el("m")
    for row in rows:
        row_element = _el("mr")
        for cell in row:
            row_element.append(_wrap("e", cell))
        matrix.append(row_element)
    return matrix
```

Расширить `_parse_sequence`, добавив ветки для команд:

- имя в `_NARY` → прочитать необязательные `_`/`^` (в любом порядке, каждый через `_parse_group`), затем распарсить остаток последовательности как тело до конца текущей группы → `_nary`
- `\lim` → прочитать необязательный `_` группой → `_limit_low([_run("lim", upright=True)], limit)`
- имя в `_ACCENTS` → группа → `_accent`
- `\overline` → группа → `_overline`; `\underline` → тот же `_overline`, но `pos` = `bot`
- `\binom` → две группы → `_delimiter("(", ")", [_fraction(a, b, no_bar=True)])`
- `\left` → прочитать следующий токен как открывающий разделитель (`(`, `[`, `\{`, `|`, `\|`, `.` = пусто), распарсить до парного `\right`, прочитать его разделитель → `_delimiter`
- `\begin` → прочитать имя окружения в `{}`; если оно в `_MATRIX_DELIMITERS`, разобрать строки, разделённые `\\`, и ячейки, разделённые `&`, до `\end{имя}` → `_matrix`, обёрнутый в `_delimiter`, если разделители непустые; иначе `UnsupportedLatexError`
- `\\` и `&` вне матрицы → `UnsupportedLatexError` с внятным сообщением

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS, все тесты модуля.

- [ ] **Step 5: Run full suite**

Expected: всё зелёное.

- [ ] **Step 6: Commit**

```bash
git add latex_omml.py tests/test_latex_omml.py
git commit -m "feat: support n-ary operators, delimiters, accents and matrices"
```

---

### Task 4: Рендерер отдаёт настоящие формулы Word

**Files:**
- Modify: `gfm_docx_renderer.py` (`_render_math_literal` → попытка OMML с откатом)
- Test: `tests/test_gfm_docx_renderer.py`

**Interfaces:**
- Consumes: `latex_to_omml`, `omml_children`, `UnsupportedLatexError` из `latex_omml`.

- [ ] **Step 1: Write the failing tests**

```python
    def test_inline_math_becomes_a_real_word_equation(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            r"Энергия $E = mc^2$ покоя."
        )
        xml = document.paragraphs[0]._p.xml
        self.assertIn("oMath", xml)
        self.assertIn("sSup", xml)
        self.assertEqual(warnings, [])

    def test_display_math_becomes_a_centred_equation(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "$$\\frac{a}{b}$$\n"
        )
        xml = "".join(p._p.xml for p in document.paragraphs)
        self.assertIn("oMath", xml)
        self.assertIn("<m:f>", xml)
        self.assertEqual(warnings, [])

    def test_unsupported_formula_falls_back_to_verbatim_text_and_warns(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            r"Формула $\qedsymbol{x}$ здесь."
        )
        self.assertIn(r"\qedsymbol{x}", document.paragraphs[0].text)
        self.assertEqual(len(warnings), 1)
        self.assertIn("qedsymbol", warnings[0])

    def test_amsmath_equation_becomes_an_equation(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "\\begin{equation}\nF = G\\frac{m_1 m_2}{r^2}\n\\end{equation}\n"
        )
        xml = "".join(p._p.xml for p in document.paragraphs)
        self.assertIn("oMath", xml)
        self.assertEqual(warnings, [])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gfm_docx_renderer -v`
Expected: FAIL — Task 1 пишет формулы текстом, `oMath` в XML нет.

- [ ] **Step 3: Write the implementation**

Импорт в `gfm_docx_renderer.py`:

```python
from latex_omml import UnsupportedLatexError, latex_to_omml
```

Заменить `_render_math_literal` парой методов:

```python
    def _render_math(self, latex: str, display: bool) -> None:
        """Insert a real Word equation, falling back to verbatim text."""
        formula = latex.strip("\n").strip()
        if not formula:
            return
        try:
            math_element = latex_to_omml(formula)
        except UnsupportedLatexError as error:
            self.warnings.append(f"Formula kept as text: {error}")
            self._render_math_literal(formula, display)
            return
        if display:
            paragraph = self.document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph._p.append(math_element)
            return
        if self._paragraph is None:
            self._paragraph = self._new_paragraph()
        self._paragraph._p.append(math_element)
```

`_render_math_literal` оставить без изменений — он теперь путь отката.

Перевести обе точки вызова из Task 1 (`math_inline`/`math_inline_double` в `_render_inline`, `math_block`/`math_block_label`/`amsmath` в `_render_block`) на `_render_math`.

Окружения `amsmath` приходят вместе с `\begin{equation}...\end{equation}`. Снять обёртку перед разбором: если содержимое начинается с `\begin{equation}` или `\begin{equation*}`, отрезать первую и последнюю строки и передать середину. Окружения `align`/`gather` разбирать построчно по `\\`, создавая по одному центрированному абзацу с формулой на строку; если строка содержит `&`, убрать его — выравнивание по `&` в OMML этой версией не поддерживается, и это не повод терять формулу.

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: всё зелёное.

- [ ] **Step 6: Commit**

```bash
git add gfm_docx_renderer.py tests/test_gfm_docx_renderer.py
git commit -m "feat: render LaTeX as native Word equations"
```

---

### Task 5: Проверка на реальном документе и пересборка

- [ ] **Step 1: Полный прогон**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter tests.test_latex_omml
```
Expected: всё зелёное.

- [ ] **Step 2: Конвертировать документ со всеми видами формул и подтвердить результат**

Собрать Markdown, содержащий: inline-формулу, display-формулу, `\begin{equation}`, сумму с пределами, интеграл, дробь с корнем, матрицу `pmatrix`, `cases`, греческие буквы, `\text{}` с кириллицей, и одну заведомо неподдерживаемую конструкцию. Конвертировать и распечатать: число элементов `m:oMath` в документе, список предупреждений, и для неподдерживаемой формулы — что она попала в текст дословно.

Expected: `m:oMath` больше нуля для каждой поддерживаемой формулы; ровно одно предупреждение — про заведомо неподдерживаемую; её текст в документе совпадает с исходником посимвольно.

- [ ] **Step 3: Открыть результат и убедиться глазами**

```bash
open <путь к получившемуся docx>
```
Формулы должны отображаться как формулы Word, а не как текст с долларами, и по клику открываться в редакторе уравнений.

- [ ] **Step 4: Пересобрать приложение**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && bash scripts/build_macos.sh
```
Expected: `Готово: .../dist/MDtoWORD.app`, codesign `valid on disk`.

**Важно для сборки:** `latex_omml.py` — новый модуль верхнего уровня. Убедиться, что PyInstaller его забирает (он импортируется напрямую из `gfm_docx_renderer.py`, поэтому анализатор должен найти его сам), и проверить это, запустив собранное приложение на файле с формулой, а не только по факту успешной сборки.
