"""Tests for OMML -> LaTeX conversion (``mdtoword.omml_latex``).

Two kinds of test: per-construct checks on hand-written OMML in the shapes
Word itself produces, and a round-trip corpus -- LaTeX through
``latex_to_omml``, back through ``omml_to_latex`` and forward again -- that
must rebuild a structurally identical OMML tree.
"""

from __future__ import annotations

import unittest

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from lxml import etree

from mdtoword.errors import ConversionWarning
from mdtoword.latex_omml import latex_to_omml
from mdtoword.omml_latex import equation_rows, omml_to_latex


def _math(body: str):
    return parse_xml(f"<m:oMath {nsdecls('m', 'w')}>{body}</m:oMath>")


def _r(text: str, properties: str = "") -> str:
    rpr = f"<m:rPr>{properties}</m:rPr>" if properties else ""
    return f'<m:r>{rpr}<m:t xml:space="preserve">{text}</m:t></m:r>'


def _e(body: str, tag: str = "e") -> str:
    return f"<m:{tag}>{body}</m:{tag}>"


def latex(body: str) -> str:
    return omml_to_latex(_math(body))


class RunTests(unittest.TestCase):
    def test_symbols_map_back_to_commands(self) -> None:
        self.assertEqual(latex(_r("α+β≤γ")), r"\alpha+\beta \leq \gamma")
        self.assertEqual(latex(_r("x→∞")), r"x \to \infty")
        self.assertEqual(latex(_r("ϕφϵε")), r"\phi\varphi\epsilon\varepsilon")

    def test_command_before_a_letter_gets_a_space(self) -> None:
        self.assertEqual(latex(_r("∂") + _r("x", '<m:sty m:val="i"/>')), r"\partial x")

    def test_latex_special_characters_are_escaped(self) -> None:
        self.assertEqual(latex(_r("{%}#&amp;$_")), r"\{\%\}\#\&\$\_")

    def test_minus_and_primes_become_ascii(self) -> None:
        self.assertEqual(latex(_r("f′(x)−1")), "f'(x)-1")

    def test_normal_text_and_function_names(self) -> None:
        normal = '<m:nor m:val="1"/>'
        self.assertEqual(latex(_r("если ", normal) + _r("x")), r"\text{если }x")
        self.assertEqual(latex(_r("sin", normal) + _r("x")), r"\sin x")
        self.assertEqual(latex(_r("lim sup", normal)), r"\limsup")
        self.assertEqual(latex(_r("a_b", normal)), r"\text{a\_b}")

    def test_run_styles(self) -> None:
        self.assertEqual(latex(_r("d", '<m:sty m:val="p"/>')), r"\mathrm{d}")
        self.assertEqual(latex(_r("cos", '<m:sty m:val="p"/>')), r"\cos")
        self.assertEqual(latex(_r("v", '<m:sty m:val="b"/>')), r"\mathbf{v}")
        self.assertEqual(latex(_r("α", '<m:sty m:val="bi"/>')), r"\boldsymbol{\alpha}")
        self.assertEqual(latex(_r("2", '<m:sty m:val="i"/>')), r"\mathit{2}")
        self.assertEqual(
            latex(_r("R", '<m:scr m:val="double-struck"/><m:sty m:val="p"/>')),
            r"\mathbb{R}",
        )

    def test_unicode_mathematical_alphanumerics(self) -> None:
        self.assertEqual(latex(_r("𝑥ℎ")), "xh")
        self.assertEqual(latex(_r("ℝ")), r"\mathbb{R}")
        self.assertEqual(latex(_r("𝐯𝐰")), r"\mathbf{vw}")
        self.assertEqual(latex(_r("ℒ")), r"\mathcal{L}")
        self.assertEqual(latex(_r("𝛼")), r"\alpha")

    def test_space_runs(self) -> None:
        self.assertEqual(latex(_r("a") + _r(" ") + _r("b")), r"a\,b")
        self.assertEqual(latex(_r("a b")), r"a\ b")

    def test_tracked_insertions_count_and_deletions_do_not(self) -> None:
        body = ("<w:ins>" + _r("a") + "</w:ins>"
                "<w:del>" + _r("b") + "</w:del>" + _r("c"))
        self.assertEqual(latex(body), "ac")

    def test_word_run_inside_a_formula_is_text(self) -> None:
        self.assertEqual(latex('<w:r><w:t>word</w:t></w:r>' + _r("x")), r"\text{word}x")


class StructureTests(unittest.TestCase):
    def test_fractions(self) -> None:
        self.assertEqual(latex(f"<m:f>{_e(_r('a'), 'num')}{_e(_r('b'), 'den')}</m:f>"),
                         r"\frac{a}{b}")
        no_bar = '<m:fPr><m:type m:val="noBar"/></m:fPr>'
        self.assertEqual(
            latex(f"<m:f>{no_bar}{_e(_r('n'), 'num')}{_e(_r('k'), 'den')}</m:f>"),
            r"{n \atop k}",
        )
        linear = '<m:fPr><m:type m:val="lin"/></m:fPr>'
        self.assertEqual(
            latex(f"<m:f>{linear}{_e(_r('a'), 'num')}{_e(_r('b'), 'den')}</m:f>"),
            "{a}/{b}",
        )

    def test_radicals(self) -> None:
        hidden = '<m:radPr><m:degHide m:val="1"/></m:radPr>'
        self.assertEqual(latex(f"<m:rad>{hidden}<m:deg/>{_e(_r('x'))}</m:rad>"),
                         r"\sqrt{x}")
        self.assertEqual(latex(f"<m:rad><m:radPr/>{_e(_r('3'), 'deg')}{_e(_r('8'))}</m:rad>"),
                         r"\sqrt[3]{8}")

    def test_scripts(self) -> None:
        self.assertEqual(latex(f"<m:sSup>{_e(_r('x'))}{_e(_r('2'), 'sup')}</m:sSup>"), "x^{2}")
        self.assertEqual(latex(f"<m:sSub>{_e(_r('ab'))}{_e(_r('i'), 'sub')}</m:sSub>"),
                         "{ab}_{i}")
        self.assertEqual(
            latex(f"<m:sSubSup>{_e(_r('x'))}{_e(_r('i'), 'sub')}{_e(_r('2'), 'sup')}</m:sSubSup>"),
            "x_{i}^{2}",
        )
        self.assertEqual(
            latex(f"<m:sPre>{_e(_r('a'), 'sub')}{_e(_r('b'), 'sup')}{_e(_r('X'))}</m:sPre>"),
            "{}_{a}^{b}X",
        )

    def test_nary_operators(self) -> None:
        body = (f"<m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr>"
                f"{_e(_r('i=1'), 'sub')}{_e(_r('n'), 'sup')}{_e(_r('i'))}</m:nary>")
        self.assertEqual(latex(body), r"\sum_{i = 1}^{n} i")
        # No m:chr means an integral; hidden limits are left out.
        hidden = ('<m:naryPr><m:subHide m:val="1"/><m:supHide m:val="1"/></m:naryPr>'
                  '<m:sub/><m:sup/>')
        self.assertEqual(latex(f"<m:nary>{hidden}{_e(_r('f'))}</m:nary>"), r"\int f")
        stacked = ('<m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="undOvr"/></m:naryPr>'
                   f"{_e(_r('0'), 'sub')}{_e(_r('1'), 'sup')}")
        self.assertEqual(latex(f"<m:nary>{stacked}{_e(_r('x'))}</m:nary>"),
                         r"\int\limits_{0}^{1} x")

    def test_operator_followed_by_more_content_is_braced(self) -> None:
        body = (f"<m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr>{_e(_r('i'), 'sub')}"
                f"<m:sup/>{_e(_r('a'))}</m:nary>" + _r("+b"))
        self.assertEqual(latex(body), r"{\sum_{i} a}+b")

    def test_delimiters(self) -> None:
        self.assertEqual(latex(f"<m:d>{_e(_r('x'))}</m:d>"), r"\left(x\right)")
        properties = '<m:dPr><m:begChr m:val="⟨"/><m:endChr m:val=""/></m:dPr>'
        self.assertEqual(latex(f"<m:d>{properties}{_e(_r('x'))}</m:d>"),
                         r"\left\langle x\right.")
        separated = '<m:dPr><m:begChr m:val="{"/><m:sepChr m:val="|"/><m:endChr m:val="}"/></m:dPr>'
        self.assertEqual(latex(f"<m:d>{separated}{_e(_r('x'))}{_e(_r('y'))}</m:d>"),
                         r"\left\{x \middle| y\right\}")

    def test_binomial_and_matrix_environments(self) -> None:
        stack = (f"<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr>"
                 f"{_e(_r('n'), 'num')}{_e(_r('k'), 'den')}</m:f>")
        self.assertEqual(latex(f"<m:d>{_e(stack)}</m:d>"), r"\binom{n}{k}")
        matrix = (f"<m:m><m:mr>{_e(_r('a'))}{_e(_r('b'))}</m:mr>"
                  f"<m:mr>{_e(_r('c'))}{_e(_r('d'))}</m:mr></m:m>")
        self.assertEqual(latex(matrix), r"\begin{matrix} a & b \\ c & d \end{matrix}")
        brackets = '<m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
        self.assertEqual(latex(f"<m:d>{brackets}{_e(matrix)}</m:d>"),
                         r"\begin{bmatrix} a & b \\ c & d \end{bmatrix}")
        bars = '<m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr>'
        self.assertEqual(latex(f"<m:d>{bars}{_e(matrix)}</m:d>"),
                         r"\begin{vmatrix} a & b \\ c & d \end{vmatrix}")

    def test_word_matrix_with_centred_column_properties(self) -> None:
        properties = ('<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/>'
                      '<m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>')
        matrix = f"<m:m>{properties}<m:mr>{_e(_r('1'))}{_e(_r('0'))}</m:mr></m:m>"
        self.assertEqual(latex(matrix), r"\begin{matrix} 1 & 0 \end{matrix}")

    def test_array_with_column_alignment(self) -> None:
        properties = ('<m:mPr><m:mcs>'
                      '<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/></m:mcPr></m:mc>'
                      '<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="right"/></m:mcPr></m:mc>'
                      '</m:mcs></m:mPr>')
        matrix = f"<m:m>{properties}<m:mr>{_e(_r('a'))}{_e(_r('b'))}</m:mr></m:m>"
        self.assertEqual(latex(matrix), r"\begin{array}{lr} a & b \end{array}")

    def test_word_cases_built_from_an_equation_array(self) -> None:
        rows = (f"<m:eqArr>{_e(_r('1') + _r(',', '<m:aln/>') + _r('x>0'))}"
                f"{_e(_r('0') + _r(',', '<m:aln/>') + _r('x≤0'))}</m:eqArr>")
        brace = '<m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr>'
        self.assertEqual(
            latex(f"<m:d>{brace}{_e(rows)}</m:d>"),
            r"\begin{cases}1 &,x > 0 \\ 0 &,x \leq 0\end{cases}",
        )

    def test_accents_and_bars(self) -> None:
        self.assertEqual(latex(f"<m:acc>{_e(_r('x'))}</m:acc>"), r"\hat{x}")
        tilde = '<m:accPr><m:chr m:val="~"/></m:accPr>'
        self.assertEqual(latex(f"<m:acc>{tilde}{_e(_r('y'))}</m:acc>"), r"\tilde{y}")
        vector = '<m:accPr><m:chr m:val="⃗"/></m:accPr>'
        self.assertEqual(latex(f"<m:acc>{vector}{_e(_r('v'))}</m:acc>"), r"\vec{v}")
        self.assertEqual(latex(f"<m:bar>{_e(_r('x'))}</m:bar>"), r"\underline{x}")
        top = '<m:barPr><m:pos m:val="top"/></m:barPr>'
        self.assertEqual(latex(f"<m:bar>{top}{_e(_r('AB'))}</m:bar>"), r"\overline{AB}")

    def test_function_application(self) -> None:
        name = _e(_r("sin", '<m:sty m:val="p"/>'), "fName")
        argument = _e(f"<m:d>{_e(_r('x'))}</m:d>")
        self.assertEqual(latex(f"<m:func>{name}{argument}</m:func>"), r"\sin\left(x\right)")
        plain_name = _e(_r("sgn"), "fName")
        self.assertEqual(latex(f"<m:func>{plain_name}{_e(_r('x'))}</m:func>"),
                         r"\operatorname{sgn} x")
        upright = '<m:sty m:val="p"/>'
        limit = (f"<m:limLow>{_e(_r('lim', upright))}"
                 f"{_e(_r('n→∞'), 'lim')}</m:limLow>")
        self.assertEqual(
            latex(f"<m:func>{_e(limit, 'fName')}{_e(_r('a') + _r('n'))}</m:func>"),
            r"\lim_{n \to \infty}{an}",
        )

    def test_limits_and_group_characters(self) -> None:
        self.assertEqual(
            latex(f"<m:limLow>{_e(_r('x'))}{_e(_r('n'), 'lim')}</m:limLow>"),
            r"\underset{n}{x}",
        )
        brace = f"<m:groupChr>{_e(_r('a+b'))}</m:groupChr>"
        self.assertEqual(latex(brace), r"\underbrace{a+b}")
        self.assertEqual(
            latex(f"<m:limLow>{_e(brace)}{_e(_r('n'), 'lim')}</m:limLow>"),
            r"\underbrace{a+b}_{n}",
        )
        top = '<m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/></m:groupChrPr>'
        over = f"<m:groupChr>{top}{_e(_r('x'))}</m:groupChr>"
        self.assertEqual(
            latex(f"<m:limUpp>{_e(over)}{_e(_r('k'), 'lim')}</m:limUpp>"),
            r"\overbrace{x}^{k}",
        )
        arrow = '<m:groupChrPr><m:chr m:val="→"/></m:groupChrPr>'
        self.assertEqual(latex(f"<m:groupChr>{arrow}{_e(_r('f'))}</m:groupChr>"),
                         r"\xrightarrow{f}")

    def test_boxes_and_phantoms(self) -> None:
        self.assertEqual(latex(f"<m:borderBox>{_e(_r('E'))}</m:borderBox>"), r"\boxed{E}")
        self.assertEqual(latex(f"<m:box>{_e(_r('dx'))}</m:box>"), "{dx}")
        self.assertEqual(latex(f"<m:phant>{_e(_r('x'))}</m:phant>"), r"\phantom{x}")

    def test_equation_arrays(self) -> None:
        aligned = (f"<m:eqArr>{_e(_r('a') + _r('=', '<m:aln/>') + _r('b'))}"
                   f"{_e(_r('c') + _r('=', '<m:aln/>') + _r('d'))}</m:eqArr>")
        self.assertEqual(latex(aligned),
                         r"\begin{aligned}a &= b \\ c &= d\end{aligned}")
        stacked = f"<m:eqArr>{_e(_r('a'))}{_e(_r('b'))}</m:eqArr>"
        self.assertEqual(latex(stacked), r"\begin{gathered}a \\ b\end{gathered}")

    def test_equation_rows(self) -> None:
        aligned = (f"<m:eqArr>{_e(_r('a') + _r('=', '<m:aln/>') + _r('b'))}"
                   f"{_e(_r('c') + _r('=', '<m:aln/>') + _r('d'))}</m:eqArr>")
        self.assertEqual(equation_rows(_math(aligned)), ["a &= b", "c &= d"])
        self.assertIsNone(equation_rows(_math(_r("x"))))

    def test_math_paragraph_stacks_its_equations(self) -> None:
        paragraph = parse_xml(
            f"<m:oMathPara {nsdecls('m')}><m:oMath>{_r('a')}</m:oMath>"
            f"<m:oMath>{_r('b')}</m:oMath></m:oMathPara>"
        )
        self.assertEqual(omml_to_latex(paragraph), r"a \\ b")

    def test_unknown_element_keeps_its_text_and_warns(self) -> None:
        warnings: list[str] = []
        result = omml_to_latex(_math(_r("x=") + f"<m:mystery>{_r('42')}</m:mystery>"), warnings)
        self.assertEqual(result, "x = 42")
        self.assertEqual(len(warnings), 1)
        self.assertIsInstance(warnings[0], ConversionWarning)
        self.assertEqual(warnings[0].code, "formula_partial")
        self.assertIn("m:mystery", warnings[0])

    def test_no_warning_for_known_constructs(self) -> None:
        warnings: list[str] = []
        omml_to_latex(latex_to_omml(r"\frac{a}{b} + \sqrt{x}"), warnings)
        self.assertEqual(warnings, [])


def _canonical(element) -> bytes:
    return etree.tostring(element, method="c14n")


# Constructs listed as supported in the README's "What's supported" table.
ROUND_TRIP_CORPUS = [
    r"E = mc^2", r"x_i^2", r"a_{ij}", r"x^{n+1}", r"e^{-x^2}", r"\alpha_{1}^{2}",
    r"\frac{a}{b}", r"\frac{1}{2}x", r"\dfrac{1}{2}", r"\frac{1}{1+\frac{1}{x}}",
    r"\sqrt{x}", r"\sqrt[3]{8}", r"\sqrt{a^2+b^2}", r"\sqrt{\frac{a}{b}}", r"\sqrt[n]{x^n}",
    r"\alpha + \beta = \gamma", r"\Delta x + \Omega", r"\phi \ne \varphi",
    r"a \leq b \geq c \neq d", r"a \approx b \equiv c", r"a \times b \div c \pm d",
    r"a \cdot b", r"x \in A \subset B \cup C \cap D", r"\forall x \exists y",
    r"A \Rightarrow B", r"x \to \infty", r"\partial f / \partial x", r"\nabla \cdot F",
    r"\sin x + \cos y", r"\log_2 n", r"\ln x", r"\det A", r"\exp(x)",
    r"\text{если } x > 0", r"\text{rank}(A)", r"\mathbf{x} + \boldsymbol{\alpha}",
    r"\sum_{i=1}^{n} i^2", r"\prod_{k=1}^{n} k", r"\int_0^1 x^2\,dx", r"\iint_D f\,dA",
    r"\oint_C F", r"\bigcup_{i} A_i", r"{\sum_{i} a_i} + b",
    r"\sum_{i=1}^{n} \frac{1}{i^2} = \frac{\pi^2}{6}",
    r"\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}",
    r"\sum_{\substack{i < j \\ i \in S}} a_{ij}",
    r"\lim_{x \to 0} \frac{\sin x}{x}",
    r"\lim_{n \to \infty} \left(1 + \frac{1}{n}\right)^n",
    r"\left( \frac{a}{b} \right)", r"\left[ x \right]", r"\left\{ x \right\}",
    r"\left| x \right|", r"\left\langle x \right\rangle", r"\left. x \right|",
    r"\left\lfloor x \right\rfloor + \left\lceil y \right\rceil",
    r"\left. \frac{df}{dx} \right|_{x=0}",
    r"\hat{x} + \tilde{y} + \bar{z} + \vec{v}", r"\dot{x}\ddot{x}\check{x}\breve{x}",
    r"\acute{x}\grave{x}", r"\hat{\theta}", r"\overline{AB}", r"\underline{x}",
    r"\underline{\overline{x}}",
    r"\binom{n}{k}", r"{n \choose k}", r"{a \over b}", r"{n \atop k}",
    r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
    r"\begin{bmatrix} 1 & 0 \\ 0 & 1 \end{bmatrix}",
    r"\begin{Bmatrix} a \end{Bmatrix}", r"\begin{vmatrix} a & b \\ c & d \end{vmatrix}",
    r"\begin{Vmatrix} x \end{Vmatrix}", r"\begin{matrix} a & b \\ c & d \end{matrix}",
    r"f(x) = \begin{cases} 1 & x > 0 \\ 0 & x \le 0 \end{cases}",
    r"\begin{array}{lcr} a & b & c \\ d & e & f \end{array}",
    r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}", r"100\%", r"\{ x : x > 0 \}",
    r"\$ \# \& \_",
]


class RoundTripTests(unittest.TestCase):
    def test_corpus_rebuilds_identical_omml(self) -> None:
        self.assertGreaterEqual(len(ROUND_TRIP_CORPUS), 60)
        for source in ROUND_TRIP_CORPUS:
            with self.subTest(source=source):
                original = latex_to_omml(source)
                warnings: list[str] = []
                converted = omml_to_latex(original, warnings)
                self.assertEqual(warnings, [])
                rebuilt = latex_to_omml(converted)
                self.assertEqual(_canonical(rebuilt), _canonical(original),
                                 f"{source!r} came back as {converted!r}")

    def test_multi_line_formulas_round_trip_through_their_rows(self) -> None:
        for source in (r"a &= b + c \\ d &= e", r"x \\ y + 1",
                       r"f(x) &= \sum_i a_i \\ &= 0"):
            with self.subTest(source=source):
                original = latex_to_omml(source)
                rows = equation_rows(original)
                self.assertIsNotNone(rows)
                rebuilt = latex_to_omml(r" \\ ".join(rows))
                self.assertEqual(_canonical(rebuilt), _canonical(original))


if __name__ == "__main__":
    unittest.main()
