import io
import unittest

from docx import Document
from docx.oxml.ns import nsmap as _nsmap
from docx.oxml.ns import qn

from latex_omml import _NOT_YET, UnsupportedLatexError, latex_to_omml

# python-docx registers a custom lxml element class per known tag, but it
# knows no `m:` (math) tags, so freshly built <m:...> elements come back as
# bare `lxml.etree._Element` without the `.xml` accessor this test module
# uses for inspection. Registering BaseOxmlElement as the namespace default
# upgrades every `m:` element so `.xml` works below. This is test-only
# convenience -- latex_omml.py itself never calls `.xml`, so the library has
# no business mutating python-docx's global registry to provide it.
try:
    from docx.oxml.parser import element_class_lookup as _element_class_lookup
    from docx.oxml.xmlchemy import BaseOxmlElement as _BaseOxmlElement

    _element_class_lookup.get_namespace(_nsmap["m"])[None] = _BaseOxmlElement
except (ImportError, KeyError, AttributeError):
    pass

MATH_NS = {"m": _nsmap["m"]}


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

    def test_unbraced_frac_argument_splits_multidigit_number(self):
        """`\\frac12x` means `\\frac{1}{2}x`, not `12/x` -- Finding 1."""
        element = latex_to_omml(r"\frac12x")
        children = list(element)
        self.assertEqual(len(children), 2)  # the fraction, then a separate "x" run
        fraction, trailing_run = children
        num = fraction.find("m:num", MATH_NS)
        den = fraction.find("m:den", MATH_NS)
        self.assertEqual([t.text for t in num.iter(qn("m:t"))], ["1"])
        self.assertEqual([t.text for t in den.iter(qn("m:t"))], ["2"])
        self.assertEqual([t.text for t in trailing_run.iter(qn("m:t"))], ["x"])

    def test_unbraced_superscript_splits_multidigit_number(self):
        """`x^12` superscripts only `1`, leaving a literal `2` outside -- Finding 1."""
        element = latex_to_omml("x^12")
        children = list(element)
        self.assertEqual(len(children), 2)  # the sSup, then a separate "2" run
        script, trailing_run = children
        sup = script.find("m:sup", MATH_NS)
        self.assertEqual([t.text for t in sup.iter(qn("m:t"))], ["1"])
        self.assertEqual([t.text for t in trailing_run.iter(qn("m:t"))], ["2"])

    def test_braced_number_arguments_still_take_the_whole_number(self):
        """Braced arguments must be unaffected by the unbraced-splitting fix."""
        frac_element = latex_to_omml(r"\frac{12}{x}")
        num = frac_element.find("m:f/m:num", MATH_NS)
        self.assertEqual([t.text for t in num.iter(qn("m:t"))], ["12"])
        sup_element = latex_to_omml("x^{10}")
        sup = sup_element.find("m:sSup/m:sup", MATH_NS)
        self.assertEqual([t.text for t in sup.iter(qn("m:t"))], ["10"])

    def test_every_not_yet_entry_still_raises(self):
        """Guards Finding 2: deleting `_NOT_YET` entries must not go unnoticed.

        `\\sum` degrading to a lone "Σ" with its limits silently dropped is
        exactly the failure mode this module exists to prevent, so every
        entry still listed as unimplemented must raise -- naming the
        offending construct -- rather than render as something plausible.
        This sweeps the table itself, so an entry that is removed without
        being implemented has nowhere to hide.
        """
        self.assertTrue(_NOT_YET, "the unimplemented table must stay guarded")
        for name in _NOT_YET:
            with self.subTest(command=name):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml("\\" + name)
                self.assertIn(name, str(caught.exception))

    def test_reserved_and_dangling_constructs_raise(self):
        """Constructs that are no longer in `_NOT_YET` but must still fail:
        unknown commands, and the halves of a pair used on their own."""
        must_raise = [
            (r"\qedsymbol", "qedsymbol"),
            (r"\begin{array}{cc} a \end{array}", "array"),
            (r"\begin{aligned} a \end{aligned}", "aligned"),
            (r"\left( x", "left"),
            (r"\right)", "right"),
            (r"\end{pmatrix}", "end"),
            (r"\begin{pmatrix} a & b", "end"),
            (r"\begin{pmatrix} a \end{bmatrix}", "bmatrix"),
            (r"\left\heartsuit x \right.", "heartsuit"),
            ("&", "&"),
            ("\\\\", "line break"),
        ]
        for latex, marker in must_raise:
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml(latex)
                self.assertIn(marker, str(caught.exception))

    def test_scripted_group_keeps_every_run_in_the_base(self):
        """`{a+b}^2` must nest all three runs -- a, +, b -- inside <m:e>,
        not just the last one (Finding 2)."""
        element = latex_to_omml("{a+b}^2")
        base = element.find("m:sSup/m:e", MATH_NS)
        self.assertIsNotNone(base)
        self.assertEqual(len(base), 3)

    def test_spacing_run_survives_docx_save_and_reopen(self):
        """A save/reopen round trip through python-docx's `remove_blank_text`
        parser must not eat the space run from `\\,` (Finding 2)."""
        document = Document()
        document.element.body.insert(0, latex_to_omml(r"a\,b"))
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        reopened = Document(buffer)
        texts = [t.text for t in reopened.element.body.iter(qn("m:t"))]
        self.assertEqual(texts, ["a", " ", "b"])

    def test_mathbf_is_bold_upright_boldsymbol_and_bm_are_bold_italic(self):
        """Finding 3: `\\mathbf` is bold upright; `\\boldsymbol`/`\\bm` are
        bold italic -- the previous implementation had these swapped."""
        mathbf_xml = xml_of(r"\mathbf{x}")
        self.assertIn('<m:sty m:val="b"/>', mathbf_xml)
        self.assertNotIn('<m:sty m:val="bi"/>', mathbf_xml)
        self.assertIn('<m:sty m:val="bi"/>', xml_of(r"\boldsymbol{x}"))
        self.assertIn('<m:sty m:val="bi"/>', xml_of(r"\bm{x}"))

    def test_square_root_with_empty_bracket_hides_degree(self):
        """`\\sqrt[]{x}` must behave like `\\sqrt{x}` -- Finding 5."""
        xml = xml_of(r"\sqrt[]{x}")
        self.assertIn('<m:degHide m:val="1"/>', xml)
        self.assertIn("<m:deg/>", xml)


def _tags(element) -> list:
    """Local names of an element's direct children, in document order."""
    return [child.tag.split("}")[-1] for child in element]


def _texts(element) -> list:
    """Every <m:t> under `element`, in document order."""
    return [t.text for t in element.iter(qn("m:t"))]


class BigConstructTests(unittest.TestCase):
    """Task 3: n-ary operators, limits, delimiters, accents, matrices.

    These assert on element structure rather than substrings: an <m:nary>
    whose limits landed in the body would satisfy every `assertIn` a
    grep-style test can make, yet render wrongly in Word.
    """

    def test_nary_sum_carries_limits(self):
        element = latex_to_omml(r"\sum_{i=1}^{n} i")
        nary = element.find("m:nary", MATH_NS)
        self.assertIsNotNone(nary)
        self.assertEqual(_tags(nary), ["naryPr", "sub", "sup", "e"])
        self.assertEqual(
            nary.find("m:naryPr/m:chr", MATH_NS).get(qn("m:val")), "∑")
        self.assertEqual(
            nary.find("m:naryPr/m:limLoc", MATH_NS).get(qn("m:val")), "undOvr")
        self.assertEqual(_texts(nary.find("m:sub", MATH_NS)), ["i", "=", "1"])
        self.assertEqual(_texts(nary.find("m:sup", MATH_NS)), ["n"])
        self.assertEqual(_texts(nary.find("m:e", MATH_NS)), ["i"])

    def test_nary_without_limits_hides_the_empty_boxes(self):
        nary = latex_to_omml(r"\prod a").find("m:nary", MATH_NS)
        self.assertEqual(
            nary.find("m:naryPr/m:chr", MATH_NS).get(qn("m:val")), "∏")
        self.assertEqual(
            nary.find("m:naryPr/m:subHide", MATH_NS).get(qn("m:val")), "1")
        self.assertEqual(
            nary.find("m:naryPr/m:supHide", MATH_NS).get(qn("m:val")), "1")
        self.assertEqual(_texts(nary.find("m:e", MATH_NS)), ["a"])

    def test_integral_uses_its_own_character_and_side_limits(self):
        nary = latex_to_omml(r"\int_0^1 x\,dx").find("m:nary", MATH_NS)
        self.assertEqual(
            nary.find("m:naryPr/m:chr", MATH_NS).get(qn("m:val")), "∫")
        self.assertEqual(
            nary.find("m:naryPr/m:limLoc", MATH_NS).get(qn("m:val")), "subSup")
        self.assertEqual(_texts(nary.find("m:sub", MATH_NS)), ["0"])
        self.assertEqual(_texts(nary.find("m:sup", MATH_NS)), ["1"])
        self.assertEqual(_texts(nary.find("m:e", MATH_NS)), ["x", " ", "d", "x"])

    def test_every_integral_sign_keeps_limits_beside_it(self):
        """∬, ∭ and ∮ are integrals too -- not just the plain ∫."""
        for latex, character in ((r"\iint x", "∬"), (r"\iiint x", "∭"),
                                 (r"\oint x", "∮")):
            with self.subTest(latex=latex):
                properties = latex_to_omml(latex).find("m:nary/m:naryPr", MATH_NS)
                self.assertEqual(
                    properties.find("m:chr", MATH_NS).get(qn("m:val")), character)
                self.assertEqual(
                    properties.find("m:limLoc", MATH_NS).get(qn("m:val")), "subSup")

    def test_limit_uses_lim_low(self):
        element = latex_to_omml(r"\lim_{x \to 0} f(x)")
        limit = element.find("m:limLow", MATH_NS)
        self.assertIsNotNone(limit)
        self.assertEqual(_tags(limit), ["e", "lim"])
        self.assertEqual(_texts(limit.find("m:e", MATH_NS)), ["lim"])
        self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["x", "→", "0"])
        # `f(x)` is the limit's operand in prose, but it stays a sibling in
        # OMML -- it must not be swallowed into <m:lim>.
        self.assertEqual(
            _texts(element), ["lim", "x", "→", "0", "f", "(", "x", ")"])

    def test_bare_limit_without_a_subscript_is_just_an_upright_run(self):
        element = latex_to_omml(r"\lim f")
        self.assertIsNone(element.find("m:limLow", MATH_NS))
        self.assertEqual(_texts(element), ["lim", "f"])
        self.assertIn('<m:nor m:val="1"/>', element.xml)

    def test_left_right_delimiters_wrap_the_whole_fraction(self):
        element = latex_to_omml(r"\left( \frac{a}{b} \right)")
        delimiter = element.find("m:d", MATH_NS)
        self.assertIsNotNone(delimiter)
        self.assertEqual(
            delimiter.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), ")")
        body = delimiter.find("m:e", MATH_NS)
        self.assertEqual(len(body), 1)
        self.assertEqual(body[0].tag, qn("m:f"))

    def test_left_right_supports_dot_braces_and_bars(self):
        pairs = {
            r"\left[ x \right]": ("[", "]"),
            r"\left\{ x \right\}": ("{", "}"),
            r"\left| x \right|": ("|", "|"),
            r"\left\| x \right\|": ("‖", "‖"),
            r"\left\langle x \right\rangle": ("⟨", "⟩"),
            r"\left. x \right)": ("", ")"),
        }
        for latex, (begin, end) in pairs.items():
            with self.subTest(latex=latex):
                properties = latex_to_omml(latex).find("m:d/m:dPr", MATH_NS)
                self.assertEqual(
                    properties.find("m:begChr", MATH_NS).get(qn("m:val")), begin)
                self.assertEqual(
                    properties.find("m:endChr", MATH_NS).get(qn("m:val")), end)

    def test_nested_left_right_pairs_match_innermost_first(self):
        outer = latex_to_omml(r"\left[ \left( a \right) + b \right]").find(
            "m:d", MATH_NS)
        self.assertEqual(
            outer.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "[")
        body = outer.find("m:e", MATH_NS)
        self.assertEqual([child.tag.split("}")[-1] for child in body],
                         ["d", "r", "r"])
        self.assertEqual(
            body.find("m:d/m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")

    def test_accents_and_overline(self):
        accent = latex_to_omml(r"\hat{x}").find("m:acc", MATH_NS)
        self.assertEqual(_tags(accent), ["accPr", "e"])
        self.assertEqual(
            accent.find("m:accPr/m:chr", MATH_NS).get(qn("m:val")), "̂")
        self.assertEqual(_texts(accent.find("m:e", MATH_NS)), ["x"])

        bar = latex_to_omml(r"\overline{AB}").find("m:bar", MATH_NS)
        self.assertEqual(_tags(bar), ["barPr", "e"])
        self.assertEqual(
            bar.find("m:barPr/m:pos", MATH_NS).get(qn("m:val")), "top")
        self.assertEqual(_texts(bar.find("m:e", MATH_NS)), ["A", "B"])

        self.assertEqual(
            latex_to_omml(r"\underline{x}")
            .find("m:bar/m:barPr/m:pos", MATH_NS).get(qn("m:val")), "bot")
        self.assertIn('<m:chr m:val="⃗"/>', xml_of(r"\vec{v}"))

    def test_accent_still_accepts_its_own_scripts(self):
        """`\\hat{x}^2` superscripts the accented base, not a bare `x`."""
        element = latex_to_omml(r"\hat{x}^2")
        base = element.find("m:sSup/m:e", MATH_NS)
        self.assertEqual(len(base), 1)
        self.assertEqual(base[0].tag, qn("m:acc"))
        self.assertEqual(_texts(element.find("m:sSup/m:sup", MATH_NS)), ["2"])

    def test_binomial_is_a_barless_fraction_in_parentheses(self):
        delimiter = latex_to_omml(r"\binom{n}{k}").find("m:d", MATH_NS)
        self.assertEqual(
            delimiter.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), ")")
        body = delimiter.find("m:e", MATH_NS)
        self.assertEqual(len(body), 1)
        fraction = body[0]
        self.assertEqual(fraction.tag, qn("m:f"))
        self.assertEqual(
            fraction.find("m:fPr/m:type", MATH_NS).get(qn("m:val")), "noBar")
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["n"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["k"])

    def test_pmatrix_builds_a_matrix_in_parentheses(self):
        element = latex_to_omml(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}")
        delimiter = element.find("m:d", MATH_NS)
        self.assertEqual(
            delimiter.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), ")")
        rows = delimiter.findall("m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            [len(row.findall("m:e", MATH_NS)) for row in rows], [2, 2])
        self.assertEqual(
            [_texts(cell) for row in rows
             for cell in row.findall("m:e", MATH_NS)],
            [["a"], ["b"], ["c"], ["d"]])

    def test_matrix_flavours_pick_their_own_fences(self):
        fences = {
            "matrix": None,
            "pmatrix": ("(", ")"),
            "bmatrix": ("[", "]"),
            "Bmatrix": ("{", "}"),
            "vmatrix": ("|", "|"),
            "Vmatrix": ("‖", "‖"),
        }
        for name, pair in fences.items():
            with self.subTest(environment=name):
                element = latex_to_omml(
                    "\\begin{%s} a & b \\\\ c & d \\end{%s}" % (name, name))
                if pair is None:
                    self.assertIsNone(element.find("m:d", MATH_NS))
                    self.assertIsNotNone(element.find("m:m", MATH_NS))
                    continue
                properties = element.find("m:d/m:dPr", MATH_NS)
                self.assertEqual(
                    properties.find("m:begChr", MATH_NS).get(qn("m:val")), pair[0])
                self.assertEqual(
                    properties.find("m:endChr", MATH_NS).get(qn("m:val")), pair[1])

    def test_cases_environment(self):
        element = latex_to_omml(
            r"\begin{cases} x & x > 0 \\ -x & x \leq 0 \end{cases}")
        delimiter = element.find("m:d", MATH_NS)
        self.assertEqual(
            delimiter.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "{")
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), "")
        rows = delimiter.findall("m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            [len(row.findall("m:e", MATH_NS)) for row in rows], [2, 2])
        self.assertEqual(
            [_texts(cell) for cell in rows[1].findall("m:e", MATH_NS)],
            [["-", "x"], ["x", "≤", "0"]])

    def test_ragged_rows_are_padded_so_word_sees_a_rectangle(self):
        rows = latex_to_omml(
            r"\begin{cases} a & b \\ c \end{cases}"
        ).findall("m:d/m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(
            [len(row.findall("m:e", MATH_NS)) for row in rows], [2, 2])
        self.assertEqual(_texts(rows[1].findall("m:e", MATH_NS)[1]), [])

    def test_trailing_row_separator_does_not_add_an_empty_row(self):
        rows = latex_to_omml(
            r"\begin{pmatrix} a & b \\ c & d \\ \end{pmatrix}"
        ).findall("m:d/m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(len(rows), 2)

    def test_nary_body_stops_at_the_enclosing_construct(self):
        """A sum's body runs to the end of its group -- but not past the
        `\\right`, the `&` or the `}` that closes it."""
        delimiter = latex_to_omml(
            r"\left( \sum_{i} a_i \right) + b").find("m:d", MATH_NS)
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), ")")
        self.assertEqual(_texts(delimiter.find("m:e", MATH_NS)), ["i", "a", "i"])

        rows = latex_to_omml(
            r"\begin{pmatrix} \sum_i a_i & b \end{pmatrix}"
        ).findall("m:d/m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(
            [len(row.findall("m:e", MATH_NS)) for row in rows], [2])
        self.assertEqual(_texts(rows[0].findall("m:e", MATH_NS)[1]), ["b"])

        group = latex_to_omml(r"{\sum_i a_i}b")
        self.assertEqual(_tags(group), ["nary", "r"])
        self.assertEqual(_texts(group[1]), ["b"])

    def test_big_constructs_survive_docx_save_and_reopen(self):
        """The shapes must come back intact through a real save/open cycle,
        including python-docx's blank-text-stripping parser."""
        document = Document()
        formulas = [
            r"\sum_{i=1}^{n} i",
            r"\int_0^1 x\,dx",
            r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
            r"\begin{cases} x & x > 0 \\ -x & x \leq 0 \end{cases}",
        ]
        for offset, formula in enumerate(formulas):
            document.element.body.insert(offset, latex_to_omml(formula))
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        body = Document(buffer).element.body

        naries = body.findall("m:oMath/m:nary", MATH_NS)
        self.assertEqual(
            [n.find("m:naryPr/m:chr", MATH_NS).get(qn("m:val")) for n in naries],
            ["∑", "∫"])
        self.assertEqual(_texts(naries[0].find("m:sub", MATH_NS)), ["i", "=", "1"])
        self.assertEqual(_texts(naries[0].find("m:e", MATH_NS)), ["i"])
        self.assertEqual(_texts(naries[1].find("m:e", MATH_NS)),
                         ["x", " ", "d", "x"])

        matrices = body.findall("m:oMath/m:d/m:e/m:m", MATH_NS)
        self.assertEqual(len(matrices), 2)
        self.assertEqual(
            [len(m.findall("m:mr", MATH_NS)) for m in matrices], [2, 2])
        self.assertEqual(_texts(matrices[0]), ["a", "b", "c", "d"])
        self.assertEqual(
            [d.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val"))
             for d in body.findall("m:oMath/m:d", MATH_NS)], ["(", "{"])


if __name__ == "__main__":
    unittest.main()
