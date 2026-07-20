import io
import unittest

from docx import Document
from docx.oxml.ns import nsmap as _nsmap
from docx.oxml.ns import qn

from latex_omml import UnsupportedLatexError, latex_to_omml

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

    def test_not_yet_supported_and_reserved_constructs_all_raise(self):
        """Guards Finding 2: deleting `_NOT_YET` entries must not go unnoticed.

        `\\sum` degrading to a lone "Σ" with its limits silently dropped is
        exactly the failure mode this module exists to prevent, so every one
        of these must raise -- naming the offending construct -- rather than
        render as something merely plausible.
        """
        must_raise = [
            (r"\sum", "sum"),
            (r"\prod", "prod"),
            (r"\int", "integral"),
            (r"\oint", "integral"),
            (r"\lim", "limit"),
            (r"\left", "left"),
            (r"\right", "right"),
            (r"\begin", "begin"),
            (r"\end", "end"),
            (r"\binom", "binom"),
            (r"\hat", "hat"),
            (r"\vec", "vec"),
            (r"\overline", "overline"),
            ("\\\\", "line break"),
            ("&", "&"),
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


if __name__ == "__main__":
    unittest.main()
