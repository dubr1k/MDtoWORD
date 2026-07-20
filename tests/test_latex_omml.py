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
