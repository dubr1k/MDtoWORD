import io
import unittest

from docx import Document
from docx.oxml.ns import nsmap as _nsmap
from docx.oxml.ns import qn

from mdtoword.latex_omml import (
    _ACCENTS,
    _ENVIRONMENT_ONLY,
    _ESCAPED,
    _LIMIT_OPERATORS,
    _NARY,
    _SPACING,
    _SYMBOLS,
    _UPRIGHT_FUNCTIONS,
    UnsupportedLatexError,
    latex_to_omml,
    split_equation_tag,
)

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

# The exact set `_ENVIRONMENT_ONLY` is expected to have right now, written
# out by hand -- not derived from `_ENVIRONMENT_ONLY` at runtime, which
# would recreate the very self-referential hole the sweep test below would
# otherwise have on its own (deleting an entry here has to be a deliberate
# second edit).
EXPECTED_ENVIRONMENT_ONLY = (
    "matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix",
    "array", "cases",
)


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

    def test_every_environment_only_name_still_raises_as_a_bare_command(self):
        """Guards Finding 2: deleting `_ENVIRONMENT_ONLY` entries must not
        go unnoticed.

        `\\sum` degrading to a lone "Σ" with its limits silently dropped is
        exactly the failure mode this module exists to prevent, so every
        name that only works as an environment must raise -- naming the
        environment form to use instead -- rather than render as something
        plausible.  This sweeps the table itself, so an entry that is
        removed without being implemented has nowhere to hide.
        """
        self.assertTrue(
            _ENVIRONMENT_ONLY, "the environment-only table must stay guarded")
        # Deleting an entry from `_ENVIRONMENT_ONLY` above is not enough on
        # its own -- this table is iterated below, so the sweep can't notice
        # a name that is simply gone. Pinning the expected names against a
        # frozen literal forces a deliberate edit here too.
        self.assertEqual(
            sorted(_ENVIRONMENT_ONLY), sorted(EXPECTED_ENVIRONMENT_ONLY))
        for name in _ENVIRONMENT_ONLY:
            with self.subTest(command=name):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml("\\" + name)
                message = str(caught.exception)
                self.assertIn(name, message)
                # Not just the command name: the generic fallback message
                # also contains that (`Unsupported LaTeX command: \foo`), so
                # asserting only the name would not notice the
                # `_ENVIRONMENT_ONLY` enforcement block itself being
                # deleted. The `\begin{...}` suggestion only appears in the
                # dedicated `_ENVIRONMENT_ONLY` message.
                self.assertIn("\\begin{%s}" % name, message)

    def test_environment_only_names_are_disjoint_from_every_fallback_table(self):
        """`_ENVIRONMENT_ONLY` is checked before several branches for
        constructs this module already implements reach their own fallback
        tables -- not merely "before the symbol table" as the old comment in
        `_parse_command` claimed. If a name were ever added to
        `_ENVIRONMENT_ONLY` that also lived in one of those tables,
        whichever branch runs first would shadow the other, the same way a
        deleted entry used to fall through unnoticed (Finding 1). This pins
        the seven tables `_ENVIRONMENT_ONLY` must stay disjoint from."""
        fallback_tables = {
            "_NARY": _NARY,
            "_ACCENTS": _ACCENTS,
            "_LIMIT_OPERATORS": _LIMIT_OPERATORS,
            "_SYMBOLS": _SYMBOLS,
            "_SPACING": _SPACING,
            "_ESCAPED": _ESCAPED,
            "_UPRIGHT_FUNCTIONS": _UPRIGHT_FUNCTIONS,
        }
        for table_name, table in fallback_tables.items():
            with self.subTest(table=table_name):
                self.assertEqual(set(_ENVIRONMENT_ONLY) & set(table), set())

    def test_reserved_and_dangling_constructs_raise(self):
        """Constructs outside `_ENVIRONMENT_ONLY` that must still fail:
        unknown commands, and the halves of a pair used on their own."""
        must_raise = [
            (r"\qedsymbol", "qedsymbol"),
            # `aligned` used to be the example here; it is supported now, so
            # an environment that never will be stands in for it.
            (r"\begin{tikzcd} a \end{tikzcd}", "tikzcd"),
            (r"\left( x", "left"),
            (r"\right)", "right"),
            (r"\end{pmatrix}", "end"),
            (r"\begin{pmatrix} a & b", "end"),
            (r"\begin{pmatrix} a \end{bmatrix}", "bmatrix"),
            (r"\left\heartsuit x \right.", "heartsuit"),
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

    def test_no_run_combines_nor_with_sty(self):
        """OOXML's `CT_RPR` makes `<m:nor>` and `<m:sty>` a *choice*, not a
        sequence, so a run may carry one or the other but never both.

        `\\mathbf{x}` -- upright *and* bold -- is the case that hits it:
        emitting `<m:nor/><m:sty m:val="b"/>` is rejected by the ISO/IEC
        29500-4 schema. `<m:sty m:val="b">` already means bold *upright*, so
        the style value alone carries both facts and `<m:nor>` is redundant
        there. Swept over every construct that can produce a styled run.
        """
        formulas = [
            r"\mathbf{x}", r"\mathbf{ab}", r"\mathbf{2}",
            r"\mathbf{\text{ab}}", r"\mathbf{\alpha}",
            r"\text{hello}", r"\mathrm{d}", r"\operatorname{sgn}",
            r"\boldsymbol{v}", r"\bm{w}", r"\mathit{y}",
            r"\mathbf{\frac{a}{b}}", r"\mathbf{\sin x}",
        ]
        for latex in formulas:
            with self.subTest(latex=latex):
                for properties in latex_to_omml(latex).iter(qn("m:rPr")):
                    present = _tags(properties)
                    self.assertFalse(
                        "nor" in present and "sty" in present,
                        f"<m:rPr> carries both nor and sty: {present}",
                    )


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

    def test_limsup_and_liminf_produce_lim_low_with_the_right_text(self):
        """`\\limsup`/`\\liminf` left the unimplemented table and were
        mapping to the upright run text "lim sup"/"lim inf" -- new behaviour
        that gained zero coverage, the exact pattern these guards exist to
        catch. `\\limsup_{n} a` must produce a `limLow` whose `<m:e>` holds
        the text "lim sup", not a bare symbol with its limit dropped."""
        element = latex_to_omml(r"\limsup_{n} a")
        limit = element.find("m:limLow", MATH_NS)
        self.assertIsNotNone(limit)
        self.assertEqual(_tags(limit), ["e", "lim"])
        self.assertEqual(_texts(limit.find("m:e", MATH_NS)), ["lim sup"])
        self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["n"])
        self.assertEqual(_texts(element), ["lim sup", "n", "a"])

        liminf_limit = latex_to_omml(r"\liminf_{n} a").find(
            "m:limLow", MATH_NS)
        self.assertIsNotNone(liminf_limit)
        self.assertEqual(_texts(liminf_limit.find("m:e", MATH_NS)), ["lim inf"])

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

    def test_every_nary_and_accent_entry_converts_without_raising(self):
        """13 of these -- \\coprod, \\bigoplus, \\bigotimes, \\bigvee,
        \\bigwedge, \\widehat, \\widetilde, \\dot, \\ddot, \\acute, \\grave,
        \\check, \\breve -- had zero coverage: new-behaviour-with-no-test is
        exactly the pattern the `_ENVIRONMENT_ONLY` guards exist to catch,
        and these tables are no different just because they are already
        implemented."""
        for name, character in _NARY.items():
            with self.subTest(command=name):
                nary = latex_to_omml("\\%s x" % name).find("m:nary", MATH_NS)
                self.assertIsNotNone(nary)
                self.assertEqual(
                    nary.find("m:naryPr/m:chr", MATH_NS).get(qn("m:val")),
                    character)

        for name, character in _ACCENTS.items():
            with self.subTest(command=name):
                accent = latex_to_omml("\\%s{x}" % name).find("m:acc", MATH_NS)
                self.assertIsNotNone(accent)
                self.assertEqual(
                    accent.find("m:accPr/m:chr", MATH_NS).get(qn("m:val")),
                    character)


class InfixCommandTests(unittest.TestCase):
    r"""``\over``, ``\atop`` and ``\choose``: TeX's infix fraction builders.

    Each one splits the group it appears in -- everything to its left is the
    numerator, everything to its right the denominator -- which is why they
    cannot be handled in ``_parse_command`` like a prefix command.
    """

    def test_over_builds_a_plain_fraction(self):
        fraction = latex_to_omml(r"{a + b \over c}").find("m:f", MATH_NS)
        self.assertIsNotNone(fraction)
        self.assertIsNone(fraction.find("m:fPr/m:type", MATH_NS))
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["a", "+", "b"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["c"])

    def test_atop_builds_a_barless_fraction(self):
        fraction = latex_to_omml(r"{n \atop k}").find("m:f", MATH_NS)
        self.assertEqual(
            fraction.find("m:fPr/m:type", MATH_NS).get(qn("m:val")), "noBar")
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["n"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["k"])
        self.assertIsNone(
            latex_to_omml(r"{n \atop k}").find("m:d", MATH_NS),
            r"\atop stacks without fences -- those belong to \choose",
        )

    def test_choose_is_a_barless_fraction_in_parentheses(self):
        r"""``{n \choose k}`` must match ``\binom{n}{k}`` exactly."""
        delimiter = latex_to_omml(r"{n \choose k}").find("m:d", MATH_NS)
        self.assertEqual(
            delimiter.find("m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")
        self.assertEqual(
            delimiter.find("m:dPr/m:endChr", MATH_NS).get(qn("m:val")), ")")
        fraction = delimiter.find("m:e/m:f", MATH_NS)
        self.assertEqual(
            fraction.find("m:fPr/m:type", MATH_NS).get(qn("m:val")), "noBar")
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["n"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["k"])

    def test_infix_without_braces_splits_the_whole_formula(self):
        fraction = latex_to_omml(r"1 \over 2").find("m:f", MATH_NS)
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["1"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["2"])

    def test_infix_stops_at_the_group_that_encloses_it(self):
        r"""``x + {a \over b} + y`` divides only inside the braces."""
        element = latex_to_omml(r"x + {a \over b} + y")
        self.assertEqual(
            _tags(element), ["r", "r", "f", "r", "r"])
        fraction = element.find("m:f", MATH_NS)
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["a"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["b"])

    def test_infix_inside_a_matrix_cell_stops_at_the_cell(self):
        rows = latex_to_omml(
            r"\begin{pmatrix} a \over b & c \end{pmatrix}"
        ).findall("m:d/m:e/m:m/m:mr", MATH_NS)
        cells = rows[0].findall("m:e", MATH_NS)
        self.assertEqual(len(cells), 2)
        self.assertEqual(_tags(cells[0]), ["f"])
        self.assertEqual(_texts(cells[1]), ["c"])

    def test_two_infix_commands_in_one_group_raise(self):
        r"""TeX itself rejects ``a \over b \over c`` as ambiguous; guessing
        an association here would silently produce one of two readings."""
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"a \over b \over c")
        self.assertIn("over", str(caught.exception))

    def test_nesting_infix_through_braces_is_allowed(self):
        r"""``{{a \over b} \over c}`` is unambiguous and must convert."""
        outer = latex_to_omml(r"{{a \over b} \over c}").find("m:f", MATH_NS)
        self.assertEqual(_tags(outer.find("m:num", MATH_NS)), ["f"])
        self.assertEqual(_texts(outer.find("m:den", MATH_NS)), ["c"])

    def test_infix_with_nothing_after_it_still_builds_an_empty_denominator(self):
        fraction = latex_to_omml(r"a \over").find("m:f", MATH_NS)
        self.assertEqual(_texts(fraction.find("m:num", MATH_NS)), ["a"])
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), [])

    def test_infix_used_as_an_argument_reports_its_two_sides(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\frac\over x")
        self.assertIn("over", str(caught.exception))


class LineBreakTests(unittest.TestCase):
    r"""``\\`` outside a matrix: an OMML equation array (``<m:eqArr>``)."""

    def test_line_break_builds_an_equation_array(self):
        element = latex_to_omml(r"a = b \\ c = d")
        array = element.find("m:eqArr", MATH_NS)
        self.assertIsNotNone(array)
        lines = array.findall("m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        self.assertEqual(_texts(lines[0]), ["a", "=", "b"])
        self.assertEqual(_texts(lines[1]), ["c", "=", "d"])

    def test_a_formula_without_a_line_break_stays_flat(self):
        self.assertIsNone(latex_to_omml("a = b").find("m:eqArr", MATH_NS))

    def test_trailing_line_break_does_not_add_an_empty_line(self):
        element = latex_to_omml(r"a \\ b \\")
        lines = element.find("m:eqArr", MATH_NS).findall("m:e", MATH_NS)
        self.assertEqual(len(lines), 2)

    def test_a_lone_line_break_produces_nothing_rather_than_raising(self):
        r"""``\\`` on its own is an empty line, not an error: the whole
        point of implementing it is that it no longer fails loudly."""
        element = latex_to_omml("\\\\")
        self.assertEqual(list(element), [])

    def test_line_breaks_nest_inside_a_group(self):
        element = latex_to_omml(r"\left\{ a \\ b \right.")
        lines = element.findall("m:d/m:e/m:eqArr/m:e", MATH_NS)
        self.assertEqual([_texts(line) for line in lines], [["a"], ["b"]])

    def test_a_line_break_ends_an_n_ary_operator_body(self):
        r"""``\sum_i a_i \\ b`` must leave ``b`` on the next line, not
        swallow it into the sum's operand."""
        element = latex_to_omml(r"\sum_i a_i \\ b")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        self.assertEqual(_tags(lines[0]), ["nary"])
        self.assertEqual(_texts(lines[0].find("m:nary/m:e", MATH_NS)), ["a", "i"])
        self.assertEqual(_texts(lines[1]), ["b"])

    def test_matrix_rows_are_still_rows_not_an_equation_array(self):
        element = latex_to_omml(r"\begin{pmatrix} a \\ b \end{pmatrix}")
        self.assertIsNone(element.find(".//m:eqArr", MATH_NS))
        self.assertEqual(len(element.findall("m:d/m:e/m:m/m:mr", MATH_NS)), 2)


class AlignmentPointTests(unittest.TestCase):
    r"""``&`` between the lines of a multi-line formula: an OMML alignment
    point, spelled ``<m:aln/>`` inside the ``<m:rPr>`` of the run that
    starts the aligned segment -- which is how Word itself writes it.

    This is what lets ``align`` stay ONE Word equation with its ``=`` signs
    lined up, instead of being cut into one centred paragraph per line.
    """

    @staticmethod
    def _aligned_runs(line):
        """Every run in `line` carrying an alignment point, with its text."""
        return [
            "".join(t.text or "" for t in run.iter(qn("m:t")))
            for run in line.iter(qn("m:r"))
            if run.find("m:rPr/m:aln", MATH_NS) is not None
        ]

    def test_ampersand_marks_the_run_that_follows_it(self):
        r"""In ``a &= b`` the ``&`` sits before ``=``, so the ``=`` run is
        the alignment point -- exactly what lines the equals signs up."""
        element = latex_to_omml(r"a &= b \\ c &= d")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        for line in lines:
            self.assertEqual(self._aligned_runs(line), ["="])
        self.assertEqual(
            [_texts(line) for line in lines],
            [["a", "=", "b"], ["c", "=", "d"]])

    def test_alignment_point_does_not_add_or_drop_any_text(self):
        r"""The marker is formatting on an existing run, not a new glyph."""
        self.assertEqual(
            _texts(latex_to_omml(r"a &= b \\ c &= d")),
            _texts(latex_to_omml(r"a = b \\ c = d")))

    def test_several_alignment_points_on_one_line(self):
        r"""``eqnarray``'s ``a &=& b`` has two, and both must survive."""
        line = latex_to_omml(r"a &=& b \\ c &=& d").findall(
            "m:eqArr/m:e", MATH_NS)[0]
        self.assertEqual(self._aligned_runs(line), ["=", "b"])

    def test_alignment_point_before_a_non_run_gets_its_own_marker(self):
        r"""``a &\frac{1}{2}`` puts the point before a fraction, which has
        no ``<m:rPr>`` to carry it, so an empty marker run stands in."""
        line = latex_to_omml(
            r"a &\frac{1}{2} \\ b &\frac{3}{4}"
        ).findall("m:eqArr/m:e", MATH_NS)[0]
        self.assertEqual(_tags(line), ["r", "r", "f"])
        self.assertEqual(self._aligned_runs(line), [""])
        # The marker's own <m:t> is empty, so it contributes no glyph.
        self.assertEqual(_texts(line), ["a", "", "1", "2"])

    def test_alignment_point_at_the_end_of_a_line_is_still_recorded(self):
        line = latex_to_omml(r"a & \\ b & c").findall(
            "m:eqArr/m:e", MATH_NS)[0]
        self.assertEqual(self._aligned_runs(line), [""])

    def test_an_alignment_point_ends_an_n_ary_operator_body(self):
        r"""``\sum_i a_i &= b`` must leave ``a_i`` as the sum's operand and
        let ``&`` start the next aligned segment.

        Without this the sum swallows the ``&`` into its own body, where --
        having no line break of its own to pair with -- it is rejected as a
        stray ampersand and the whole formula fails to convert. Found by
        validating against the OOXML schema, not by an earlier test.
        """
        element = latex_to_omml(r"\sum_i a_i &= b \\ c &= d")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        self.assertEqual(_tags(lines[0]), ["nary", "r", "r"])
        nary = lines[0].find("m:nary", MATH_NS)
        self.assertEqual(_texts(nary.find("m:e", MATH_NS)), ["a", "i"])
        self.assertEqual(self._aligned_runs(lines[0]), ["="])

    def test_an_alignment_point_ends_an_infix_denominator(self):
        r"""Same rule for ``\over``: ``{a \over b} &= c`` style input must
        not pull the ``&`` into the denominator."""
        element = latex_to_omml(r"a \over b &= c \\ d &= e")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        fraction = lines[0].find("m:f", MATH_NS)
        self.assertEqual(_texts(fraction.find("m:den", MATH_NS)), ["b"])
        self.assertEqual(self._aligned_runs(lines[0]), ["="])

    def test_stray_ampersand_in_a_single_line_formula_still_raises(self):
        r"""Without a ``\\`` there is nothing to align against, so ``&`` is
        far more likely a literal ampersand the user forgot to escape.
        Keeping the loud error is what stops ``Tom & Jerry`` inside ``$…$``
        from turning into an invisible alignment marker."""
        for latex in ("a & b", "&", r"\text{x} & y"):
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml(latex)
                self.assertIn("&", str(caught.exception))

    def test_matrix_ampersands_stay_column_separators(self):
        r"""A matrix consumes ``&`` as a cell break long before the
        alignment logic sees it -- no marker may leak into one."""
        element = latex_to_omml(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}")
        self.assertIsNone(element.find(".//m:aln", MATH_NS))
        self.assertEqual(
            len(element.findall("m:d/m:e/m:m/m:mr/m:e", MATH_NS)), 4)

    def test_array_ampersands_stay_column_separators(self):
        element = latex_to_omml(
            r"\begin{array}{cc} a & b \\ c & d \end{array}")
        self.assertIsNone(element.find(".//m:aln", MATH_NS))

    def test_alignment_point_survives_a_docx_save_and_reopen(self):
        document = Document()
        document.element.body.insert(0, latex_to_omml(r"a &= b \\ c &= d"))
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        lines = Document(buffer).element.body.findall(
            "m:oMath/m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        for line in lines:
            self.assertEqual(self._aligned_runs(line), ["="])


class ArrayAndSubstackTests(unittest.TestCase):
    r"""``\begin{array}{...}`` and ``\substack{...}``."""

    def test_array_builds_an_unfenced_matrix(self):
        element = latex_to_omml(
            r"\begin{array}{cc} a & b \\ c & d \end{array}")
        self.assertIsNone(element.find("m:d", MATH_NS))
        rows = element.findall("m:m/m:mr", MATH_NS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            [_texts(cell) for row in rows
             for cell in row.findall("m:e", MATH_NS)],
            [["a"], ["b"], ["c"], ["d"]])

    def test_array_column_specification_sets_each_column_justification(self):
        matrix = latex_to_omml(
            r"\begin{array}{lcr} a & b & c \end{array}").find("m:m", MATH_NS)
        columns = matrix.findall("m:mPr/m:mcs/m:mc", MATH_NS)
        self.assertEqual(
            [c.find("m:mcPr/m:mcJc", MATH_NS).get(qn("m:val")) for c in columns],
            ["left", "center", "right"])
        self.assertEqual(
            [c.find("m:mcPr/m:count", MATH_NS).get(qn("m:val")) for c in columns],
            ["1", "1", "1"])

    def test_array_properties_come_before_the_rows(self):
        """OOXML requires <m:mPr> first; Word rejects the file otherwise."""
        matrix = latex_to_omml(
            r"\begin{array}{c} a \end{array}").find("m:m", MATH_NS)
        self.assertEqual(_tags(matrix)[0], "mPr")

    def test_array_fenced_by_left_right_keeps_both(self):
        element = latex_to_omml(
            r"\left( \begin{array}{cc} a & b \end{array} \right)")
        self.assertEqual(
            element.find("m:d/m:dPr/m:begChr", MATH_NS).get(qn("m:val")), "(")
        self.assertIsNotNone(element.find("m:d/m:e/m:m", MATH_NS))

    def test_array_without_a_column_specification_says_so(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{array} a \end{array}")
        self.assertIn("column specification", str(caught.exception))

    def test_array_vertical_rule_is_rejected_rather_than_dropped(self):
        r"""OMML has no vertical rule inside a matrix, so ``{c|c}`` cannot
        be honoured -- silently dropping the rule would turn an augmented
        matrix into an ordinary one."""
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{array}{c|c} a & b \end{array}")
        message = str(caught.exception)
        self.assertIn("|", message)
        self.assertIn("array", message)

    def test_array_with_more_columns_than_declared_is_rejected(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{array}{c} a & b \end{array}")
        self.assertIn("column", str(caught.exception))

    def test_array_padded_column_from_the_specification_stays_empty(self):
        rows = latex_to_omml(
            r"\begin{array}{ccc} a & b \end{array}").findall("m:m/m:mr", MATH_NS)
        cells = rows[0].findall("m:e", MATH_NS)
        self.assertEqual(len(cells), 3)
        self.assertEqual(_texts(cells[2]), [])

    def test_substack_stacks_its_lines_in_a_single_column(self):
        matrix = latex_to_omml(
            r"\substack{i < j \\ i \in S}").find("m:m", MATH_NS)
        self.assertIsNotNone(matrix)
        rows = matrix.findall("m:mr", MATH_NS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            [len(row.findall("m:e", MATH_NS)) for row in rows], [1, 1])
        self.assertEqual(_texts(rows[0]), ["i", "<", "j"])
        self.assertEqual(_texts(rows[1]), ["i", "∈", "S"])

    def test_substack_serves_as_an_n_ary_limit(self):
        r"""Its whole reason to exist: ``\sum_{\substack{...}}``."""
        nary = latex_to_omml(
            r"\sum_{\substack{i = 1 \\ i \neq j}} a_i").find("m:nary", MATH_NS)
        self.assertEqual(
            len(nary.findall("m:sub/m:m/m:mr", MATH_NS)), 2)
        self.assertEqual(_texts(nary.find("m:e", MATH_NS)), ["a", "i"])

    def test_substack_without_a_brace_group_says_so(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\substack x")
        self.assertIn("substack", str(caught.exception))

    def test_new_constructs_survive_a_docx_save_and_reopen(self):
        """The shapes must come back intact through a real save/open cycle,
        including python-docx's blank-text-stripping parser."""
        document = Document()
        formulas = [
            r"\begin{array}{lc} a & b \\ c & d \end{array}",
            r"{n \choose k}",
            r"x = 1 \\ y = 2",
            r"\sum_{\substack{i \\ j}} a",
        ]
        for offset, formula in enumerate(formulas):
            document.element.body.insert(offset, latex_to_omml(formula))
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        body = Document(buffer).element.body

        array = body.find("m:oMath/m:m", MATH_NS)
        self.assertEqual(
            [c.find("m:mcPr/m:mcJc", MATH_NS).get(qn("m:val"))
             for c in array.findall("m:mPr/m:mcs/m:mc", MATH_NS)],
            ["left", "center"])
        self.assertEqual(_texts(array), ["a", "b", "c", "d"])

        binomial = body.find("m:oMath/m:d/m:e/m:f", MATH_NS)
        self.assertEqual(
            binomial.find("m:fPr/m:type", MATH_NS).get(qn("m:val")), "noBar")

        lines = body.findall("m:oMath/m:eqArr/m:e", MATH_NS)
        self.assertEqual(
            [_texts(line) for line in lines], [["x", "=", "1"], ["y", "=", "2"]])

        self.assertEqual(
            len(body.findall("m:oMath/m:nary/m:sub/m:m/m:mr", MATH_NS)), 2)


def _val(element, path):
    """The `m:val` of the element at `path` under `element`."""
    found = element.find(path, MATH_NS)
    if found is None:
        raise AssertionError(f"{path} not found")
    return found.get(qn("m:val"))


def _aligned_runs(line):
    """Every run in `line` carrying an alignment point, with its text."""
    return [
        "".join(t.text or "" for t in run.iter(qn("m:t")))
        for run in line.iter(qn("m:r"))
        if run.find("m:rPr/m:aln", MATH_NS) is not None
    ]


def _run_properties(run):
    """`{local-name: m:val}` of a run's <m:rPr> children."""
    properties = run.find("m:rPr", MATH_NS)
    if properties is None:
        return {}
    return {child.tag.split("}")[-1]: child.get(qn("m:val"))
            for child in properties}


def _word_tags(run):
    """Local names of the <w:rPr> children of an <m:r>."""
    properties = run.find(qn("w:rPr"))
    return [] if properties is None else _tags(properties)


def _runs(element):
    return list(element.iter(qn("m:r")))


class EnvironmentTests(unittest.TestCase):
    r"""``aligned``, ``gathered``, ``split``, ``alignedat``, the top-level
    amsmath environments, ``smallmatrix`` and the ``cases`` family."""

    def test_aligned_is_one_equation_array_aligned_on_its_ampersands(self):
        r"""The single most common LLM shape: ``$$\begin{aligned} a &= b \\
        c &= d \end{aligned}$$`` must be ONE equation array whose ``=``
        signs are the alignment points -- not a refusal, and not two
        separate equations."""
        element = latex_to_omml(
            r"\begin{aligned} a &= b \\ c &= d \end{aligned}")
        self.assertEqual(_tags(element), ["eqArr"])
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual([_texts(line) for line in lines],
                         [["a", "=", "b"], ["c", "=", "d"]])
        for line in lines:
            self.assertEqual(_aligned_runs(line), ["="])

    def test_aligned_matches_the_bare_multi_line_formula(self):
        r"""``aligned`` means exactly what a bare ``a &= b \\ c &= d``
        already meant, so the two must produce the same XML."""
        self.assertEqual(
            xml_of(r"\begin{aligned} a &= b \\ c &= d \end{aligned}"),
            xml_of(r"a &= b \\ c &= d"))

    def test_split_and_alignat_family_align_the_same_way(self):
        for latex in (
            r"\begin{split} a &= b \\ c &= d \end{split}",
            r"\begin{align*} a &= b \\ c &= d \end{align*}",
            r"\begin{align} a &= b \\ c &= d \end{align}",
            r"\begin{flalign} a &= b \\ c &= d \end{flalign}",
            r"\begin{eqnarray*} a &= b \\ c &= d \end{eqnarray*}",
            r"\begin{alignat}{1} a &= b \\ c &= d \end{alignat}",
            r"\begin{alignedat}{1} a &= b \\ c &= d \end{alignedat}",
        ):
            with self.subTest(latex=latex):
                lines = latex_to_omml(latex).findall("m:eqArr/m:e", MATH_NS)
                self.assertEqual(len(lines), 2)
                for line in lines:
                    self.assertEqual(_aligned_runs(line), ["="])

    def test_alignedat_strips_its_column_count_and_keeps_every_point(self):
        r"""``{2}`` is TeX's column-pair count, not content: it must not
        show up as a "2" in the first line."""
        element = latex_to_omml(
            r"\begin{alignedat}{2} a &= b &\quad c &= d \\ "
            r"e &= f &\quad g &= h \end{alignedat}")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        self.assertEqual(_texts(lines[0])[0], "a")
        self.assertNotIn("2", _texts(lines[0]))
        self.assertEqual(len(_aligned_runs(lines[0])), 3)

    def test_aligned_takes_an_optional_vertical_position(self):
        element = latex_to_omml(
            r"\begin{aligned}[t] a &= b \\ c &= d \end{aligned}")
        self.assertEqual(
            _texts(element.findall("m:eqArr/m:e", MATH_NS)[0]), ["a", "=", "b"])
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{aligned}[x] a &= b \end{aligned}")
        self.assertIn("aligned", str(caught.exception))

    def test_single_line_aligned_needs_no_array_and_no_marker(self):
        r"""Inside ``aligned`` the ``&`` is unambiguous, so one line with
        one is not the "literal ampersand" error -- it just has nothing to
        align against."""
        element = latex_to_omml(r"\begin{aligned} a &= b \end{aligned}")
        self.assertEqual(_texts(element), ["a", "=", "b"])
        self.assertIsNone(element.find(".//m:eqArr", MATH_NS))
        self.assertIsNone(element.find(".//m:aln", MATH_NS))

    def test_trailing_line_break_and_its_spacing_argument_are_dropped(self):
        r"""``\\[4pt]`` is vertical spacing, not a bracketed "4pt"; a
        final ``\\`` adds no empty line."""
        element = latex_to_omml(
            r"\begin{aligned} a &= b \\[4pt] c &= d \\ \end{aligned}")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual([_texts(line) for line in lines],
                         [["a", "=", "b"], ["c", "=", "d"]])

    def test_a_bracket_that_is_not_a_length_stays_content(self):
        lines = latex_to_omml(r"a \\ [b, c]").findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(_texts(lines[1]), ["[", "b", ",", "c", "]"])

    def test_gathered_stacks_lines_without_alignment(self):
        element = latex_to_omml(
            r"\begin{gathered} a = b \\ c + d = e \end{gathered}")
        lines = element.findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        self.assertIsNone(element.find(".//m:aln", MATH_NS))
        for latex in (r"\begin{gather*} a \\ b \end{gather*}",
                      r"\begin{multline} a \\ b \end{multline}"):
            with self.subTest(latex=latex):
                self.assertEqual(
                    len(latex_to_omml(latex).findall("m:eqArr/m:e", MATH_NS)), 2)

    def test_ampersand_inside_gathered_is_refused_naming_it(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{gathered} a &= b \\ c \end{gathered}")
        self.assertIn("gathered", str(caught.exception))

    def test_equation_environment_is_just_its_content(self):
        self.assertEqual(
            xml_of(r"\begin{equation*} E = mc^2 \end{equation*}"),
            xml_of(r"E = mc^2"))

    def test_aligned_nests_inside_a_brace(self):
        element = latex_to_omml(
            r"\left\{ \begin{aligned} x &= 1 \\ y &= 2 \end{aligned} \right.")
        self.assertEqual(_val(element, "m:d/m:dPr/m:begChr"), "{")
        lines = element.findall("m:d/m:e/m:eqArr/m:e", MATH_NS)
        self.assertEqual(len(lines), 2)
        for line in lines:
            self.assertEqual(_aligned_runs(line), ["="])

    def test_aligned_nests_inside_a_matrix_cell(self):
        r"""The inner environment owns its own ``&`` and ``\\``; the
        matrix must still see exactly two cells."""
        rows = latex_to_omml(
            r"\begin{pmatrix} \begin{aligned} a &= b \\ c &= d \end{aligned}"
            r" & x \end{pmatrix}").findall("m:d/m:e/m:m/m:mr", MATH_NS)
        self.assertEqual(len(rows), 1)
        cells = rows[0].findall("m:e", MATH_NS)
        self.assertEqual(len(cells), 2)
        self.assertEqual(_tags(cells[0]), ["eqArr"])
        self.assertEqual(_texts(cells[1]), ["x"])

    def test_mismatched_end_is_refused(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\begin{aligned} a &= b \end{gathered}")
        self.assertIn("gathered", str(caught.exception))

    def test_smallmatrix_is_a_plain_matrix(self):
        element = latex_to_omml(
            r"\begin{smallmatrix} a & b \\ c & d \end{smallmatrix}")
        self.assertEqual(_tags(element), ["m"])
        self.assertEqual(len(element.findall("m:m/m:mr", MATH_NS)), 2)
        self.assertEqual(
            _val(latex_to_omml(
                r"\begin{psmallmatrix} a \end{psmallmatrix}"),
                "m:d/m:dPr/m:begChr"), "(")

    def test_starred_matrix_takes_one_column_alignment(self):
        matrix = latex_to_omml(
            r"\begin{pmatrix*}[r] -1 & 2 \\ 3 & -4 \end{pmatrix*}"
        ).find("m:d/m:e/m:m", MATH_NS)
        self.assertEqual(
            [_val(column, "m:mcPr/m:mcJc")
             for column in matrix.findall("m:mPr/m:mcs/m:mc", MATH_NS)],
            ["right", "right"])

    def test_cases_family_fences_and_left_aligned_columns(self):
        r"""``cases`` columns are left-aligned in LaTeX (``{ll}``), not
        centred; ``rcases`` moves the brace to the right."""
        fences = {
            "cases": ("{", ""), "dcases": ("{", ""),
            "rcases": ("", "}"), "drcases": ("", "}"),
        }
        for name, (begin, end) in fences.items():
            with self.subTest(environment=name):
                element = latex_to_omml(
                    "\\begin{%s} x & x > 0 \\\\ -x & x \\le 0 \\end{%s}"
                    % (name, name))
                self.assertEqual(_val(element, "m:d/m:dPr/m:begChr"), begin)
                self.assertEqual(_val(element, "m:d/m:dPr/m:endChr"), end)
                matrix = element.find("m:d/m:e/m:m", MATH_NS)
                self.assertEqual(
                    [_val(column, "m:mcPr/m:mcJc")
                     for column in matrix.findall("m:mPr/m:mcs/m:mc", MATH_NS)],
                    ["left", "left"])
                self.assertEqual(len(matrix.findall("m:mr", MATH_NS)), 2)

    def test_starred_cases_set_the_condition_column_as_text(self):
        r"""``cases*``: everything after ``&`` is text, with ``$...$`` for
        math -- "if" must stay a word, not two italic variables."""
        rows = latex_to_omml(
            r"\begin{cases*} 1 & if $x > 0$ \\ 0 & otherwise \end{cases*}"
        ).findall("m:d/m:e/m:m/m:mr", MATH_NS)
        condition = rows[0].findall("m:e", MATH_NS)[1]
        runs = _runs(condition)
        self.assertEqual(_texts(condition), ["if ", "x", ">", "0"])
        self.assertEqual(_run_properties(runs[0]), {"nor": "1"})
        self.assertEqual(_run_properties(runs[1]), {"sty": "i"})
        self.assertEqual(
            _texts(rows[1].findall("m:e", MATH_NS)[1]), ["otherwise"])

    def test_subarray_is_an_array(self):
        matrix = latex_to_omml(
            r"\begin{subarray}{l} i < j \\ k \end{subarray}").find("m:m", MATH_NS)
        self.assertEqual(len(matrix.findall("m:mr", MATH_NS)), 2)
        self.assertEqual(_val(matrix, "m:mPr/m:mcs/m:mc/m:mcPr/m:mcJc"), "left")


class MathAlphabetTests(unittest.TestCase):
    r"""``\mathbb`` and friends through ``<m:scr>``, ``\mathrm`` as upright
    math, and the ``\text..`` family."""

    def test_script_alphabets_set_scr_and_an_upright_style(self):
        alphabets = {
            "mathbb": "double-struck", "Bbb": "double-struck",
            "mathbbm": "double-struck", "mathcal": "script",
            "mathscr": "script", "mathfrak": "fraktur",
            "mathsf": "sans-serif", "mathtt": "monospace",
        }
        for name, script in alphabets.items():
            with self.subTest(command=name):
                runs = _runs(latex_to_omml("\\%s{R}" % name))
                self.assertEqual(len(runs), 1)
                self.assertEqual(_tags(runs[0].find("m:rPr", MATH_NS)),
                                 ["scr", "sty"])
                self.assertEqual(_run_properties(runs[0]),
                                 {"scr": script, "sty": "p"})

    def test_alphabet_covers_its_whole_argument_and_only_it(self):
        element = latex_to_omml(r"\mathbb{R}^n")
        base, exponent = (element.find("m:sSup/m:e/m:r", MATH_NS),
                          element.find("m:sSup/m:sup/m:r", MATH_NS))
        self.assertEqual(_run_properties(base)["scr"], "double-struck")
        self.assertNotIn("scr", _run_properties(exponent))

    def test_bold_italic_spellings(self):
        for latex in (r"\boldsymbol{x}", r"\bm{x}", r"\pmb{x}", r"\mathbfit{x}"):
            with self.subTest(latex=latex):
                self.assertEqual(
                    _run_properties(_runs(latex_to_omml(latex))[0]), {"sty": "bi"})

    def test_boldsymbol_makes_a_script_alphabet_bold(self):
        run = _runs(latex_to_omml(r"\boldsymbol{\mathcal{A}}"))[0]
        self.assertEqual(_run_properties(run), {"scr": "script", "sty": "b"})
        outer = _runs(latex_to_omml(r"\mathcal{\boldsymbol{A}}"))[0]
        self.assertEqual(_run_properties(outer), {"scr": "script", "sty": "b"})

    def test_mathrm_is_upright_math_not_literal_text(self):
        r"""``\mathrm{m^2}`` is a superscript in LaTeX -- reading it as the
        literal text "m^2" was silently wrong."""
        element = latex_to_omml(r"\mathrm{m^2}")
        self.assertEqual(_tags(element), ["sSup"])
        self.assertEqual(
            _run_properties(element.find("m:sSup/m:e/m:r", MATH_NS)),
            {"sty": "p"})
        self.assertEqual(
            _run_properties(_runs(latex_to_omml(r"\mathrm{\mu}"))[0]),
            {"sty": "p"})
        self.assertEqual(_texts(latex_to_omml(r"\mathrm{d}x")), ["d", "x"])
        # Letters outside A-Z tokenize as plain characters; they must be
        # upright too, or Word would italicise them.
        for run in _runs(latex_to_omml(r"\mathrm{км}")):
            self.assertEqual(_run_properties(run), {"sty": "p"})

    def test_old_style_font_switches_last_to_the_end_of_their_group(self):
        element = latex_to_omml(r"{\rm d}x + {\bf v} + {\cal L}")
        runs = _runs(element)
        self.assertEqual([_run_properties(run) for run in runs], [
            {"sty": "p"}, {"sty": "i"}, {}, {"sty": "b"}, {},
            {"scr": "script", "sty": "p"},
        ])

    def test_text_bold_and_italic_are_normal_text_with_word_formatting(self):
        r"""``<m:nor>`` and ``<m:sty>`` exclude each other, so bold/italic
        normal text is spelled with ``<w:b/>``/``<w:i/>`` -- placed after
        ``<m:rPr>`` and before ``<m:t>``, as CT_R orders them."""
        cases = {
            r"\textbf{bold}": ["b"], r"\textit{word}": ["i"],
            r"\emph{word}": ["i"], r"\text{plain}": [],
        }
        for latex, word in cases.items():
            with self.subTest(latex=latex):
                run = _runs(latex_to_omml(latex))[0]
                self.assertEqual(_run_properties(run), {"nor": "1"})
                self.assertEqual(_word_tags(run), word)
                self.assertEqual(_tags(run)[-1], "t")
                if word:
                    self.assertEqual(_tags(run)[:2], ["rPr", "rPr"])

    def test_text_sans_and_monospace_use_the_math_alphabet(self):
        for latex, script in ((r"\textsf{sans}", "sans-serif"),
                              (r"\texttt{x y}", "monospace")):
            with self.subTest(latex=latex):
                run = _runs(latex_to_omml(latex))[0]
                self.assertEqual(_run_properties(run),
                                 {"scr": script, "sty": "p"})
        self.assertEqual(_texts(latex_to_omml(r"\texttt{x y}")), ["x y"])

    def test_dollar_math_inside_text_is_math_again(self):
        element = latex_to_omml(r"\text{if $x > 0$ then}")
        runs = _runs(element)
        self.assertEqual(_texts(element), ["if ", "x", ">", "0", " then"])
        self.assertEqual(_run_properties(runs[0]), {"nor": "1"})
        self.assertEqual(_run_properties(runs[1]), {"sty": "i"})

    def test_braces_inside_text_group_rather_than_print(self):
        self.assertEqual(_texts(latex_to_omml(r"\text{a{b}c}")), ["abc"])

    def test_no_new_style_combines_nor_with_sty(self):
        formulas = [
            r"\mathbb{R}", r"\mathcal{L}", r"\textbf{x}", r"\textit{x}",
            r"\texttt{x}", r"\textsf{x}", r"\mathrm{x}", r"{\rm x}",
            r"\boldsymbol{\mathcal{A}}", r"\mathbf{\textbf{x}}",
            r"\varGamma",
        ]
        for latex in formulas:
            with self.subTest(latex=latex):
                for properties in latex_to_omml(latex).iter(qn("m:rPr")):
                    present = _tags(properties)
                    self.assertFalse("nor" in present and "sty" in present)


class OverUnderTests(unittest.TestCase):
    r"""``\overset``, ``\underset``, braces, over/under arrows and the
    extensible arrows."""

    def test_overset_and_stackrel_are_upper_limits(self):
        for latex in (r"\overset{!}{=}", r"\stackrel{!}{=}"):
            with self.subTest(latex=latex):
                limit = latex_to_omml(latex).find("m:limUpp", MATH_NS)
                self.assertEqual(_tags(limit), ["e", "lim"])
                self.assertEqual(_texts(limit.find("m:e", MATH_NS)), ["="])
                self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["!"])

    def test_underset_is_a_lower_limit(self):
        limit = latex_to_omml(
            r"\underset{x}{\operatorname{argmin}}").find("m:limLow", MATH_NS)
        self.assertEqual(_texts(limit.find("m:e", MATH_NS)), ["argmin"])
        self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["x"])

    def test_braces_are_group_characters(self):
        shapes = {
            r"\overbrace{a+b}": ("⏞", "top", "bot"),
            r"\underbrace{a+b}": ("⏟", "bot", "top"),
            r"\overparen{a}": ("⏜", "top", "bot"),
            r"\underbracket{a}": ("⎵", "bot", "top"),
        }
        for latex, (character, position, vertical) in shapes.items():
            with self.subTest(latex=latex):
                group = latex_to_omml(latex).find("m:groupChr", MATH_NS)
                properties = group.find("m:groupChrPr", MATH_NS)
                self.assertEqual(_tags(properties), ["chr", "pos", "vertJc"])
                self.assertEqual(_val(properties, "m:chr"), character)
                self.assertEqual(_val(properties, "m:pos"), position)
                self.assertEqual(_val(properties, "m:vertJc"), vertical)

    def test_brace_annotations_become_limits_around_the_brace(self):
        r"""``\underbrace{a+b}_{n}`` writes n under the brace -- Word's own
        "underbrace with text" is a groupChr inside a limLow."""
        lower = latex_to_omml(r"\underbrace{a+b}_{n}").find("m:limLow", MATH_NS)
        self.assertEqual(_tags(lower.find("m:e", MATH_NS)), ["groupChr"])
        self.assertEqual(_texts(lower.find("m:lim", MATH_NS)), ["n"])
        upper = latex_to_omml(r"\overbrace{a+b}^{n}").find("m:limUpp", MATH_NS)
        self.assertEqual(_tags(upper.find("m:e", MATH_NS)), ["groupChr"])
        self.assertEqual(_texts(upper.find("m:lim", MATH_NS)), ["n"])

    def test_over_and_under_arrows_stretch_over_their_argument(self):
        for latex, character, position in (
            (r"\overrightarrow{AB}", "→", "top"),
            (r"\overleftarrow{AB}", "←", "top"),
            (r"\overleftrightarrow{AB}", "↔", "top"),
            (r"\underrightarrow{AB}", "→", "bot"),
        ):
            with self.subTest(latex=latex):
                group = latex_to_omml(latex).find("m:groupChr", MATH_NS)
                self.assertEqual(_val(group, "m:groupChrPr/m:chr"), character)
                self.assertEqual(_val(group, "m:groupChrPr/m:pos"), position)
                self.assertEqual(_texts(group.find("m:e", MATH_NS)), ["A", "B"])

    def test_an_arrow_takes_ordinary_scripts_not_limits(self):
        element = latex_to_omml(r"\overrightarrow{AB}^2")
        self.assertEqual(_tags(element), ["sSup"])

    def test_extensible_arrow_with_text_above(self):
        group = latex_to_omml(r"\xrightarrow{f}").find("m:groupChr", MATH_NS)
        self.assertEqual(_val(group, "m:groupChrPr/m:chr"), "→")
        self.assertEqual(_val(group, "m:groupChrPr/m:pos"), "bot")
        self.assertEqual(_val(group, "m:groupChrPr/m:vertJc"), "bot")
        self.assertEqual(_texts(group.find("m:e", MATH_NS)), ["f"])

    def test_extensible_arrow_with_text_above_and_below(self):
        lower = latex_to_omml(r"\xleftarrow[g]{f}").find("m:limLow", MATH_NS)
        self.assertEqual(
            _val(lower, "m:e/m:groupChr/m:groupChrPr/m:chr"), "←")
        self.assertEqual(_texts(lower.find("m:e", MATH_NS)), ["f"])
        self.assertEqual(_texts(lower.find("m:lim", MATH_NS)), ["g"])

    def test_extensible_arrow_with_text_below_only(self):
        group = latex_to_omml(r"\xrightarrow[g]{}").find("m:groupChr", MATH_NS)
        self.assertEqual(_val(group, "m:groupChrPr/m:pos"), "top")
        self.assertEqual(_texts(group.find("m:e", MATH_NS)), ["g"])
        self.assertEqual(_texts(latex_to_omml(r"\xrightarrow{}")), ["→"])


class BoxAndDecorationTests(unittest.TestCase):
    r"""``\boxed``, ``\cancel``, the phantoms, ``\hspace`` and colour."""

    def test_boxed_is_a_plain_border_box(self):
        box = latex_to_omml(r"\boxed{E = mc^2}").find("m:borderBox", MATH_NS)
        self.assertEqual(_tags(box), ["e"])
        framed = latex_to_omml(r"\fbox{two words}").find("m:borderBox", MATH_NS)
        self.assertEqual(_texts(framed), ["two words"])
        self.assertEqual(_run_properties(_runs(framed)[0]), {"nor": "1"})

    def test_cancels_hide_the_frame_and_strike_diagonally(self):
        strikes = {
            r"\cancel{x}": ["strikeBLTR"], r"\bcancel{x}": ["strikeTLBR"],
            r"\xcancel{x}": ["strikeBLTR", "strikeTLBR"],
        }
        for latex, expected in strikes.items():
            with self.subTest(latex=latex):
                properties = latex_to_omml(latex).find(
                    "m:borderBox/m:borderBoxPr", MATH_NS)
                # CT_BorderBoxPr's sequence order, hides first.
                self.assertEqual(
                    _tags(properties),
                    ["hideTop", "hideBot", "hideLeft", "hideRight", *expected])

    def test_phantoms_hide_and_zero_the_right_dimensions(self):
        shapes = {
            r"\phantom{x}": {"show": "0"},
            r"\hphantom{x}": {"show": "0", "zeroAsc": "1", "zeroDesc": "1"},
            r"\vphantom{x}": {"show": "0", "zeroWid": "1"},
            r"\smash{x}": {"zeroAsc": "1", "zeroDesc": "1"},
            r"\smash[b]{x}": {"zeroDesc": "1"},
        }
        for latex, expected in shapes.items():
            with self.subTest(latex=latex):
                properties = latex_to_omml(latex).find(
                    "m:phant/m:phantPr", MATH_NS)
                self.assertEqual(
                    {child.tag.split("}")[-1]: child.get(qn("m:val"))
                     for child in properties}, expected)
                order = ["show", "zeroWid", "zeroAsc", "zeroDesc"]
                tags = _tags(properties)
                self.assertEqual(tags, sorted(tags, key=order.index))
        self.assertEqual(
            _texts(latex_to_omml(r"\phantom{abc}").find(
                "m:phant/m:e", MATH_NS)), ["a", "b", "c"])

    def test_hspace_accepts_any_length_and_drops_negative_space(self):
        for latex in (r"a\hspace{1cm}b", r"a\hspace*{2em}b",
                      r"a\hspace{10pt}b", r"a\hspace{0.5in}b",
                      r"a\mkern18mu b", r"a\kern 3pt b", r"a\mspace{4mu}b"):
            with self.subTest(latex=latex):
                texts = _texts(latex_to_omml(latex))
                self.assertEqual((texts[0], texts[-1]), ("a", "b"))
                self.assertEqual(len(texts), 3)
                self.assertTrue(texts[1].strip() == "" and texts[1])
        self.assertEqual(_texts(latex_to_omml(r"a\hspace{-1em}b")), ["a", "b"])
        self.assertEqual(_texts(latex_to_omml(r"a\mkern-3mu b")), ["a", "b"])
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"a\hspace{\fill}b")
        self.assertIn("hspace", str(caught.exception))

    def test_color_switch_lasts_to_the_end_of_its_group(self):
        element = latex_to_omml(r"{\color{red} x + y} + z")
        colors = [run.find("w:rPr/w:color", {"w": _nsmap["w"]})
                  for run in _runs(element)]
        self.assertEqual(
            [None if c is None else c.get(qn("w:val")) for c in colors],
            ["FF0000", "FF0000", "FF0000", None, None])

    def test_color_with_a_braced_argument_and_textcolor(self):
        for latex in (r"\color{blue}{x}", r"\textcolor{blue}{x}"):
            with self.subTest(latex=latex):
                run = _runs(latex_to_omml(latex))[0]
                self.assertEqual(
                    run.find(qn("w:rPr")).find(qn("w:color")).get(qn("w:val")),
                    "0000FF")
                self.assertEqual(_tags(run), ["rPr", "rPr", "t"])
        plain = _runs(latex_to_omml(r"\textcolor{blue}{x} + y"))[-1]
        self.assertEqual(_word_tags(plain), [])

    def test_color_specifications(self):
        cases = {
            r"\color[HTML]{ff8800}{x}": "FF8800",
            r"\color{#1E90FF}{x}": "1E90FF",
            r"\color{#f80}{x}": "FF8800",
            r"\color[rgb]{1,0.5,0}{x}": "FF8000",
            r"\color[RGB]{255,128,0}{x}": "FF8000",
            r"\color[gray]{0.5}{x}": "808080",
            r"\color{grey}{x}": "808080",
            r"\color{green}{x}": "008000",
        }
        for latex, value in cases.items():
            with self.subTest(latex=latex):
                color = _runs(latex_to_omml(latex))[0].find(
                    qn("w:rPr")).find(qn("w:color"))
                self.assertEqual(color.get(qn("w:val")), value)

    def test_unknown_colour_is_refused(self):
        for latex in (r"\color{notacolor}{x}", r"\color[cmyk]{0,1,1,0}{x}",
                      r"\textcolor[HTML]{GG0000}{x}"):
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError):
                    latex_to_omml(latex)

    def test_color_reaches_structure_glyphs_through_ctrlpr(self):
        r"""A fraction bar or radical sign is not a run: its colour lives
        in the ``<m:ctrlPr>`` that closes the structure's properties."""
        fraction = latex_to_omml(r"\color{red}{\frac{a}{b}}").find("m:f", MATH_NS)
        properties = fraction.find("m:fPr", MATH_NS)
        self.assertEqual(_tags(fraction)[0], "fPr")
        self.assertEqual(_tags(properties)[-1], "ctrlPr")
        self.assertEqual(
            properties.find("m:ctrlPr", MATH_NS).find(qn("w:rPr"))
            .find(qn("w:color")).get(qn("w:val")), "FF0000")
        nary = latex_to_omml(r"\color{red} \sum_i a_i").find("m:nary", MATH_NS)
        self.assertEqual(_tags(nary.find("m:naryPr", MATH_NS))[-1], "ctrlPr")

    def test_inner_colour_wins_over_outer(self):
        element = latex_to_omml(r"\color{red}{a + \textcolor{blue}{b}}")
        values = [run.find(qn("w:rPr")).find(qn("w:color")).get(qn("w:val"))
                  for run in _runs(element)]
        self.assertEqual(values, ["FF0000", "FF0000", "0000FF"])

    def test_colour_ends_at_an_alignment_cell(self):
        r"""amsmath makes each ``&`` cell a group, so a colour set before
        ``&`` does not leak into the next cell."""
        line = latex_to_omml(
            r"\begin{aligned} \color{red} a &= b \\ c &= d \end{aligned}"
        ).findall("m:eqArr/m:e", MATH_NS)[0]
        colored = [run.find(qn("w:rPr")) is not None for run in _runs(line)]
        self.assertEqual(colored, [True, False, False])


class NumberingTests(unittest.TestCase):
    r"""``\tag``, ``\label``, ``\nonumber``, ``\notag`` and
    `split_equation_tag`."""

    def test_label_nonumber_and_notag_produce_nothing(self):
        for latex in (r"E = mc^2 \label{eq:einstein}", r"E = mc^2 \nonumber",
                      r"E = mc^2 \notag", r"\label{a:b_c} E = mc^2"):
            with self.subTest(latex=latex):
                self.assertEqual(xml_of(latex), xml_of(r"E = mc^2"))

    def test_tag_is_an_upright_number_set_apart_at_the_end(self):
        element = latex_to_omml(r"E = mc^2 \tag{1}")
        runs = list(element)[-2:]
        self.assertEqual(_texts(runs[0]), [" "])
        self.assertEqual(_texts(runs[1]), ["(1)"])
        self.assertEqual(_run_properties(runs[1]), {"nor": "1"})
        starred = list(latex_to_omml(r"E = mc^2 \tag*{A}"))[-1]
        self.assertEqual(_texts(starred), ["A"])

    def test_tag_moves_to_the_end_of_its_line_wherever_it_is_written(self):
        self.assertEqual(xml_of(r"\tag{2} a = b"), xml_of(r"a = b \tag{2}"))
        element = latex_to_omml(r"\sum_i a_i \tag{3}")
        self.assertEqual(_tags(element), ["nary", "r", "r"])
        self.assertEqual(_texts(element.find("m:nary/m:e", MATH_NS)), ["a", "i"])

    def test_each_line_keeps_its_own_tag(self):
        lines = latex_to_omml(
            r"\begin{align} a &= b \tag{1} \\ c &= d \tag{2} \end{align}"
        ).findall("m:eqArr/m:e", MATH_NS)
        self.assertEqual([_texts(line)[-1] for line in lines], ["(1)", "(2)"])
        for line in lines:
            self.assertEqual(_aligned_runs(line), ["="])

    def test_two_tags_on_one_line_are_refused(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"a \tag{1} \tag{2}")
        self.assertIn("tag", str(caught.exception))

    def test_split_equation_tag_returns_the_tag_text(self):
        self.assertEqual(split_equation_tag(r"E = mc^2 \tag{3}"),
                         ("E = mc^2", "3"))
        self.assertEqual(split_equation_tag(r"E = mc^2 \tag*{A}"),
                         ("E = mc^2", "A"))
        self.assertEqual(split_equation_tag(r"\tag {1.2a} x"), ("x", "1.2a"))
        self.assertEqual(split_equation_tag(r"a \tag{\text{a}{b}}"),
                         ("a", r"\text{a}{b}"))

    def test_split_equation_tag_drops_labels_and_no_number_markers(self):
        self.assertEqual(
            split_equation_tag(r"a = b \label{eq:x} \nonumber"), ("a = b", None))
        self.assertEqual(split_equation_tag(r"a \notag"), ("a", None))
        self.assertEqual(
            split_equation_tag(r"a = b \label{eq:x} \tag{7}"), ("a = b", "7"))
        self.assertEqual(split_equation_tag("x"), ("x", None))

    def test_split_equation_tag_leaves_per_line_tags_in_place(self):
        source = r"a \tag{1} \\ b \tag{2}"
        self.assertEqual(split_equation_tag(source), (source, None))
        remaining, tag = split_equation_tag(r"a \tag{1} \label{x} \\ b \tag{2}")
        self.assertEqual((remaining, tag), (r"a \tag{1}  \\ b \tag{2}", None))

    def test_split_equation_tag_never_raises_and_respects_escapes(self):
        for source in (r"x \tag 1", r"x \tag{1", r"a\tagx{1}",
                       "a \\\\tag{1}", "\\", r"\tag"):
            with self.subTest(source=source):
                self.assertEqual(split_equation_tag(source),
                                 (source.strip(), None))

    def test_split_result_converts(self):
        remaining, tag = split_equation_tag(
            r"\begin{aligned} a &= b \\ c &= d \end{aligned} \label{e} \tag{4}")
        self.assertEqual(tag, "4")
        self.assertEqual(len(latex_to_omml(remaining).findall(
            "m:eqArr/m:e", MATH_NS)), 2)


class StyleSwitchAndDelimiterTests(unittest.TestCase):
    r"""``\displaystyle`` and friends, ``\limits``/``\nolimits``, ``\big``
    sizes and ``\middle``."""

    def test_style_and_size_switches_are_no_ops(self):
        for switch in ("displaystyle", "textstyle", "scriptstyle",
                       "scriptscriptstyle", "small", "Large"):
            with self.subTest(switch=switch):
                self.assertEqual(
                    xml_of("\\%s \\frac{a}{b}" % switch), xml_of(r"\frac{a}{b}"))
        with self.assertRaises(UnsupportedLatexError):
            latex_to_omml(r"x^\displaystyle")

    def test_limits_and_nolimits_set_the_nary_limit_location(self):
        cases = {
            r"\sum\limits_{i=1}^n a_i": "undOvr",
            r"\sum\nolimits_{i=1}^n a_i": "subSup",
            r"\int\limits_0^1 f": "undOvr",
            r"\int\nolimits_0^1 f": "subSup",
            r"\sum\nolimits\limits_i a": "undOvr",
        }
        for latex, location in cases.items():
            with self.subTest(latex=latex):
                nary = latex_to_omml(latex).find("m:nary", MATH_NS)
                self.assertEqual(_val(nary, "m:naryPr/m:limLoc"), location)
                self.assertEqual(_tags(nary.find("m:naryPr", MATH_NS))[:2],
                                 ["chr", "limLoc"])

    def test_limits_on_function_names(self):
        self.assertEqual(_tags(latex_to_omml(r"\max_{x} f")), ["limLow", "r"])
        self.assertEqual(_tags(latex_to_omml(r"\lim\nolimits_{x} f")),
                         ["sSub", "r"])
        self.assertEqual(_tags(latex_to_omml(r"\sin\limits_{x} f")),
                         ["limLow", "r"])
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"x\limits_0")
        self.assertIn("limits", str(caught.exception))

    def test_big_delimiters_are_plain_characters(self):
        element = latex_to_omml(r"\bigl( x \bigr) \Bigl[ y \Bigr] \bigg\{ z \bigg\}")
        self.assertIsNone(element.find(".//m:d", MATH_NS))
        self.assertEqual(_texts(element),
                         ["(", "x", ")", "[", "y", "]", "{", "z", "}"])
        self.assertEqual(_texts(latex_to_omml(r"\Big\langle x \Big\rangle")),
                         ["⟨", "x", "⟩"])
        scripted = latex_to_omml(r"\frac{dy}{dx}\bigg|_{x=0}")
        self.assertEqual(_texts(scripted.find("m:sSub/m:e", MATH_NS)), ["|"])

    def test_middle_becomes_a_separator(self):
        delimiter = latex_to_omml(r"\left( a \middle| b \right)").find("m:d", MATH_NS)
        self.assertEqual(_tags(delimiter.find("m:dPr", MATH_NS)),
                         ["begChr", "sepChr", "endChr"])
        self.assertEqual(_val(delimiter, "m:dPr/m:sepChr"), "|")
        self.assertEqual([_texts(e) for e in delimiter.findall("m:e", MATH_NS)],
                         [["a"], ["b"]])
        braces = latex_to_omml(
            r"\left\{ x \middle\| y \middle\| z \right\}").find("m:d", MATH_NS)
        self.assertEqual(_val(braces, "m:dPr/m:sepChr"), "‖")
        self.assertEqual(len(braces.findall("m:e", MATH_NS)), 3)

    def test_mixed_or_stray_middle_is_refused(self):
        for latex in (r"\left( a \middle| b \middle\| c \right)",
                      r"\left( a \middle. b \right)", r"\middle| x"):
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml(latex)
                self.assertIn("middle", str(caught.exception))


class CommandTests(unittest.TestCase):
    r"""``\cfrac``, ``\dbinom``, ``\pmod``, ``\operatorname*``, ``\not``
    and the other commands added alongside them."""

    def test_cfrac_is_a_fraction(self):
        outer = latex_to_omml(r"\cfrac{1}{1 + \cfrac{1}{x}}").find("m:f", MATH_NS)
        self.assertIsNotNone(outer.find("m:den/m:f", MATH_NS))
        self.assertEqual(xml_of(r"\cfrac[l]{a}{b}"), xml_of(r"\frac{a}{b}"))

    def test_brace_and_brack_infixes_fence_a_barless_stack(self):
        for latex, pair in ((r"{n \brace k}", ("{", "}")),
                            (r"{n \brack k}", ("[", "]"))):
            with self.subTest(latex=latex):
                element = latex_to_omml(latex)
                self.assertEqual((_val(element, "m:d/m:dPr/m:begChr"),
                                  _val(element, "m:d/m:dPr/m:endChr")), pair)
                self.assertEqual(
                    _val(element, "m:d/m:e/m:f/m:fPr/m:type"), "noBar")

    def test_dbinom_and_tbinom_match_binom(self):
        for name in ("dbinom", "tbinom"):
            with self.subTest(command=name):
                self.assertEqual(xml_of("\\%s{n}{k}" % name),
                                 xml_of(r"\binom{n}{k}"))

    def test_pmod_puts_an_upright_mod_in_parentheses(self):
        element = latex_to_omml(r"a \equiv b \pmod{n}")
        delimiter = element.find("m:d", MATH_NS)
        self.assertEqual(_val(delimiter, "m:dPr/m:begChr"), "(")
        runs = _runs(delimiter)
        self.assertEqual([_texts(run) for run in runs], [["mod"], [" "], ["n"]])
        self.assertEqual(_run_properties(runs[0]), {"nor": "1"})

    def test_bmod_is_an_upright_operator(self):
        element = latex_to_omml(r"a \bmod b")
        self.assertEqual(_texts(element), ["a", " ", "mod", " ", "b"])
        self.assertEqual(_run_properties(_runs(element)[2]), {"nor": "1"})

    def test_pr_is_an_upright_function_name(self):
        run = _runs(latex_to_omml(r"\Pr(A)"))[0]
        self.assertEqual(_texts(run), ["Pr"])
        self.assertEqual(_run_properties(run), {"nor": "1"})

    def test_operatorname_star_takes_limits(self):
        element = latex_to_omml(r"\operatorname*{argmax}_{x \in X} f(x)")
        limit = element.find("m:limLow", MATH_NS)
        self.assertEqual(_texts(limit.find("m:e", MATH_NS)), ["argmax"])
        self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["x", "∈", "X"])
        plain = latex_to_omml(r"\operatorname{argmax}_x f")
        self.assertEqual(_tags(plain), ["sSub", "r"])

    def test_operatorname_star_without_braces_names_the_command(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\operatorname* x")
        self.assertIn("\\operatorname*", str(caught.exception))
        self.assertIn("{", str(caught.exception))

    def test_mathop_takes_limits(self):
        limit = latex_to_omml(
            r"\mathop{\mathrm{arg\,min}}_{x} f").find("m:limLow", MATH_NS)
        self.assertEqual(_texts(limit.find("m:lim", MATH_NS)), ["x"])

    def test_not_composes_the_negated_relation(self):
        expected = {
            r"\not=": "≠", r"\not\in": "∉", r"\not\subset": "⊄",
            r"\not\subseteq": "⊈", r"\not\equiv": "≢", r"\not\sim": "≁",
            r"\not<": "≮", r"\not>": "≯", r"\not\le": "≰", r"\not\leq": "≰",
            r"\not\ge": "≱", r"\not\approx": "≉", r"\not\cong": "≇",
            r"\not\parallel": "∦", r"\not\exists": "∄", r"\not\ni": "∌",
            r"\not \mid": "∤",
        }
        for latex, character in expected.items():
            with self.subTest(latex=latex):
                self.assertEqual(_texts(latex_to_omml(latex)), [character])
        # No precomposed form: the combining long solidus overlay.
        self.assertEqual(_texts(latex_to_omml(r"\not\propto")), ["∝̸"])

    def test_not_refuses_anything_but_a_single_symbol(self):
        for latex in (r"\not\frac{a}{b}", r"\not", r"\not}"):
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError) as caught:
                    latex_to_omml(latex if latex != r"\not}" else r"{\not}")
                self.assertIn("not", str(caught.exception))

    def test_double_bars_outside_and_inside_left_right(self):
        self.assertEqual(_texts(latex_to_omml(r"\|x\|")), ["‖", "x", "‖"])
        self.assertEqual(_texts(latex_to_omml(r"\lVert x \rVert")),
                         ["‖", "x", "‖"])
        self.assertEqual(_texts(latex_to_omml(r"\lvert x \rvert")),
                         ["|", "x", "|"])
        for latex, pair in ((r"\left\lVert x \right\rVert", ("‖", "‖")),
                            (r"\left\lvert x \right\rvert", ("|", "|")),
                            (r"\left< x \right>", ("⟨", "⟩"))):
            with self.subTest(latex=latex):
                element = latex_to_omml(latex)
                self.assertEqual(
                    (_val(element, "m:d/m:dPr/m:begChr"),
                     _val(element, "m:d/m:dPr/m:endChr")), pair)

    def test_dirac_notation(self):
        ket = latex_to_omml(r"\ket{\psi}")
        self.assertEqual((_val(ket, "m:d/m:dPr/m:begChr"),
                          _val(ket, "m:d/m:dPr/m:endChr")), ("|", "⟩"))
        braket = latex_to_omml(r"\braket{\phi | \psi}").find("m:d", MATH_NS)
        self.assertEqual(_val(braket, "m:dPr/m:sepChr"), "|")
        self.assertEqual(len(braket.findall("m:e", MATH_NS)), 2)

    def test_prime_and_tie(self):
        self.assertEqual(_texts(latex_to_omml("f'(x)")),
                         ["f", "′", "(", "x", ")"])
        self.assertEqual(_texts(latex_to_omml("a~b")), ["a", " ", "b"])

    def test_sideset_is_refused_with_a_clear_message(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\sideset{}{'}\sum_{n} a_n")
        message = str(caught.exception)
        self.assertIn("sideset", message)
        self.assertIn("not supported", message)

    def test_missing_argument_names_the_command(self):
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\frac{a}")
        self.assertIn("\\frac", str(caught.exception))
        with self.assertRaises(UnsupportedLatexError) as caught:
            latex_to_omml(r"\text x")
        self.assertIn("\\text", str(caught.exception))

    def test_deliberate_refusals_still_hold(self):
        must_raise = [
            r"\begin{array}{c|c} a & b \end{array}",
            r"\begin{array}{p{2cm}} a \end{array}",
            r"\begin{array}{@{}c} a \end{array}",
            r"\begin{array}{cc} \hline a & b \end{array}",
            r"\matrix{a & b}", r"\cases{a & b}",
            r"a \over b \over c", "a & b",
        ]
        for latex in must_raise:
            with self.subTest(latex=latex):
                with self.assertRaises(UnsupportedLatexError):
                    latex_to_omml(latex)


class SymbolTests(unittest.TestCase):
    """The added symbols map to the right Unicode code points."""

    EXPECTED = {
        "top": "⊤", "bot": "⊥", "dagger": "†", "ddagger": "‡", "lnot": "¬",
        "implies": "⟹", "impliedby": "⟸", "iff": "⟺",
        "uparrow": "↑", "downarrow": "↓", "updownarrow": "↕",
        "Uparrow": "⇑", "Downarrow": "⇓", "Updownarrow": "⇕",
        "nearrow": "↗", "searrow": "↘", "nwarrow": "↖", "swarrow": "↙",
        "ni": "∋", "mid": "∣", "nmid": "∤",
        "Longrightarrow": "⟹", "Longleftarrow": "⟸",
        "Longleftrightarrow": "⟺", "longrightarrow": "⟶",
        "longleftarrow": "⟵", "longleftrightarrow": "⟷", "longmapsto": "⟼",
        "hookrightarrow": "↪", "hookleftarrow": "↩",
        "rightleftharpoons": "⇌", "leftrightharpoons": "⇋",
        "rightharpoonup": "⇀", "leftharpoonup": "↼",
        "doteq": "≐", "triangleq": "≜", "coloneqq": "≔", "eqqcolon": "≕",
        "wp": "℘", "square": "□", "blacksquare": "■", "checkmark": "✓",
        "bigcirc": "◯", "diamond": "⋄", "Diamond": "◇", "triangle": "△",
        "triangledown": "▽", "therefore": "∴", "because": "∵",
        "nexists": "∄", "complement": "∁", "hslash": "ℏ", "imath": "ı",
        "jmath": "ȷ", "mho": "℧", "leqslant": "⩽", "geqslant": "⩾",
        "lesssim": "≲", "gtrsim": "≳", "prec": "≺", "succ": "≻",
        "preceq": "⪯", "succeq": "⪰", "sqsubset": "⊏", "sqsupset": "⊐",
        "sqsubseteq": "⊑", "sqsupseteq": "⊒", "vdash": "⊢", "dashv": "⊣",
        "models": "⊨", "asymp": "≍", "bowtie": "⋈", "smile": "⌣",
        "frown": "⌢", "amalg": "⨿", "wr": "≀", "uplus": "⊎", "sqcup": "⊔",
        "sqcap": "⊓", "odot": "⊙", "ominus": "⊖", "oslash": "⊘",
        "circledast": "⊛", "boxplus": "⊞", "boxtimes": "⊠", "lhd": "⊲",
        "rhd": "⊳", "unlhd": "⊴", "unrhd": "⊵", "nleq": "≰", "ngeq": "≱",
        "nsubseteq": "⊈", "nsupseteq": "⊉", "ncong": "≇", "nsim": "≁",
        "varkappa": "ϰ", "varsigma": "ς", "varrho": "ϱ", "varpi": "ϖ",
        "digamma": "ϝ", "backslash": "\\", "measuredangle": "∡",
        "sphericalangle": "∢", "degree": "°", "surd": "√", "flat": "♭",
        "sharp": "♯", "natural": "♮", "clubsuit": "♣", "diamondsuit": "♢",
        "heartsuit": "♡", "spadesuit": "♠", "lozenge": "◊",
        "blacklozenge": "⧫", "bigstar": "★", "circledR": "®",
        "copyright": "©", "pounds": "£", "yen": "¥", "euro": "€",
        "ldots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱", "colon": ":",
        "langle": "⟨", "rangle": "⟩", "lfloor": "⌊", "rceil": "⌉",
        "lt": "<", "gt": ">", "vee": "∨", "wedge": "∧",
        # LaTeX's \epsilon and \phi are the lunate and straight forms.
        "epsilon": "ϵ", "varepsilon": "ε", "phi": "ϕ", "varphi": "φ",
    }

    def test_every_added_symbol_has_its_code_point(self):
        for name, character in self.EXPECTED.items():
            with self.subTest(command=name):
                self.assertEqual(_SYMBOLS.get(name), character)
                self.assertEqual(_texts(latex_to_omml("\\" + name)), [character])

    def test_added_big_operators_are_n_ary(self):
        for name, character, location in (
            ("bigodot", "⨀", "undOvr"), ("biguplus", "⨄", "undOvr"),
            ("bigsqcup", "⨆", "undOvr"), ("iiiint", "⨌", "subSup"),
            ("oiint", "∯", "subSup"),
        ):
            with self.subTest(command=name):
                nary = latex_to_omml("\\%s_i A_i" % name).find("m:nary", MATH_NS)
                self.assertEqual(_val(nary, "m:naryPr/m:chr"), character)
                self.assertEqual(_val(nary, "m:naryPr/m:limLoc"), location)

    def test_variant_capitals_are_italic(self):
        run = _runs(latex_to_omml(r"\varGamma"))[0]
        self.assertEqual(_texts(run), ["Γ"])
        self.assertEqual(_run_properties(run), {"sty": "i"})


# Common LLM-style formulas: every one must convert without raising.
SMOKE_FORMULAS = [
    r"\begin{aligned} a &= b \\ c &= d \end{aligned}",
    r"\begin{aligned} f(x) &= (x+1)^2 \\ &= x^2 + 2x + 1 \end{aligned}",
    r"\begin{aligned} \nabla \cdot \mathbf{E} &= \frac{\rho}{\varepsilon_0} \\ \nabla \cdot \mathbf{B} &= 0 \end{aligned}",
    r"\begin{alignedat}{2} x &= 1 &\quad y &= 2 \end{alignedat}",
    r"\begin{split} a &= b + c \\ &= d \end{split}",
    r"\begin{gathered} a = b \\ c = d \end{gathered}",
    r"\begin{align*} a &= b \\ c &= d \end{align*}",
    r"\begin{gather} x \\ y \end{gather}",
    r"\begin{equation} E = mc^2 \end{equation}",
    r"\begin{smallmatrix} a & b \\ c & d \end{smallmatrix}",
    r"\begin{pmatrix} 1 & 0 \\ 0 & 1 \end{pmatrix}",
    r"\begin{bmatrix} a_{11} & a_{12} \\ a_{21} & a_{22} \end{bmatrix}",
    r"\begin{vmatrix} a & b \\ c & d \end{vmatrix} = ad - bc",
    r"\begin{cases} x & \text{if } x \ge 0 \\ -x & \text{otherwise} \end{cases}",
    r"\begin{dcases} \frac{1}{2} & x > 0 \\ 0 & x \le 0 \end{dcases}",
    r"\begin{rcases} a & b \\ c & d \end{rcases} \implies e",
    r"\begin{cases*} 1 & if $x$ is odd \\ 0 & otherwise \end{cases*}",
    r"\left\{ \begin{aligned} x + y &= 3 \\ x - y &= 1 \end{aligned} \right.",
    r"\mathbb{R}^n", r"x \in \mathbb{Z}_{\ge 0}", r"\mathbb{E}[X] = \mu",
    r"\mathcal{L}(\theta)", r"\mathcal{O}(n \log n)", r"\mathscr{F}",
    r"\mathfrak{g}", r"\mathsf{T}", r"\mathtt{x}", r"\mathbbm{1}_{A}",
    r"\boldsymbol{\mu}", r"\boldsymbol{\Sigma}^{-1}", r"\bm{x}",
    r"\mathbf{x}^\top \mathbf{A} \mathbf{x}", r"\textbf{Note:}\ x > 0",
    r"\textit{see } x", r"\texttt{id}", r"\textsf{A}",
    r"\text{if } x > 0 \text{ and $y$ is odd}",
    r"\overset{!}{=}", r"\stackrel{\text{def}}{=}",
    r"\underset{x}{\operatorname{argmin}} f(x)",
    r"\overbrace{a + b + c}^{3}",
    r"\underbrace{1 + 2 + \cdots + n}_{n \text{ terms}}",
    r"\overrightarrow{AB}", r"\overleftarrow{AB}", r"\overleftrightarrow{AB}",
    r"A \xrightarrow{f} B", r"A \xleftarrow[g]{f} B",
    r"\boxed{E = mc^2}", r"\cancel{x} + \bcancel{y} + \xcancel{z}",
    r"a\phantom{bc}d", r"a\hphantom{bc}d", r"a\vphantom{\frac{1}{2}}d",
    r"a \hspace{1cm} b", r"\color{red}{x} + y", r"{\color{blue} x + y} + z",
    r"\textcolor{green}{x}", r"\color[HTML]{FF8800}{x}",
    r"E = mc^2 \label{eq:einstein}", r"E = mc^2 \tag{1}", r"E \tag*{A}",
    r"a = b \nonumber", r"a = b \notag",
    r"\displaystyle \sum_{i=1}^n i", r"\textstyle \int_0^1 f",
    r"\sum\limits_{i=1}^n a_i", r"\int\limits_0^1 f(x)\,dx",
    r"\sum\nolimits_{i} a_i", r"\bigl( x \bigr)", r"\Bigl[ x \Bigr]",
    r"\bigg\{ x \bigg\}", r"\left( a \middle| b \right)",
    r"\cfrac{1}{1 + \cfrac{1}{x}}", r"\dbinom{n}{k}", r"\tbinom{n}{k}",
    r"a \equiv b \pmod{n}", r"a \bmod b", r"\Pr(A \mid B)",
    r"\operatorname*{argmax}_{x \in X} f(x)", r"\operatorname{tr}(A)",
    r"\mathop{\mathrm{arg\,min}}_{x} f(x)", r"a \not= b", r"x \not\in A",
    r"A \not\subseteq B", r"a \not\equiv b \pmod{p}", r"\|x\|_2",
    r"\lVert x \rVert", r"\left\lVert x \right\rVert_\infty",
    r"\lvert x \rvert", r"\top", r"A^\dagger", r"\neg p \lor q",
    r"p \implies q", r"p \iff q", r"\uparrow \downarrow",
    r"\mathbb{N} \ni n", r"a \mid b", r"a \nmid b",
    r"f \colon X \to Y", r"x \mapsto x^2", r"A \Longrightarrow B",
    r"\hookrightarrow", r"\rightleftharpoons", r"a \doteq b",
    r"a \triangleq b", r"x \coloneqq 1", r"\wp", r"\square",
    r"\blacksquare", r"\checkmark", r"\therefore", r"\because",
    r"\nexists x", r"A^\complement", r"\hslash", r"\imath",
    r"a \leqslant b", r"a \lesssim b", r"a \prec b", r"A \sqsubseteq B",
    r"\Gamma \vdash \phi", r"\models", r"A \uplus B", r"A \sqcup B",
    r"x \odot y", r"\bigodot_i A_i", r"\biguplus_i A_i", r"\bigsqcup_i A_i",
    r"\heartsuit", r"\pounds 5", r"90^\circ", r"90\degree",
    r"\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}",
    r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}",
    r"\lim_{n \to \infty} \left(1 + \frac{1}{n}\right)^n = e",
    r"\sum_{k=0}^{\infty} \frac{x^k}{k!} = e^x",
    r"P(A \mid B) = \frac{P(B \mid A)\,P(A)}{P(B)}",
    r"\nabla \times \mathbf{E} = -\frac{\partial \mathbf{B}}{\partial t}",
    r"\langle u, v \rangle", r"\lfloor x \rfloor + \lceil y \rceil",
    r"f'(x) + f''(x)", r"\left.\frac{df}{dx}\right|_{x=0}",
    r"\frac{dy}{dx}\bigg|_{x=0}",
    r"\mathrm{Attention}(Q,K,V) = \mathrm{softmax}\left(\frac{QK^T}{\sqrt{d_k}}\right)V",
    r"\mathcal{N}(\mu, \sigma^2)", r"D_{\mathrm{KL}}(P \| Q)",
    r"\theta \leftarrow \theta - \eta \nabla_\theta \mathcal{L}",
    r"\mathbb{E}_{x \sim p}[f(x)]",
    r"\Delta x \, \Delta p \geq \frac{\hbar}{2}",
    r"i\hbar\frac{\partial}{\partial t}\Psi = \hat{H}\Psi",
    r"\ket{\psi} = \alpha\ket{0} + \beta\ket{1}", r"\braket{\phi | \psi}",
    r"\max_{x \in S} f(x)", r"\arg\max_\theta L(\theta)",
    r"\det(A - \lambda I) = 0", r"\mathrm{H_2O}", r"\mathrm{m/s^2}",
    r"{\rm d}x", r"\vec{F} = m\vec{a}", r"\mathring{A}",
    r"{}^{14}_{6}\mathrm{C}", r"\sum_{\substack{i=1 \\ i \neq j}}^{n} a_i",
    r"a \lt b", r"\oint_C \mathbf{F} \cdot d\mathbf{r}",
    r"\{x \in \mathbb{R} : x > 0\}", r"1{,}000", r"a~b",
    r"\begin{cases} 1 & x > 0 \\[4pt] 0 & \text{otherwise} \end{cases}",
    r"\Big( \frac{a}{b} \Big)^2", r"\color{red} \sum_i a_i",
    r"\begin{equation*} a \tag{2.1} \end{equation*}",
    r"\frac{\partial^2 u}{\partial t^2} = c^2 \nabla^2 u",
    r"\mathbf{W}^{[l]}", r"x \in [0, 1)",
    r"\binom{n}{k} p^k (1-p)^{n-k}", r"e^{i\pi} + 1 = 0",
    r"\forall \epsilon > 0, \exists \delta > 0",
    r"\sigma(z) = \frac{1}{1 + e^{-z}}",
    r"\hat{\beta} = (X^\top X)^{-1} X^\top y",
    r"\text{softmax}(z)_i = \frac{e^{z_i}}{\sum_{j} e^{z_j}}",
    r"\smash{x}", r"a \mkern-3mu b", r"\overparen{AB}",
    r"\begin{array}{lcr} a & b & c \end{array}",
]


class SmokeTests(unittest.TestCase):
    def test_common_formulas_convert_without_raising(self):
        self.assertGreaterEqual(len(SMOKE_FORMULAS), 150)
        for latex in SMOKE_FORMULAS:
            with self.subTest(latex=latex):
                element = latex_to_omml(latex)
                self.assertEqual(element.tag, qn("m:oMath"))

    def test_new_constructs_survive_a_docx_save_and_reopen(self):
        document = Document()
        formulas = [
            r"\begin{aligned} a &= b \\ c &= d \end{aligned}",
            r"\mathbb{R}",
            r"\textbf{v}",
            r"\underbrace{a+b}_{n}",
            r"\color{red}{\frac{a}{b}}",
            r"\phantom{x}",
        ]
        for offset, formula in enumerate(formulas):
            document.element.body.insert(offset, latex_to_omml(formula))
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        body = Document(buffer).element.body
        lines = body.findall("m:oMath/m:eqArr/m:e", MATH_NS)
        self.assertEqual([_aligned_runs(line) for line in lines], [["="], ["="]])
        self.assertEqual(
            _val(body, "m:oMath/m:r/m:rPr/m:scr"), "double-struck")
        self.assertIsNotNone(body.find(".//" + qn("w:b")))
        self.assertEqual(
            _val(body, "m:oMath/m:limLow/m:e/m:groupChr/m:groupChrPr/m:chr"), "⏟")
        self.assertIsNotNone(
            body.find("m:oMath/m:f/m:fPr/m:ctrlPr", MATH_NS))
        self.assertEqual(_val(body, "m:oMath/m:phant/m:phantPr/m:show"), "0")


if __name__ == "__main__":
    unittest.main()
