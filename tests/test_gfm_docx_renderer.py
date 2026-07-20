from pathlib import Path
import tempfile
import unittest

from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.shared import RGBColor

from gfm_docx_renderer import GfmDocxRenderer


_MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _equations(paragraph):
    """Every Word equation directly inside this paragraph."""
    return paragraph._p.findall(f"{_MATH_NS}oMath")


class GfmDocxRendererTests(unittest.TestCase):
    def setUp(self):
        self.renderer = GfmDocxRenderer("Arial", Pt(12))

    def test_renders_gfm_blocks_and_inline_formatting(self):
        document, warnings = self.renderer.render(
            "# Heading\n\n"
            "Paragraph with **bold**, *italic*, ~~struck~~, and `code`.\n\n"
            "> quoted text\n\n"
            "- [x] done\n"
            "  - nested\n\n"
            "| left | right |\n"
            "| :--- | ---: |\n"
            "| one | two |\n\n"
            "```python\n"
            "print('x')\n"
            "```"
        )

        self.assertEqual(document.paragraphs[0].style.name, "Heading 1")
        inline_paragraph = document.paragraphs[1]
        self.assertTrue(any(run.bold for run in inline_paragraph.runs))
        self.assertTrue(any(run.italic for run in inline_paragraph.runs))
        self.assertTrue(any(run.font.strike for run in inline_paragraph.runs))
        self.assertIn("quoted text", "\n".join(p.text for p in document.paragraphs))
        self.assertIn("☒ done", "\n".join(p.text for p in document.paragraphs))
        self.assertEqual(len(document.tables), 1)
        self.assertIn("print('x')", "\n".join(p.text for p in document.paragraphs))
        self.assertEqual(warnings, [])

    def test_renders_links_and_falls_back_when_local_image_is_missing(self):
        document, warnings = self.renderer.render(
            "[MDtoWord](https://example.com) and ![diagram](missing.png)",
            source_path=Path("/tmp/source.md"),
        )

        self.assertIn("MDtoWord", document.paragraphs[0].text)
        self.assertIn("[diagram]", document.paragraphs[0].text)
        self.assertEqual(warnings, ["Image not found: missing.png"])

    def test_renders_footnotes_as_a_final_section(self):
        document, warnings = self.renderer.render(
            "Text with a footnote.[^1]\n\n[^1]: Footnote text."
        )

        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("Footnotes", text)
        self.assertIn("Footnote text.", text)
        self.assertEqual(warnings, [])

    def test_headings_are_black(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "# Заголовок\n\n## Второй\n\n### Третий\n"
        )
        for level in (1, 2, 3):
            style = document.styles[f"Heading {level}"]
            self.assertEqual(style.font.color.rgb, RGBColor(0, 0, 0))

    def test_hyperlink_is_black_and_underlined(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "[сайт](https://example.com)\n"
        )
        xml = document.paragraphs[0]._p.xml
        self.assertIn('w:val="000000"', xml)
        self.assertNotIn("0563C1", xml)
        self.assertIn("w:u", xml)

    def test_body_paragraphs_are_justified(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "Обычный абзац текста.\n\n- пункт списка\n\n> цитата\n"
        )
        justified = [
            p for p in document.paragraphs
            if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
        ]
        self.assertEqual(len(justified), 3)

    def test_headings_and_code_are_not_justified(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "# Заголовок\n\n```python\nx = 1\n```\n"
        )
        for paragraph in document.paragraphs:
            self.assertNotEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)

    def test_quote_style_color_is_explicit_black_not_theme_color(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "> цитата\n"
        )
        style = document.styles["Quote"]
        self.assertEqual(style.font.color.rgb, RGBColor(0, 0, 0))
        self.assertNotIn("themeColor", style.element.xml)

    def test_table_has_explicit_borders(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "| a | b |\n|---|---|\n| 1 | 2 |\n"
        )
        table_xml = document.tables[0]._tbl.tblPr.xml
        self.assertIn("tblBorders", table_xml)
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            self.assertIn(f"w:{edge}", table_xml)
        tbl_pr = document.tables[0]._tbl.tblPr
        child_names = [child.tag.rsplit("}", 1)[-1] for child in tbl_pr]
        self.assertIn("tblBorders", child_names)
        self.assertIn("tblLook", child_names)
        self.assertLess(
            child_names.index("tblBorders"),
            child_names.index("tblLook"),
            "w:tblBorders must precede w:tblLook in the tblPr child sequence "
            "per the OOXML CT_TblPrBase schema order",
        )

    def test_table_respects_markdown_column_alignment(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "| left | right | center |\n|:---|---:|:---:|\n| a | b | c |\n"
        )
        body = document.tables[0].rows[1]
        self.assertEqual(body.cells[0].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(body.cells[1].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(body.cells[2].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_escaped_inline_math_becomes_equations_not_text(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            r"Формула $a\,b$ и $a*b*c$ и $50\%$ в строке."
        )
        paragraph = document.paragraphs[0]
        self.assertEqual(len(_equations(paragraph)), 3)
        self.assertEqual(paragraph.text, "Формула  и  и  в строке.")
        self.assertEqual(warnings, [])

    def test_dollar_amounts_in_prose_survive_verbatim(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "This costs $5 and that costs $10, total $15."
        )
        text = document.paragraphs[0].text
        self.assertEqual(
            text, "This costs $5 and that costs $10, total $15."
        )
        self.assertEqual(warnings, [])

    def test_display_math_becomes_a_structured_equation(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "$$\n\\int_0^\\infty e^{-x^2}\\,dx = \\frac{\\sqrt{\\pi}}{2}\n$$\n"
        )
        with_equations = [p for p in document.paragraphs if _equations(p)]
        self.assertEqual(len(with_equations), 1)
        xml = with_equations[0]._p.xml
        self.assertIn("<m:nary>", xml)  # the integral
        self.assertIn("<m:f>", xml)  # the fraction
        self.assertIn("<m:rad>", xml)  # the square root
        self.assertEqual(warnings, [])

    def test_equation_label_is_not_dropped(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "$$\nx = 1\n$$ (1)\n"
        )
        labelled = [p for p in document.paragraphs if "(1)" in p.text]
        self.assertEqual(len(labelled), 1, "equation label was dropped")
        self.assertEqual(
            len(_equations(labelled[0])),
            1,
            "label paragraph lost its equation",
        )
        self.assertEqual(warnings, [])

    def test_amsmath_align_becomes_one_equation_per_line(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "\\begin{align}\n"
            "a &= b + c \\\\\n"
            "d &= e + f\n"
            "\\end{align}\n"
        )
        with_equations = [p for p in document.paragraphs if _equations(p)]
        self.assertEqual(len(with_equations), 2)
        for paragraph in with_equations:
            self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(len(_equations(paragraph)), 1)
        self.assertEqual(warnings, [])

    def test_amsmath_matrix_environment_is_a_single_equation(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "\\begin{pmatrix}\n"
            "a & b \\\\\n"
            "c & d\n"
            "\\end{pmatrix}\n"
        )
        with_equations = [p for p in document.paragraphs if _equations(p)]
        self.assertEqual(len(with_equations), 1)
        self.assertEqual(len(_equations(with_equations[0])), 1)
        self.assertIn("<m:m>", with_equations[0]._p.xml)
        self.assertEqual(warnings, [])

    def test_unsupported_display_environment_is_kept_byte_for_byte(self):
        source = (
            "\\begin{align}\n"
            "a &= \\qedsymbol{b} \\\\\n"
            "c &= d\n"
            "\\end{align}"
        )
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            source + "\n"
        )
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertEqual(text, source)
        self.assertEqual(
            [eq for p in document.paragraphs for eq in _equations(p)],
            [],
            "a partially converted environment must not leak equations",
        )
        self.assertEqual(len(warnings), 1)
        self.assertIn("qedsymbol", warnings[0])

    def test_whitespace_only_inline_math_keeps_its_literal_dollars(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "x $ $ y\n"
        )
        self.assertEqual(document.paragraphs[0].text, "x $ $ y")
        self.assertEqual(_equations(document.paragraphs[0]), [])
        self.assertEqual(warnings, [])

    def test_empty_display_math_does_not_crash_or_warn(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "before\n\n$$\n$$\n\nafter\n"
        )
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("before", text)
        self.assertIn("after", text)
        self.assertEqual(warnings, [])

    def test_blank_line_does_not_widen_display_math(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "Первый абзац.\n"
            "\n"
            "$$\n"
            "\n"
            "Второй абзац, ничего общего с формулой.\n"
            "\n"
            "$$\n"
            "\n"
            "Третий абзац.\n"
        )
        middle_text = "Второй абзац, ничего общего с формулой."
        matches = [p for p in document.paragraphs if middle_text in p.text]
        self.assertTrue(matches, "middle paragraph text did not survive verbatim")
        middle_paragraph = matches[0]
        self.assertTrue(middle_paragraph.runs, "middle paragraph has no runs")
        for run in middle_paragraph.runs:
            self.assertNotEqual(run.font.name, "Courier New")

    def test_stray_dollar_pair_with_cyrillic_content_warns(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "Переменные $HOME и $PATH в шелле."
        )
        self.assertEqual(len(warnings), 1)
        self.assertIn(r"\$", warnings[0])
        self.assertIn("HOME", warnings[0])

    def test_genuine_formula_does_not_warn(self):
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "$E = mc^2$"
        )
        self.assertEqual(warnings, [])

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

    def test_equations_survive_a_save_and_reopen_round_trip(self):
        unsupported = r"\qedsymbol{x}"
        document, warnings = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "Inline $E = mc^2$ here.\n\n"
            "$$\\frac{a}{b}$$\n\n"
            "\\begin{equation}\n"
            "F = G\\frac{m_1 m_2}{r^2}\n"
            "\\end{equation}\n\n"
            "\\begin{align}\n"
            "a &= b + c \\\\\n"
            "d &= e + f\n"
            "\\end{align}\n\n"
            f"Broken ${unsupported}$ formula.\n"
        )

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "round-trip.docx"
            document.save(str(path))
            reopened = Document(str(path))

        equations = [eq for p in reopened.paragraphs for eq in _equations(p)]
        # inline + display + equation + two align lines
        self.assertEqual(len(equations), 5)

        self.assertEqual(len(warnings), 1)
        self.assertIn("qedsymbol", warnings[0])

        text = "\n".join(p.text for p in reopened.paragraphs)
        self.assertIn(f"Broken {unsupported} formula.", text)

    def test_footnote_paragraphs_are_justified_like_body_lists(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "1. пункт списка\n\n"
            "Текст со сноской[^1]\n\n"
            "[^1]: Содержимое сноски.\n"
        )
        list_number_paragraphs = [
            paragraph
            for paragraph in document.paragraphs
            if paragraph.style.name == "List Number"
        ]
        self.assertEqual(len(list_number_paragraphs), 2)
        for paragraph in list_number_paragraphs:
            self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)


if __name__ == "__main__":
    unittest.main()
