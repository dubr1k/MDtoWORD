"""Markdown fidelity: lists, footnotes, images, metadata, anchors, HTML, layout.

Each test names one behaviour the Word output must have and checks it on
the document model (and, where the property lives in XML only, on the XML).
"""

from io import BytesIO
from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image

from mdtoword.gfm_renderer import GfmDocxRenderer
from mdtoword.options import DocumentOptions


_MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _render(markdown, source_path=None, **options):
    renderer = GfmDocxRenderer(
        "Times New Roman", Pt(12), document_options=DocumentOptions(**options)
    )
    return renderer.render(markdown, source_path=source_path)


def _num(paragraph):
    """(numId, ilvl) of a numbered paragraph, or None."""
    ppr = paragraph._p.pPr
    if ppr is None or ppr.numPr is None:
        return None
    return int(ppr.numPr.numId.val), int(ppr.numPr.ilvl.val)


def _start_of(document, num_id):
    """The start value Word will use for level 0 of *num_id*."""
    numbering = document.part.numbering_part.element
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != str(num_id):
            continue
        override = num.find(f"{qn('w:lvlOverride')}/{qn('w:startOverride')}")
        if override is not None:
            return int(override.get(qn("w:val")))
        abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
        for abstract in numbering.findall(qn("w:abstractNum")):
            if abstract.get(qn("w:abstractNumId")) == abstract_id:
                start = abstract.find(f"{qn('w:lvl')}/{qn('w:start')}")
                return int(start.get(qn("w:val"))) if start is not None else 1
    return None


def _png(width, height):
    buffer = BytesIO()
    Image.new("RGB", (width, height), (200, 30, 30)).save(buffer, format="PNG")
    return buffer.getvalue()


class ListTests(unittest.TestCase):
    def test_each_ordered_list_restarts_and_honours_its_start(self):
        document, warnings = _render(
            "1. one\n2. two\n\nText between.\n\n1. again one\n\nAnd:\n\n5. five\n6. six\n"
        )
        numbered = [p for p in document.paragraphs if _num(p)]
        self.assertEqual([p.text for p in numbered], ["one", "two", "again one", "five", "six"])
        first, second, third = _num(numbered[0])[0], _num(numbered[2])[0], _num(numbered[3])[0]
        self.assertEqual(_num(numbered[1])[0], first)
        self.assertNotEqual(first, second)
        self.assertNotEqual(second, third)
        self.assertEqual(_start_of(document, first), 1)
        self.assertEqual(_start_of(document, second), 1)
        self.assertEqual(_start_of(document, third), 5)
        self.assertEqual(warnings, [])

    def test_nested_lists_use_word_levels(self):
        document, _ = _render("1. top\n   - nested bullet\n     1. deep\n2. back\n")
        levels = {p.text: _num(p)[1] for p in document.paragraphs if _num(p)}
        self.assertEqual(levels, {"top": 0, "nested bullet": 1, "deep": 2, "back": 0})

    def test_second_paragraph_of_an_item_is_not_numbered(self):
        document, _ = _render("1. first\n\n   continuation\n\n   ```\n   code\n   ```\n2. second\n")
        texts = {p.text: p for p in document.paragraphs}
        self.assertIsNotNone(_num(texts["first"]))
        self.assertIsNone(_num(texts["continuation"]))
        self.assertIsNone(_num(texts["code"]))
        self.assertGreater(texts["continuation"].paragraph_format.left_indent, 0)
        self.assertGreater(texts["code"].paragraph_format.left_indent, 0)
        self.assertEqual(_num(texts["second"])[0], _num(texts["first"])[0])

    def test_item_starting_with_a_code_block_still_gets_its_number(self):
        document, _ = _render("1. ```\n   code\n   ```\n")
        numbered = [p for p in document.paragraphs if _num(p)]
        self.assertEqual(len(numbered), 1)

    def test_task_list_glyphs(self):
        document, _ = _render("- [ ] todo\n- [x] done\n")
        self.assertEqual([p.text for p in document.paragraphs], ["☐ todo", "☒ done"])


class LineBreakTests(unittest.TestCase):
    def test_single_newline_is_a_space_by_default(self):
        document, _ = _render("first line\nsecond line\n")
        self.assertEqual(document.paragraphs[0].text, "first line second line")

    def test_preserve_mode_keeps_line_breaks(self):
        document, _ = _render("first line\nsecond line\n", line_breaks="preserve")
        self.assertEqual(document.paragraphs[0].text, "first line\nsecond line")

    def test_hard_breaks_always_break(self):
        document, _ = _render("two spaces  \nbackslash\\\nend\n")
        self.assertEqual(document.paragraphs[0].text, "two spaces\nbackslash\nend")

    def test_justified_lines_ending_in_a_break_are_not_stretched(self):
        document, _ = _render("x\n")
        compat = document.settings.element.find(qn("w:compat"))
        self.assertIsNotNone(compat.find(qn("w:doNotExpandShiftReturn")))


class FrontMatterTests(unittest.TestCase):
    def test_front_matter_becomes_title_block_and_properties(self):
        document, warnings = _render(
            "---\ntitle: Отчёт\nauthor: [Иванов, Петров]\ntags: [a, b]\nlang: ru\n---\n\nТекст.\n"
        )
        self.assertEqual(document.paragraphs[0].style.name, "Title")
        self.assertEqual(document.paragraphs[0].text, "Отчёт")
        self.assertEqual(document.core_properties.title, "Отчёт")
        self.assertEqual(document.core_properties.author, "Иванов, Петров")
        self.assertEqual(document.core_properties.keywords, "a, b")
        texts = [p.text for p in document.paragraphs]
        self.assertNotIn("title: Отчёт", "\n".join(texts))
        self.assertEqual(warnings, [])


class AnchorTests(unittest.TestCase):
    def test_heading_links_point_at_bookmarks(self):
        document, warnings = _render(
            "См. [таблицы](#таблицы) и [дубль](#раздел-1).\n\n# Таблицы\n\n# Раздел\n\n# Раздел\n"
        )
        xml = document.element.body.xml
        bookmarks = document.element.body.findall(".//" + qn("w:bookmarkStart"))
        names = [bookmark.get(qn("w:name")) for bookmark in bookmarks]
        self.assertEqual(len(names), 3)
        links = document.paragraphs[0]._p.findall(qn("w:hyperlink"))
        anchors = [link.get(qn("w:anchor")) for link in links]
        self.assertEqual(anchors, [names[0], names[2]])
        self.assertNotIn('r:id=""', xml)
        self.assertEqual(warnings, [])

    def test_missing_anchor_warns_and_keeps_the_text(self):
        document, warnings = _render("[нет](#nowhere)\n")
        self.assertEqual(document.paragraphs[0].text, "нет")
        self.assertEqual([w.code for w in warnings], ["link_anchor_missing"])


class ImageTests(unittest.TestCase):
    def test_large_image_is_scaled_to_the_text_width(self):
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "big.png").write_bytes(_png(3000, 600))
            document, warnings = _render("![Схема](big.png)\n", source_path=Path(directory) / "a.md")
        shape = document.inline_shapes[0]
        section = document.sections[0]
        text_width = section.page_width - section.left_margin - section.right_margin
        self.assertLessEqual(shape.width, text_width)
        self.assertAlmostEqual(shape.width / shape.height, 5.0, places=2)
        self.assertEqual(warnings, [])

    def test_standalone_image_is_a_centred_numbered_figure(self):
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "a.png").write_bytes(_png(10, 10))
            document, _ = _render(
                "Текст на русском.\n\n![Схема системы](a.png)\n", source_path=Path(directory) / "a.md"
            )
        figure = [p for p in document.paragraphs if p._p.findall(".//" + qn("w:drawing"))][0]
        self.assertEqual(figure.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        captions = [p for p in document.paragraphs if p.style.name == "Caption"]
        self.assertEqual(len(captions), 1)
        self.assertTrue(captions[0].text.startswith("Рисунок "))
        self.assertTrue(captions[0].text.endswith("Схема системы"))
        self.assertIn("SEQ Figure", captions[0]._p.xml)
        self.assertIn('descr="Схема системы"', figure._p.xml)

    def test_inline_image_gets_no_caption(self):
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "a.png").write_bytes(_png(10, 10))
            document, _ = _render("Иконка ![i](a.png) в тексте.\n", source_path=Path(directory) / "a.md")
        self.assertEqual([p for p in document.paragraphs if p.style.name == "Caption"], [])

    def test_svg_is_embedded_with_a_png_fallback(self):
        svg = b'<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><rect width="200" height="100"/></svg>'
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "icon.svg").write_bytes(svg)
            document, warnings = _render("![](icon.svg)\n", source_path=Path(directory) / "a.md")
        self.assertEqual(warnings, [])
        self.assertIn("svgBlip", document.element.body.xml)

    def test_webp_is_converted(self):
        buffer = BytesIO()
        Image.new("RGB", (20, 10), (0, 0, 255)).save(buffer, format="WEBP")
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "x.webp").write_bytes(buffer.getvalue())
            document, warnings = _render("![x](x.webp)\n", source_path=Path(directory) / "a.md")
        self.assertEqual(warnings, [])
        self.assertEqual(len(document.inline_shapes), 1)

    def test_percent_encoded_cyrillic_file_name_is_found(self):
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "схема.png").write_bytes(_png(10, 10))
            document, warnings = _render("![x](схема.png)\n", source_path=Path(directory) / "a.md")
        self.assertEqual(warnings, [])
        self.assertEqual(len(document.inline_shapes), 1)


class InlineExtensionTests(unittest.TestCase):
    def test_sub_sup_mark(self):
        document, warnings = _render("H~2~O, x^2^ и ==важно==.\n")
        runs = document.paragraphs[0].runs
        self.assertTrue(any(run.font.subscript and run.text == "2" for run in runs))
        self.assertTrue(any(run.font.superscript and run.text == "2" for run in runs))
        self.assertTrue(any(run.font.highlight_color is not None and run.text == "важно" for run in runs))
        self.assertEqual(warnings, [])

    def test_inline_html_formatting_tags(self):
        document, warnings = _render("a<sub>1</sub> b<sup>2</sup> <kbd>Ctrl</kbd> <mark>m</mark> x<br>y\n")
        paragraph = document.paragraphs[0]
        self.assertNotIn("<", paragraph.text)
        runs = paragraph.runs
        self.assertTrue(any(run.font.subscript and run.text == "1" for run in runs))
        self.assertTrue(any(run.font.superscript and run.text == "2" for run in runs))
        self.assertTrue(any(run.font.name == "Courier New" and run.text == "Ctrl" for run in runs))
        self.assertIn("\n", paragraph.text)
        self.assertEqual(warnings, [])

    def test_unknown_html_is_stripped_with_one_warning(self):
        document, warnings = _render("<blink>x</blink> <blink>y</blink>\n")
        self.assertEqual(document.paragraphs[0].text, "x y")
        self.assertEqual([w.code for w in warnings], ["html_dropped"])

    def test_html_comment_is_dropped_silently(self):
        document, warnings = _render("a <!-- hidden --> b\n\n<!-- block comment -->\n")
        self.assertEqual(document.paragraphs[0].text, "a  b")
        self.assertEqual(warnings, [])

    def test_details_block_keeps_its_text(self):
        document, _ = _render("<details>\n<summary>Спойлер</summary>\nСкрытый текст\n</details>\n")
        texts = [p.text.strip() for p in document.paragraphs if p.text.strip()]
        self.assertEqual(texts, ["Спойлер", "Скрытый текст"])
        summary = [p for p in document.paragraphs if p.text.strip() == "Спойлер"][0]
        self.assertTrue(all(run.bold for run in summary.runs if run.text.strip()))

    def test_inline_code_in_heading_scales_with_the_heading(self):
        document, _ = _render("# Title `code`\n")
        code_run = [run for run in document.paragraphs[0].runs if run.text == "code"][0]
        self.assertGreater(code_run.font.size.pt, 12)


class BlockTests(unittest.TestCase):
    def test_github_alert_becomes_a_callout(self):
        document, warnings = _render("> [!WARNING]\n> Осторожно с этим.\n")
        texts = [p.text for p in document.paragraphs]
        self.assertEqual(texts, ["Внимание", "Осторожно с этим."])
        self.assertIn("w:shd", document.paragraphs[1]._p.xml)
        self.assertEqual(warnings, [])

    def test_nested_quotes_indent_deeper(self):
        document, _ = _render("> outer\n>\n> > inner\n")
        outer, inner = document.paragraphs
        self.assertGreater(inner.paragraph_format.left_indent, outer.paragraph_format.left_indent)

    def test_list_inside_quote_keeps_the_quote_indent(self):
        document, _ = _render("> - item\n")
        paragraph = document.paragraphs[0]
        self.assertIsNotNone(_num(paragraph))
        self.assertIn("w:pBdr", paragraph._p.xml)

    def test_definition_list(self):
        document, _ = _render("Term\n: Definition text\n")
        term, definition = document.paragraphs
        self.assertEqual(term.text, "Term")
        self.assertTrue(all(run.bold for run in term.runs))
        self.assertGreater(definition.paragraph_format.left_indent, 0)

    def test_code_block_uses_the_source_code_style(self):
        document, _ = _render("```python\nx = 1\n```\n")
        self.assertEqual([p.style.name for p in document.paragraphs], ["Source Code"])
        self.assertEqual(document.paragraphs[0].text, "x = 1")

    def test_table_caption_goes_above_the_table(self):
        document, warnings = _render("| a |\n|---|\n| 1 |\n\nTable: Results\n")
        body = list(document.element.body)
        caption_index = next(
            i for i, element in enumerate(body) if element.tag == qn("w:p") and "Results" in "".join(element.itertext())
        )
        table_index = next(i for i, element in enumerate(body) if element.tag == qn("w:tbl"))
        self.assertLess(caption_index, table_index)
        self.assertEqual(warnings, [])

    def test_toc_marker_inserts_a_toc_field(self):
        document, _ = _render("[TOC]\n\n# One\n\n## Two\n")
        xml = document.element.body.xml
        self.assertIn("TOC \\o", xml)
        self.assertIsNotNone(document.settings.element.find(qn("w:updateFields")))

    def test_numbered_display_equation_via_tag_or_label(self):
        document, warnings = _render("$$\nx = 1\n$$ (3)\n")
        labelled = [p for p in document.paragraphs if "(3)" in p.text]
        self.assertEqual(len(labelled), 1)
        self.assertEqual(warnings, [])

    def test_tag_numbers_display_and_amsmath_equations(self):
        cases = {
            "$$x = 1 \\tag{2}$$\n": "(2)",
            "$$x = 1 \\tag*{A}$$\n": "A",
            "\\begin{align}\na &= b \\tag{4}\n\\end{align}\n": "(4)",
        }
        for markdown, number in cases.items():
            with self.subTest(markdown=markdown):
                document, warnings = _render(markdown)
                paragraph = document.paragraphs[0]
                self.assertEqual(paragraph.text.strip(), number)
                self.assertEqual(len(list(paragraph._p.iter(f"{_MATH_NS}oMath"))), 1)
                self.assertNotIn("tag", "".join(paragraph._p.itertext()))
                self.assertEqual(warnings, [])

    def test_display_equation_is_wrapped_in_omathpara(self):
        document, _ = _render("$$\\sum_{i=1}^n i$$\n")
        self.assertIn("oMathPara", document.element.body.xml)


class LayoutTests(unittest.TestCase):
    def test_a4_by_default(self):
        document, _ = _render("x\n")
        section = document.sections[0]
        self.assertAlmostEqual(section.page_width.mm, 210, places=0)
        self.assertAlmostEqual(section.page_height.mm, 297, places=0)

    def test_letter_on_request(self):
        document, _ = _render("x\n", page_size="Letter")
        self.assertAlmostEqual(document.sections[0].page_width.mm, 215.9, places=0)

    def test_language_is_detected(self):
        document, _ = _render("Это русский текст документа.\n")
        lang = document.styles.element.find(f"{qn('w:docDefaults')}//{qn('w:lang')}")
        self.assertEqual(lang.get(qn("w:val")), "ru-RU")

    def test_gost_preset(self):
        renderer = GfmDocxRenderer(
            "Times New Roman", Pt(14), document_options=DocumentOptions(preset="gost")
        )
        document, _ = renderer.render("Абзац текста.\n\n![Схема](missing.png)\n")
        section = document.sections[0]
        self.assertAlmostEqual(section.left_margin.mm, 30, places=0)
        self.assertAlmostEqual(section.right_margin.mm, 15, places=0)
        normal = document.styles["Normal"].paragraph_format
        self.assertEqual(normal.line_spacing, 1.5)
        self.assertAlmostEqual(normal.first_line_indent.cm, 1.25, places=2)
        self.assertIn("PAGE", section.footer.paragraphs[0]._p.xml)

    def test_template_supplies_styles_but_not_content(self):
        with tempfile.TemporaryDirectory() as directory:
            template = Document()
            template.add_paragraph("TEMPLATE BODY TEXT")
            template.styles["Normal"].font.name = "Georgia"
            path = Path(directory) / "template.docx"
            template.save(str(path))
            document, _ = _render("Hello.\n", template=path)
        self.assertNotIn("TEMPLATE BODY TEXT", "\n".join(p.text for p in document.paragraphs))
        self.assertEqual(document.styles["Normal"].font.name, "Georgia")
        self.assertIsNone(document.paragraphs[0].runs[0].font.name)

    def test_missing_template_fails_clearly(self):
        with self.assertRaises(FileNotFoundError):
            _render("x\n", template=Path("/nonexistent/template.docx"))


class WarningMetadataTests(unittest.TestCase):
    def test_warnings_carry_code_and_line(self):
        _, warnings = _render("Intro.\n\nSecond paragraph\nwith $\\qedsymbol$ on line four.\n")
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0].code, "formula_unsupported")
        self.assertEqual(warnings[0].line, 4)

    def test_unreferenced_footnote_is_reported_with_its_line(self):
        _, warnings = _render("text[^a]\n\n[^a]: used\n[^x]: orphan\n")
        self.assertEqual([(w.code, w.line) for w in warnings], [("footnote_unreferenced", 4)])

    def test_document_saves_and_reopens(self):
        document, _ = _render(
            "---\ntitle: T\n---\n\n[TOC]\n\n# A\n\n1. x[^1]\n\n> [!NOTE]\n> n\n\n| a |\n|---|\n| $x$ |\n\n[^1]: note\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "out.docx"
            document.save(str(path))
            reopened = Document(str(path))
        self.assertTrue(reopened.paragraphs)


if __name__ == "__main__":
    unittest.main()
