from pathlib import Path
import unittest

from docx.shared import Pt

from gfm_docx_renderer import GfmDocxRenderer


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


if __name__ == "__main__":
    unittest.main()
