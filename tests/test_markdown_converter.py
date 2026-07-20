from pathlib import Path
import tempfile
import unittest

from md_to_word_converter import MarkdownToWordConverter


class MarkdownConverterTests(unittest.TestCase):
    def test_convert_file_reports_nonfatal_renderer_warnings(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.md"
            output = Path(directory) / "source.docx"
            source.write_text("![diagram](missing.png)", encoding="utf-8")

            success, message = MarkdownToWordConverter().convert_file(source, output)

            self.assertTrue(success)
            self.assertTrue(output.is_file())
            self.assertIn("Image not found: missing.png", message)


if __name__ == "__main__":
    unittest.main()
