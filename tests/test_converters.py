"""Тесты ядра конвертации.

Модуль намеренно не импортирует ``mdtoword.app``: ядро не должно зависеть
от PyQt6, и этот файл — исполняемое доказательство того, что не зависит.
"""

from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from docx import Document

from mdtoword.converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)


class MarkdownToWordConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    def test_convert_content_writes_the_document_and_returns_no_warnings(self) -> None:
        output = self.root / "out.docx"

        warnings = MarkdownToWordConverter().convert_content("# Заголовок", output)

        self.assertEqual(warnings, [])
        self.assertEqual(Document(str(output)).paragraphs[0].text, "Заголовок")

    def test_convert_file_returns_nonfatal_renderer_warnings_as_a_list(self) -> None:
        source = self.root / "source.md"
        source.write_text("![diagram](missing.png)", encoding="utf-8")
        output = self.root / "source.docx"

        warnings = MarkdownToWordConverter().convert_file(source, output)

        self.assertEqual(warnings, ["Image not found: missing.png"])
        self.assertTrue(output.is_file())

    def test_missing_source_raises_conversion_error(self) -> None:
        with self.assertRaises(ConversionError):
            MarkdownToWordConverter().convert_file(
                self.root / "нет-такого.md", self.root / "out.docx"
            )

    def test_unwritable_output_raises_conversion_error(self) -> None:
        source = self.root / "source.md"
        source.write_text("# Заголовок", encoding="utf-8")
        # Каталог вместо файла: python-docx не может писать поверх директории.
        blocked = self.root / "blocked.docx"
        blocked.mkdir()

        with self.assertRaises(ConversionError):
            MarkdownToWordConverter().convert_file(source, blocked)

    def test_font_settings_reach_the_rendered_document(self) -> None:
        output = self.root / "styled.docx"

        MarkdownToWordConverter(font_name="Georgia").convert_content("Текст", output)

        document = Document(str(output))
        self.assertEqual(document.styles["Normal"].font.name, "Georgia")


class WordToMarkdownConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    def test_headings_and_bold_runs_round_trip_into_markdown(self) -> None:
        docx_path = self.root / "source.docx"
        document = Document()
        document.add_heading("Раздел", level=2)
        paragraph = document.add_paragraph()
        paragraph.add_run("жирный").bold = True
        document.save(str(docx_path))
        output = self.root / "source.md"

        warnings = WordToMarkdownConverter().convert_file(docx_path, output)

        text = output.read_text(encoding="utf-8")
        self.assertEqual(warnings, [])
        self.assertIn("## Раздел", text)
        self.assertIn("**жирный**", text)

    def test_missing_source_raises_conversion_error(self) -> None:
        with self.assertRaises(ConversionError):
            WordToMarkdownConverter().convert_file(
                self.root / "нет-такого.docx", self.root / "out.md"
            )


class CoreIsolationTests(unittest.TestCase):
    def test_importing_the_core_does_not_pull_in_pyqt(self) -> None:
        # Обязательно в подпроцессе: в общем прогоне tests/test_drop_queue.py
        # импортирует PyQt6, и проверка sys.modules в текущем процессе
        # проходила бы или падала в зависимости от порядка сбора тестов.
        repo_root = Path(__file__).resolve().parent.parent
        result = subprocess.run(
            [
                sys.executable,
                "-c",
                "import mdtoword.converters, sys; "
                "sys.exit(1 if 'PyQt6' in sys.modules else 0)",
            ],
            capture_output=True,
            text=True,
            cwd=repo_root,
        )

        self.assertEqual(result.returncode, 0, "ядро подтянуло PyQt6:\n" + result.stderr)


if __name__ == "__main__":
    unittest.main()
