import tempfile
import unittest
from pathlib import Path

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication

from mdtoword.app import ConverterGUI, _describe_warning
from mdtoword.errors import ConversionWarning
from mdtoword.theme import ThemeManager


class GuiDocumentOptionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _window(self, directory: str) -> ConverterGUI:
        settings = QSettings(str(Path(directory) / "theme.ini"), QSettings.Format.IniFormat)
        return ConverterGUI(theme_manager=ThemeManager(settings=settings))

    def test_gost_preset_reaches_the_converter_and_bumps_the_default_size(self):
        with tempfile.TemporaryDirectory() as directory:
            window = self._window(directory)
            window.preset_combobox.setCurrentIndex(1)

            self.assertEqual(window.converter.document_options.preset, "gost")
            self.assertEqual(window.size_spinbox.value(), 14)
            self.assertEqual(window.converter.default_font_size.pt, 14)

            window.preset_combobox.setCurrentIndex(0)
            self.assertEqual(window.converter.document_options.preset, "default")
            self.assertEqual(window.size_spinbox.value(), 12)

    def test_custom_size_survives_a_preset_change(self):
        with tempfile.TemporaryDirectory() as directory:
            window = self._window(directory)
            window.size_spinbox.setValue(16)
            window.preset_combobox.setCurrentIndex(1)
            self.assertEqual(window.size_spinbox.value(), 16)

    def test_toc_checkbox_and_options_survive_a_mode_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            window = self._window(directory)
            window.toc_checkbox.setChecked(True)
            window._toggle_converter_type()
            window._toggle_converter_type()
            self.assertTrue(window.converter.document_options.toc)

    def test_warning_description_carries_the_source_line(self):
        warning = ConversionWarning("Image not found: a.png", code="image_not_found", line=12)
        self.assertEqual(_describe_warning("doc.md", warning), "doc.md:12: Image not found: a.png")
        self.assertEqual(_describe_warning("doc.md", "plain"), "doc.md: plain")


if __name__ == "__main__":
    unittest.main()
