import tempfile
import unittest
from pathlib import Path

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication

from gui_theme import ThemeManager


class ThemeManagerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        app = QApplication.instance() or QApplication([])
        assert isinstance(app, QApplication)
        cls.app: QApplication = app

    def setUp(self):
        self._tmpdir = tempfile.TemporaryDirectory()
        self.settings = QSettings(
            str(Path(self._tmpdir.name) / "theme.ini"),
            QSettings.Format.IniFormat,
        )

    def tearDown(self):
        self._tmpdir.cleanup()

    def test_defaults_to_dark_and_styles_tab_pane(self):
        manager = ThemeManager(settings=self.settings)

        self.assertEqual(manager.theme, "dark")
        stylesheet = manager.stylesheet("dark")
        self.assertIn("QTabWidget::pane", stylesheet)
        self.assertIn("QComboBox QAbstractItemView", stylesheet)
        self.assertIn("QScrollBar:vertical", stylesheet)
        self.assertIn("#0D1117", stylesheet)

    def test_toggle_applies_and_persists_theme(self):
        manager = ThemeManager(settings=self.settings)

        self.assertEqual(manager.toggle(), "light")
        manager.apply(self.app)

        restored = ThemeManager(settings=self.settings)
        self.assertEqual(restored.theme, "light")
        self.assertIn("#F6F8FC", self.app.styleSheet())


if __name__ == "__main__":
    unittest.main()
