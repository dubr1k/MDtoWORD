import tempfile
import unittest
from pathlib import Path

from PyQt6.QtCore import QSettings
from PyQt6.QtGui import QPalette
from PyQt6.QtWidgets import QApplication

from mdtoword.theme import ThemeManager


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
        self.icon_cache_dir = Path(self._tmpdir.name) / "icons"

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
        manager = ThemeManager(settings=self.settings, icon_cache_dir=self.icon_cache_dir)

        self.assertEqual(manager.toggle(), "light")
        manager.apply(self.app)

        restored = ThemeManager(settings=self.settings)
        self.assertEqual(restored.theme, "light")
        self.assertIn("#F6F8FC", self.app.styleSheet())

    def test_apply_sets_fusion_palette(self):
        manager = ThemeManager(settings=self.settings, icon_cache_dir=self.icon_cache_dir)
        manager.apply(self.app)
        self.assertEqual(
            self.app.palette().color(QPalette.ColorRole.Window).name().upper(),
            "#0D1117",
        )

    def test_stylesheet_draws_themed_drop_down_and_spin_buttons(self):
        # Fusion's native arrows are what we're replacing, so the stylesheet
        # must now define real drop-down/spin-button subcontrols instead of
        # leaving them absent (the old assertion belonged to the "let Fusion
        # draw it" version of this fix).
        stylesheet = ThemeManager.stylesheet("dark")
        self.assertIn("QComboBox::drop-down", stylesheet)
        self.assertIn("subcontrol-origin: padding", stylesheet)
        self.assertIn("subcontrol-position: center right", stylesheet)
        self.assertIn("QSpinBox::up-button", stylesheet)
        self.assertIn("QSpinBox::down-button", stylesheet)
        self.assertIn("subcontrol-origin: border", stylesheet)

    def test_apply_writes_themed_chevron_icons(self):
        manager = ThemeManager(settings=self.settings, icon_cache_dir=self.icon_cache_dir)
        manager.apply(self.app)

        down_path = self.icon_cache_dir / "chevron-down-dark.svg"
        up_path = self.icon_cache_dir / "chevron-up-dark.svg"
        self.assertTrue(down_path.exists())
        self.assertTrue(up_path.exists())
        self.assertIn("#A9B1D6", down_path.read_text(encoding="utf-8"))
        self.assertIn("#A9B1D6", up_path.read_text(encoding="utf-8"))

        stylesheet = self.app.styleSheet()
        self.assertIn(down_path.as_posix(), stylesheet)
        self.assertIn(up_path.as_posix(), stylesheet)
        self.assertIn("image:", stylesheet)
        self.assertTrue(down_path.is_absolute())
        self.assertTrue(up_path.is_absolute())

    def test_switching_theme_points_stylesheet_at_the_other_icon_and_colour(self):
        manager = ThemeManager(settings=self.settings, icon_cache_dir=self.icon_cache_dir)
        manager.apply(self.app)
        dark_down_path = self.icon_cache_dir / "chevron-down-dark.svg"
        self.assertIn(dark_down_path.as_posix(), self.app.styleSheet())

        manager.toggle()
        manager.apply(self.app)
        light_down_path = self.icon_cache_dir / "chevron-down-light.svg"
        light_stylesheet = self.app.styleSheet()

        self.assertIn(light_down_path.as_posix(), light_stylesheet)
        self.assertNotIn(dark_down_path.as_posix(), light_stylesheet)
        self.assertIn("#5D687A", light_down_path.read_text(encoding="utf-8"))
        self.assertNotIn("#A9B1D6", light_down_path.read_text(encoding="utf-8"))

    def test_apply_survives_unwritable_icon_cache_and_omits_image_rule(self):
        # A regular file sitting where the cache directory needs to be makes
        # that directory literally uncreatable regardless of OS permission
        # bits or whether the test runs as root, so this reliably exercises
        # the fail-soft path without mocking anything away.
        blocked = Path(self._tmpdir.name) / "blocked-cache"
        blocked.write_text("not a directory", encoding="utf-8")

        manager = ThemeManager(settings=self.settings, icon_cache_dir=blocked)
        manager.apply(self.app)  # must not raise

        stylesheet = self.app.styleSheet()
        self.assertNotIn("image:", stylesheet)
        # The surrounding layout styling should still be present so the
        # controls look deliberate even without themed arrow icons.
        self.assertIn("QComboBox::drop-down", stylesheet)
        self.assertIn("QSpinBox::up-button", stylesheet)


if __name__ == "__main__":
    unittest.main()
