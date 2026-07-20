import unittest
from unittest.mock import patch
import tempfile
from pathlib import Path

from PyQt6.QtCore import QMimeData, QPoint, QPointF, QSettings, QUrl, Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

from gui_theme import ThemeManager

from md_to_word_converter import ConverterGUI, DropFileList


class DropFileListTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_accepts_local_file_and_directory_urls(self):
        widget = DropFileList()
        mime = QMimeData()
        mime.setUrls([
            QUrl.fromLocalFile("/tmp/input.md"),
            QUrl.fromLocalFile("/tmp/folder"),
        ])
        event = QDragEnterEvent(
            QPoint(0, 0), Qt.DropAction.CopyAction, mime,
            Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
        )

        widget.dragEnterEvent(event)

        self.assertTrue(event.isAccepted())

    def test_emits_dropped_local_paths(self):
        widget = DropFileList()
        received = []
        widget.paths_dropped.connect(received.extend)
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile("/tmp/input.md")])
        event = QDropEvent(
            QPointF(0, 0), Qt.DropAction.CopyAction, mime,
            Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
        )

        widget.dropEvent(event)

        self.assertEqual(received, ["/tmp/input.md"])

    def test_window_accepts_dropped_files(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            md_file = Path(directory) / "input.md"
            md_file.write_text("# hi", encoding="utf-8")

            mime = QMimeData()
            mime.setUrls([QUrl.fromLocalFile(str(md_file))])
            enter = QDragEnterEvent(
                QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            )
            window.dragEnterEvent(enter)
            self.assertTrue(enter.isAccepted())

            drop = QDropEvent(
                QPointF(10, 10), Qt.DropAction.CopyAction, mime,
                Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            )
            window.dropEvent(drop)
            self.assertEqual(window.selected_files, [md_file.resolve()])

    def test_theme_button_switches_application_theme(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertEqual(window.theme_manager.theme, "dark")
            window.theme_button.click()

            self.assertEqual(window.theme_manager.theme, "light")
            self.assertEqual(window.theme_button.text(), "☾")
            self.assertIn("#F6F8FC", self.app.styleSheet())

    def test_drop_target_has_room_for_dragging(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertGreaterEqual(window.drop_hint.minimumHeight(), 80)

    def test_window_fits_small_screens(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            window.show()
            QApplication.processEvents()

            # Smallest supported screen (13" MacBook Air, 956pt) leaves ~900pt
            # usable after the menu bar; 800 keeps headroom while still
            # failing decisively on the 944px regression this guard catches.
            self.assertLessEqual(window.minimumSizeHint().height(), 800)

            window.hide()

    def test_drop_zone_click_requests_files(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            received = []
            window.drop_hint.clicked.connect(lambda: received.append(True))

            # QFileDialog.exec() blocks waiting for a modal close event that
            # never arrives under QT_QPA_PLATFORM=offscreen (verified: this
            # hangs even with DontUseNativeDialog), so the dialog call itself
            # is mocked here; the click -> clicked signal -> _select_files
            # wiring under test is untouched.
            with patch(
                "md_to_word_converter.QFileDialog.getOpenFileNames",
                return_value=([], ""),
            ):
                QTest.mouseClick(window.drop_hint, Qt.MouseButton.LeftButton)

            self.assertEqual(received, [True])

    def test_progress_indicator_is_hidden_while_idle(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertTrue(window.progress.isHidden())

    def test_text_tab_hidden_in_word_mode(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            text_index = window.tabs.indexOf(window.text_tab)
            window.tabs.setCurrentIndex(text_index)

            window.toggle_button.click()

            self.assertFalse(window.tabs.isTabVisible(text_index))
            self.assertNotEqual(window.tabs.currentIndex(), text_index)

            window.toggle_button.click()
            self.assertTrue(window.tabs.isTabVisible(text_index))

    def test_card_title_row_has_uniform_background(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            window.show()
            QApplication.processEvents()

            image = window.grab().toImage()
            group = window.output_group
            origin = group.mapTo(window, QPoint(0, 0))
            left = image.pixelColor(origin.x() + 40, origin.y() + 4)
            right = image.pixelColor(
                origin.x() + group.width() - 40, origin.y() + 4
            )

            self.assertEqual(left.name(), right.name())
            self.assertEqual(left.name(), "#161b22")

            window.hide()

    def test_progress_reserves_space_while_hidden(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertTrue(window.progress.isHidden())
            self.assertTrue(window.progress.sizePolicy().retainSizeWhenHidden())

    def test_queue_buttons_follow_queue_state(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            md_file = Path(directory) / "input.md"
            md_file.write_text("# hi", encoding="utf-8")

            self.assertFalse(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())

            window._add_sources([str(md_file)])
            self.assertTrue(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())
            self.assertIn("В очереди: 1", window.status_label.text())

            window.files_listbox.setCurrentRow(0)
            self.assertTrue(window.remove_button.isEnabled())

            window._clear_files()
            self.assertFalse(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())


if __name__ == "__main__":
    unittest.main()
