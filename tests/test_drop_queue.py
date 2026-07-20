import unittest
import tempfile
from pathlib import Path

from PyQt6.QtCore import QMimeData, QPoint, QPointF, QSettings, QUrl, Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
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

            self.assertGreaterEqual(window.drop_hint.minimumHeight(), 100)

    def test_progress_indicator_is_hidden_while_idle(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertTrue(window.progress.isHidden())


if __name__ == "__main__":
    unittest.main()
