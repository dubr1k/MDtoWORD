import sys
from typing import Any, cast
from pathlib import Path

from docx.shared import Pt
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent, QIcon
from PyQt6.QtWidgets import (
    QApplication, QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QGroupBox,
    QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QSpinBox, QTabWidget,
    QVBoxLayout, QWidget,
)

from .converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)
from .gui.texts import TRANSLATIONS, describe_warning as _describe_warning
from .gui.widgets import (
    DropFileList,
    DropZoneLabel,
    accept_local_paths_event as _accept_local_paths_event,
    dropped_local_paths as _dropped_local_paths,
)
from .options import PRESET_FONT_SIZE, DocumentOptions
from .workflow import discover_sources, resolve_output_paths
from .theme import ThemeManager

__all__ = ["ConverterGUI", "DropFileList", "DropZoneLabel", "main"]


class ConverterGUI(QMainWindow):
    """Compact GUI for interactive conversion batches."""

    def __init__(self, theme_manager: ThemeManager | None = None):
        super().__init__()
        self.theme_manager = theme_manager or ThemeManager()
        app = QApplication.instance()
        if isinstance(app, QApplication):
            self.theme_manager.apply(app)
        self.setWindowTitle("MDtoWord")
        self.resize(860, 720)
        self.selected_files: list[Path] = []
        self.output_directory: Path | None = None
        self.current_converter_type = "md_to_word"
        self.converter: MarkdownToWordConverter | WordToMarkdownConverter = MarkdownToWordConverter()
        self.current_language = "ru"
        self.fonts = ["Arial", "Times New Roman", "Calibri", "Georgia", "Helvetica", "Courier New"]
        self.translations = TRANSLATIONS
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.footnotes_heading = self._text["footnotes_heading"]
        self._set_icon()
        self._create_widgets()
        self.setAcceptDrops(True)

    def _set_icon(self) -> None:
        module_dir = Path(__file__).resolve().parent
        # assets/ лежит в корне проекта — на уровень выше пакета mdtoword,
        # а в сборке PyInstaller распаковывается рядом с кодом в sys._MEIPASS.
        search_roots = [module_dir.parent, module_dir]
        bundle_root = getattr(sys, "_MEIPASS", None)
        if bundle_root:
            search_roots.insert(0, Path(bundle_root))
        for icon_name in (("macos-icon.png", "ico.png") if sys.platform == "darwin" else ("ico.png", "macos-icon.png")):
            for root in search_roots:
                icon_path = root / "assets" / icon_name
                if icon_path.exists():
                    self.setWindowIcon(QIcon(str(icon_path)))
                    return

    @property
    def _text(self) -> dict[str, str]:
        return self.translations[self.current_language]

    def dragEnterEvent(self, event: QDragEnterEvent | None) -> None:
        _accept_local_paths_event(event)

    def dragMoveEvent(self, event: QDragMoveEvent | None) -> None:
        _accept_local_paths_event(event)

    def dropEvent(self, event: QDropEvent | None) -> None:
        if event is None:
            return
        paths = _dropped_local_paths(event)
        if paths:
            self._add_sources(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _create_widgets(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 14, 20, 16)
        layout.setSpacing(10)

        self.title_label = QLabel()
        self.title_label.setObjectName("title-label")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title_label)

        self.settings_group = QGroupBox()
        settings = QHBoxLayout(self.settings_group)
        self.font_label = QLabel()
        self.font_combobox = QComboBox()
        self.font_combobox.addItems(self.fonts)
        self.font_combobox.setCurrentText("Times New Roman")
        self.font_combobox.currentTextChanged.connect(self._on_font_change)
        self.size_label = QLabel()
        self.size_spinbox = QSpinBox()
        self.size_spinbox.setRange(6, 72)
        self.size_spinbox.setValue(12)
        self.size_spinbox.valueChanged.connect(self._on_size_change)
        self.preset_label = QLabel()
        self.preset_combobox = QComboBox()
        self.preset_combobox.addItem("", "default")
        self.preset_combobox.addItem("", "gost")
        self.preset_combobox.currentIndexChanged.connect(self._on_preset_change)
        self.toc_checkbox = QCheckBox()
        self.toc_checkbox.toggled.connect(self._apply_document_options)
        settings.addWidget(self.font_label)
        settings.addWidget(self.font_combobox, 1)
        settings.addWidget(self.size_label)
        settings.addWidget(self.size_spinbox)
        settings.addWidget(self.preset_label)
        settings.addWidget(self.preset_combobox)
        settings.addWidget(self.toc_checkbox)
        layout.addWidget(self.settings_group)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)
        self.files_tab = QWidget()
        self.files_tab.setObjectName("tab-page")
        files_layout = QVBoxLayout(self.files_tab)
        files_layout.setContentsMargins(16, 16, 16, 16)
        files_layout.setSpacing(10)
        self.drop_hint = DropZoneLabel()
        self.drop_hint.clicked.connect(self._select_files)
        files_layout.addWidget(self.drop_hint)
        actions = QHBoxLayout()
        self.add_files_button = QPushButton()
        self.add_files_button.clicked.connect(self._select_files)
        self.add_folder_button = QPushButton()
        self.add_folder_button.clicked.connect(self._select_folder)
        actions.addWidget(self.add_files_button)
        actions.addWidget(self.add_folder_button)
        actions.addStretch()
        files_layout.addLayout(actions)
        self.files_listbox = DropFileList()
        self.files_listbox.paths_dropped.connect(self._add_sources)
        self.files_listbox.itemSelectionChanged.connect(self._update_queue_buttons)
        self.files_listbox.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.files_listbox.setMinimumHeight(90)
        files_layout.addWidget(self.files_listbox, 1)
        removal = QHBoxLayout()
        self.remove_button = QPushButton()
        self.remove_button.setObjectName("danger-button")
        self.remove_button.clicked.connect(self._remove_selected_files)
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("danger-button")
        self.clear_button.clicked.connect(self._clear_files)
        removal.addWidget(self.remove_button)
        removal.addWidget(self.clear_button)
        removal.addStretch()
        files_layout.addLayout(removal)
        self.tabs.addTab(self.files_tab, "")

        self.text_tab = QWidget()
        self.text_tab.setObjectName("tab-page")
        text_layout = QVBoxLayout(self.text_tab)
        self.text_label = QLabel()
        self.text_input = QPlainTextEdit()
        text_layout.addWidget(self.text_label)
        text_layout.addWidget(self.text_input)
        self.tabs.addTab(self.text_tab, "")

        self.output_group = QGroupBox()
        output = QHBoxLayout(self.output_group)
        self.output_label = QLabel()
        self.output_label.setObjectName("output-path")
        self.choose_output_button = QPushButton()
        self.choose_output_button.clicked.connect(self._select_output_directory)
        self.reset_output_button = QPushButton()
        self.reset_output_button.clicked.connect(self._reset_output_directory)
        output.addWidget(self.output_label, 1)
        output.addWidget(self.choose_output_button)
        output.addWidget(self.reset_output_button)
        layout.addWidget(self.output_group)

        self.progress = QProgressBar()
        policy = self.progress.sizePolicy()
        policy.setRetainSizeWhenHidden(True)
        self.progress.setSizePolicy(policy)
        self.progress.hide()
        layout.addWidget(self.progress)
        self.status_label = QLabel()
        self.status_label.setObjectName("status-label")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        self.convert_button = QPushButton()
        self.convert_button.setObjectName("primary-button")
        self.convert_button.clicked.connect(self._convert_files)
        layout.addWidget(self.convert_button)

        footer = QHBoxLayout()
        self.toggle_button = QPushButton()
        self.toggle_button.clicked.connect(self._toggle_converter_type)
        self.language_button = QPushButton("EN")
        self.language_button.clicked.connect(self._toggle_language)
        self.theme_button = QPushButton()
        self.theme_button.setObjectName("theme-button")
        self.theme_button.clicked.connect(self._toggle_theme)
        footer.addWidget(self.toggle_button)
        footer.addStretch()
        footer.addWidget(self.theme_button)
        footer.addWidget(self.language_button)
        self._update_theme_button()
        layout.addLayout(footer)
        self._update_ui()

    def _toggle_converter_type(self) -> None:
        self.current_converter_type = (
            "word_to_md" if self.current_converter_type == "md_to_word" else "md_to_word"
        )
        self.converter = (
            WordToMarkdownConverter() if self.current_converter_type == "word_to_md" else MarkdownToWordConverter()
        )
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.default_font_name = self.font_combobox.currentText()
            self.converter.default_font_size = Pt(self.size_spinbox.value())
            self.converter.footnotes_heading = self._text["footnotes_heading"]
            self._apply_document_options()
        self.selected_files = discover_sources(self.selected_files, self.current_converter_type)
        self._update_ui()

    def _toggle_language(self) -> None:
        self.current_language = "en" if self.current_language == "ru" else "ru"
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.footnotes_heading = self._text["footnotes_heading"]
        self._update_ui()

    def _toggle_theme(self) -> None:
        self.theme_manager.toggle()
        app = QApplication.instance()
        if isinstance(app, QApplication):
            self.theme_manager.apply(app)
        self._update_theme_button()

    def _update_theme_button(self) -> None:
        is_dark = self.theme_manager.theme == "dark"
        self.theme_button.setText("☀" if is_dark else "☾")
        tooltip = self._text["theme_dark" if is_dark else "theme_light"]
        self.theme_button.setToolTip(tooltip)
        self.theme_button.setAccessibleName(tooltip)

    def _update_ui(self) -> None:
        text = self._text
        is_markdown = self.current_converter_type == "md_to_word"
        self.setWindowTitle(text["title_md"] if is_markdown else text["title_word"])
        self.title_label.setText(self.windowTitle())
        self.settings_group.setTitle(text["settings"])
        self.settings_group.setVisible(is_markdown)
        self.font_label.setText(text["font"])
        self.size_label.setText(text["size"])
        self.preset_label.setText(text["preset"])
        self.preset_combobox.setItemText(0, text["preset_default"])
        self.preset_combobox.setItemText(1, text["preset_gost"])
        self.toc_checkbox.setText(text["toc"])
        self.drop_hint.setText(text["drop_md"] if is_markdown else text["drop_word"])
        self.add_files_button.setText(text["add_files"])
        self.add_folder_button.setText(text["add_folder"])
        self.remove_button.setText(text["remove"])
        self.clear_button.setText(text["clear"])
        self.tabs.setTabText(self.tabs.indexOf(self.files_tab), text["files_tab"])
        self.tabs.setTabText(self.tabs.indexOf(self.text_tab), text["text_tab"])
        self.text_label.setText(text["text_label"])
        self.tabs.setTabVisible(self.tabs.indexOf(self.text_tab), is_markdown)
        self.output_group.setTitle(text["output"])
        self.output_label.setText(str(self.output_directory) if self.output_directory else text["output_auto"])
        self.choose_output_button.setText(text["choose_output"])
        self.reset_output_button.setText(text["reset_output"])
        self.reset_output_button.setVisible(self.output_directory is not None)
        self.toggle_button.setText(text["toggle_md"] if is_markdown else text["toggle_word"])
        self.language_button.setText("EN" if self.current_language == "ru" else "RU")
        self._update_theme_button()
        self._refresh_queue()

    def _refresh_queue(self) -> None:
        self.files_listbox.clear()
        for source in self.selected_files:
            self.files_listbox.addItem(f"{source.name}\n{source.parent}")
        count = len(self.selected_files)
        if count:
            self.status_label.setText(self._text["queued"].format(count=count))
        else:
            self.status_label.setText(self._text["ready"])
        self.convert_button.setText(
            self._text["convert"] if not count else f"{self._text['convert']} ({count})"
        )
        self._update_queue_buttons()

    def _update_queue_buttons(self) -> None:
        self.clear_button.setEnabled(bool(self.selected_files))
        self.remove_button.setEnabled(bool(self.files_listbox.selectedItems()))

    def _on_font_change(self, font_name: str) -> None:
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.default_font_name = font_name

    def _on_size_change(self, value: int) -> None:
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.default_font_size = Pt(value)

    def _on_preset_change(self, _index: int) -> None:
        # Кегль следует за пресетом, только если пользователь его не трогал:
        # 12 pt для обычного оформления, 14 pt для ГОСТ.
        preset = self.preset_combobox.currentData() or "default"
        if self.size_spinbox.value() in {int(size) for size in PRESET_FONT_SIZE.values()}:
            self.size_spinbox.setValue(int(PRESET_FONT_SIZE[preset]))
        self._apply_document_options()

    def _apply_document_options(self, *_args: Any) -> None:
        if isinstance(self.converter, MarkdownToWordConverter):
            self.converter.document_options = DocumentOptions(
                preset=self.preset_combobox.currentData() or "default",
                toc=self.toc_checkbox.isChecked(),
            )

    def _select_files(self) -> None:
        suffix = "*.md *.markdown" if self.current_converter_type == "md_to_word" else "*.docx"
        paths, _ = QFileDialog.getOpenFileNames(self, self._text["add_files"], "", f"Supported files ({suffix})")
        self._add_sources(paths)

    def _select_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, self._text["add_folder"])
        if folder:
            self._add_sources([folder])

    def _add_sources(self, paths: list[str] | list[Path]) -> None:
        discovered = discover_sources((Path(path) for path in paths), self.current_converter_type)
        known = set(self.selected_files)
        self.selected_files.extend(path for path in discovered if path not in known)
        self._refresh_queue()

    def _remove_selected_files(self) -> None:
        selected = {index.row() for index in self.files_listbox.selectedIndexes()}
        self.selected_files = [
            source for index, source in enumerate(self.selected_files) if index not in selected
        ]
        self._refresh_queue()

    def _clear_files(self) -> None:
        self.selected_files.clear()
        self._refresh_queue()

    def _select_output_directory(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, self._text["choose_output"])
        if directory:
            self.output_directory = Path(directory)
            self._update_ui()

    def _reset_output_directory(self) -> None:
        self.output_directory = None
        self._update_ui()

    def _convert_files(self) -> None:
        text_tab_index = self.tabs.indexOf(self.text_tab)
        if self.current_converter_type == "md_to_word" and self.tabs.currentIndex() == text_tab_index:
            self._convert_text()
            return
        if not self.selected_files:
            QMessageBox.warning(self, self.windowTitle(), self._text["no_files"])
            return

        suffix = ".docx" if self.current_converter_type == "md_to_word" else ".md"
        queue = list(self.selected_files)
        outputs = resolve_output_paths(queue, self.output_directory, suffix)

        lockable_widgets = (
            self.convert_button,
            self.files_listbox,
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.clear_button,
            self.drop_hint,
        )
        try:
            for widget in lockable_widgets:
                widget.setEnabled(False)
            self.setAcceptDrops(False)
            self.files_listbox.setAcceptDrops(False)
            self.progress.show()
            self.progress.setRange(0, len(queue))
            success_count = 0
            errors: list[str] = []
            warnings: list[str] = []
            for index, source in enumerate(queue, start=1):
                self.status_label.setText(self._text["converting"].format(filename=source.name))
                QApplication.processEvents()
                try:
                    file_warnings = self.converter.convert_file(source, outputs[source])
                except ConversionError as error:
                    errors.append(
                        f"{source.name}: " + self._text["convert_failed"].format(error=error)
                    )
                else:
                    success_count += 1
                    warnings.extend(
                        _describe_warning(source.name, warning) for warning in file_warnings
                    )
                self.progress.setValue(index)
                QApplication.processEvents()

            details = errors + warnings
            result = self._text["result"].format(success=success_count, errors=len(errors))
            if details:
                QMessageBox.warning(self, self._text["errors"], result + "\n\n" + "\n".join(details))
            else:
                QMessageBox.information(self, self.windowTitle(), result)
            self.status_label.setText(self._text["finished"])
        finally:
            self.progress.hide()
            self.progress.reset()
            self.setAcceptDrops(True)
            self.files_listbox.setAcceptDrops(True)
            for widget in lockable_widgets:
                widget.setEnabled(True)
            self._update_queue_buttons()

    def _convert_text(self) -> None:
        if not isinstance(self.converter, MarkdownToWordConverter):
            return
        content = self.text_input.toPlainText()
        if not content.strip():
            QMessageBox.warning(self, self.windowTitle(), self._text["empty_text"])
            return
        output_path, _ = QFileDialog.getSaveFileName(self, self._text["save_as"], "", "Word files (*.docx)")
        if not output_path:
            return
        output = Path(output_path).with_suffix(".docx")
        try:
            warnings = cast(MarkdownToWordConverter, self.converter).convert_content(content, output)
        except ConversionError as error:
            QMessageBox.critical(
                self, self._text["errors"], self._text["convert_failed"].format(error=error)
            )
            return
        message = self._text["converted_ok"]
        if warnings:
            message += "\n\n" + "\n".join(
                _describe_warning(self._text["text_tab"], warning) for warning in warnings
            )
        QMessageBox.information(self, self.windowTitle(), message)


def main():
    """Главная функция"""
    app = QApplication(sys.argv)
    window = ConverterGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
