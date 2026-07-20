from __future__ import annotations

from dataclasses import dataclass

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication


@dataclass(frozen=True)
class ThemePalette:
    background: str
    surface: str
    surface_raised: str
    border: str
    border_strong: str
    text: str
    text_muted: str
    accent: str
    accent_hover: str
    accent_pressed: str
    selection: str
    destructive: str
    destructive_hover: str
    input: str
    disabled: str


_PALETTES = {
    "dark": ThemePalette(
        background="#0D1117",
        surface="#161B22",
        surface_raised="#1C2129",
        border="#21262D",
        border_strong="#30363D",
        text="#F0F6FC",
        text_muted="#A9B1D6",
        accent="#7C6CFF",
        accent_hover="#9386FF",
        accent_pressed="#6657E8",
        selection="#2B2754",
        destructive="#D95763",
        destructive_hover="#EB6A75",
        input="#0F1520",
        disabled="#6E7681",
    ),
    "light": ThemePalette(
        background="#F6F8FC",
        surface="#FFFFFF",
        surface_raised="#F0F3F8",
        border="#D8DEE9",
        border_strong="#BEC7D5",
        text="#19202B",
        text_muted="#5D687A",
        accent="#5E50D6",
        accent_hover="#7062E9",
        accent_pressed="#4D40BC",
        selection="#E8E5FF",
        destructive="#C83F4D",
        destructive_hover="#DE5261",
        input="#FFFFFF",
        disabled="#8B95A5",
    ),
}


class ThemeManager:
    """Persist and apply the application's complete Qt widget theme."""

    def __init__(self, settings: QSettings | None = None) -> None:
        self._settings = settings or QSettings("dubr1k", "MDtoWord")
        stored_theme = str(self._settings.value("theme", "dark"))
        self.theme = stored_theme if stored_theme in _PALETTES else "dark"

    def toggle(self) -> str:
        self.theme = "light" if self.theme == "dark" else "dark"
        self._settings.setValue("theme", self.theme)
        self._settings.sync()
        return self.theme

    def apply(self, app: QApplication) -> None:
        app.setStyle("Fusion")
        app.setStyleSheet(self.stylesheet(self.theme))

    @staticmethod
    def stylesheet(theme: str) -> str:
        palette = _PALETTES.get(theme, _PALETTES["dark"])
        return f"""
            QWidget {{
                background: {palette.background};
                color: {palette.text};
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                font-size: 13px;
            }}

            QMainWindow, QDialog {{ background: {palette.background}; }}
            QLabel {{ background: transparent; }}
            QLabel#title-label {{
                color: {palette.text};
                font-size: 28px;
                font-weight: 700;
                letter-spacing: -0.3px;
            }}
            QLabel#status-label {{ color: {palette.text_muted}; font-weight: 600; }}
            QLabel#drop-zone {{ color: {palette.accent}; font-size: 15px; font-weight: 600; }}
            QLabel#output-path {{ background: {palette.surface_raised}; border-radius: 8px; color: {palette.text_muted}; padding: 9px 11px; }}

            QGroupBox {{
                background: {palette.surface};
                border: 1px solid {palette.border};
                border-radius: 12px;
                margin-top: 16px;
                padding: 18px 16px 15px 16px;
                font-weight: 700;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 14px;
                background: {palette.surface};
                border: none;
                border-radius: 4px;
                padding: 1px 8px;
                color: {palette.text};
            }}

            QLineEdit, QPlainTextEdit, QListWidget, QComboBox, QSpinBox {{
                background: {palette.input};
                border: 1px solid {palette.border};
                border-radius: 8px;
                color: {palette.text};
                padding: 8px 10px;
                selection-background-color: {palette.selection};
                selection-color: {palette.text};
            }}
            QPlainTextEdit {{ padding: 12px; }}
            QListWidget {{ outline: 0; padding: 5px; }}
            QListWidget::item {{
                border-radius: 6px;
                padding: 7px 8px;
                margin: 1px 0;
            }}
            QListWidget::item:selected {{ background: {palette.selection}; color: {palette.text}; }}
            QLineEdit:focus, QPlainTextEdit:focus, QListWidget:focus,
            QComboBox:focus, QSpinBox:focus {{ border: 1px solid {palette.accent}; }}
            QLineEdit:disabled, QPlainTextEdit:disabled, QListWidget:disabled,
            QComboBox:disabled, QSpinBox:disabled {{ color: {palette.disabled}; }}

            QComboBox::drop-down {{
                border: 0;
                width: 28px;
            }}
            QComboBox::down-arrow {{
                width: 8px;
                height: 8px;
            }}
            QComboBox QAbstractItemView {{
                background: {palette.surface_raised};
                border: 1px solid {palette.border_strong};
                color: {palette.text};
                outline: 0;
                padding: 4px;
                selection-background-color: {palette.selection};
            }}

            QTabWidget::pane {{
                background: {palette.surface};
                border: 1px solid {palette.border};
                border-radius: 14px;
                top: -1px;
            }}
            QWidget#tab-page {{ background: transparent; }}
            QTabWidget::tab-bar {{ alignment: center; }}
            QTabBar {{ background: transparent; }}
            QTabBar::tab {{
                background: {palette.surface_raised};
                border: 1px solid transparent;
                border-radius: 7px;
                color: {palette.text_muted};
                margin: 0 3px 5px 3px;
                padding: 7px 17px;
                font-weight: 600;
            }}
            QTabBar::tab:hover {{ color: {palette.text}; background: {palette.border}; }}
            QTabBar::tab:selected {{
                background: {palette.accent};
                color: white;
            }}

            QPushButton {{
                background: transparent;
                border: 1px solid {palette.border_strong};
                border-radius: 8px;
                color: {palette.text};
                padding: 8px 12px;
                font-weight: 600;
            }}
            QPushButton:hover {{ background: {palette.surface_raised}; border-color: {palette.text_muted}; }}
            QPushButton:pressed {{ background: {palette.border}; }}
            QPushButton:disabled {{ color: {palette.disabled}; border-color: {palette.border}; }}
            QPushButton#primary-button {{
                background: {palette.accent};
                border-color: {palette.accent};
                color: white;
                font-size: 15px;
                font-weight: 700;
                padding: 11px 18px;
            }}
            QPushButton#primary-button:hover {{ background: {palette.accent_hover}; border-color: {palette.accent_hover}; }}
            QPushButton#primary-button:pressed {{ background: {palette.accent_pressed}; border-color: {palette.accent_pressed}; }}
            QPushButton#danger-button:hover {{ background: {palette.destructive}; border-color: {palette.destructive}; color: white; }}
            QPushButton#theme-button {{
                border-radius: 15px;
                min-width: 30px;
                max-width: 30px;
                min-height: 30px;
                max-height: 30px;
                padding: 0;
            }}

            QLabel#drop-zone {{
                background: {palette.surface_raised};
                border: 2px dashed {palette.accent};
                border-radius: 12px;
            }}
            QProgressBar {{
                background: {palette.input};
                border: 0;
                border-radius: 4px;
                color: transparent;
                min-height: 7px;
            }}
            QProgressBar::chunk {{ background: {palette.accent}; border-radius: 4px; }}

            QScrollBar:vertical {{
                background: transparent;
                width: 11px;
                margin: 4px 2px;
            }}
            QScrollBar::handle:vertical {{
                background: {palette.border_strong};
                border-radius: 4px;
                min-height: 28px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {palette.text_muted}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
            QScrollBar:horizontal {{
                background: transparent;
                height: 11px;
                margin: 2px 4px;
            }}
            QScrollBar::handle:horizontal {{ background: {palette.border_strong}; border-radius: 4px; min-width: 28px; }}
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

            QToolTip {{
                background: {palette.surface_raised};
                border: 1px solid {palette.border_strong};
                border-radius: 6px;
                color: {palette.text};
                padding: 5px 7px;
            }}
        """
