from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PyQt6.QtCore import QSettings, QStandardPaths
from PyQt6.QtGui import QColor, QPalette
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
        border="#2A3038",
        border_strong="#3B4552",
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
        border="#D0D7E2",
        border_strong="#AEB9C9",
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


# Chevron paths tuned to render as clean strokes at a 12x12 viewBox. Filled
# triangles and CSS border-triangle tricks were tried first and both looked
# bad in Qt (the border trick renders as squares, not triangles) — a small
# stroked SVG chevron is what actually looks right.
_CHEVRON_PATH_D = {
    "down": "M2.5 4.5 L6 8 L9.5 4.5",
    "up": "M2.5 7.5 L6 4 L9.5 7.5",
}


def _chevron_svg(color: str, direction: str) -> str:
    """Render a minimal chevron as inline SVG markup, tinted with ``color``."""
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12">'
        f'<path d="{_CHEVRON_PATH_D[direction]}" fill="none" stroke="{color}" stroke-width="1.6" '
        'stroke-linecap="round" stroke-linejoin="round"/></svg>'
    )


def _arrow_css(selector: str, image_path: str | None, size: int) -> str:
    """Build a ``{selector} { image: url(...); width; height; }`` rule.

    Returns an empty string when ``image_path`` is falsy. Qt silently drops
    the arrow entirely if width/height are set on an arrow subcontrol without
    a matching image, so the two must always travel together — omitting both
    leaves the native Fusion arrow in place instead of an invisible one.
    """
    if not image_path:
        return ""
    return (
        f"{selector} {{\n"
        f'                image: url("{image_path}");\n'
        f"                width: {size}px;\n"
        f"                height: {size}px;\n"
        "            }"
    )


class ThemeManager:
    """Persist and apply the application's complete Qt widget theme."""

    def __init__(
        self,
        settings: QSettings | None = None,
        icon_cache_dir: str | Path | None = None,
    ) -> None:
        self._settings = settings or QSettings("dubr1k", "MDtoWord")
        stored_theme = str(self._settings.value("theme", "dark"))
        self.theme = stored_theme if stored_theme in _PALETTES else "dark"
        # Overrides QStandardPaths' cache location; primarily for tests that
        # need a deterministic (or deliberately unwritable) directory.
        self._icon_cache_dir = icon_cache_dir

    def toggle(self) -> str:
        self.theme = "light" if self.theme == "dark" else "dark"
        self._settings.setValue("theme", self.theme)
        self._settings.sync()
        return self.theme

    def apply(self, app: QApplication) -> None:
        app.setStyle("Fusion")
        palette = _PALETTES[self.theme]
        app.setPalette(self._widget_palette(palette))
        chevron_down, chevron_up = self._ensure_chevron_icons(self.theme, palette)
        app.setStyleSheet(
            self.stylesheet(self.theme, chevron_down=chevron_down, chevron_up=chevron_up)
        )

    def _ensure_chevron_icons(
        self, theme: str, palette: ThemePalette
    ) -> tuple[str | None, str | None]:
        """Write themed chevron SVGs to the icon cache; return their paths.

        Returns ``(None, None)`` if the cache directory cannot be created or
        written to for any reason (permissions, a read-only filesystem, a
        path that collides with an existing file, ...). Callers must treat
        that as "fall back to native arrows" rather than pass a dangling
        path into the stylesheet — a missing ``image:`` target renders as
        nothing at all, which is worse than the plain Fusion triangle.
        """
        try:
            if self._icon_cache_dir is not None:
                directory = Path(self._icon_cache_dir)
            else:
                location = QStandardPaths.writableLocation(
                    QStandardPaths.StandardLocation.CacheLocation
                )
                directory = Path(location) if location else Path.cwd()
            directory.mkdir(parents=True, exist_ok=True)

            down_path = directory / f"chevron-down-{theme}.svg"
            up_path = directory / f"chevron-up-{theme}.svg"
            down_path.write_text(_chevron_svg(palette.text_muted, "down"), encoding="utf-8")
            up_path.write_text(_chevron_svg(palette.text_muted, "up"), encoding="utf-8")
            return down_path.as_posix(), up_path.as_posix()
        except OSError:
            return None, None

    @staticmethod
    def _widget_palette(palette: ThemePalette) -> QPalette:
        qpalette = QPalette()
        color_roles = {
            QPalette.ColorRole.Window: palette.background,
            QPalette.ColorRole.WindowText: palette.text,
            QPalette.ColorRole.Base: palette.input,
            QPalette.ColorRole.AlternateBase: palette.surface,
            QPalette.ColorRole.Text: palette.text,
            QPalette.ColorRole.Button: palette.surface,
            QPalette.ColorRole.ButtonText: palette.text,
            QPalette.ColorRole.BrightText: palette.text,
            QPalette.ColorRole.Highlight: palette.accent,
            QPalette.ColorRole.HighlightedText: "#FFFFFF",
            QPalette.ColorRole.Link: palette.accent,
            QPalette.ColorRole.ToolTipBase: palette.surface_raised,
            QPalette.ColorRole.ToolTipText: palette.text,
            QPalette.ColorRole.PlaceholderText: palette.text_muted,
        }
        for role, value in color_roles.items():
            qpalette.setColor(role, QColor(value))
        disabled = QColor(palette.disabled)
        for role in (
            QPalette.ColorRole.Text,
            QPalette.ColorRole.ButtonText,
            QPalette.ColorRole.WindowText,
        ):
            qpalette.setColor(QPalette.ColorGroup.Disabled, role, disabled)
        return qpalette

    @staticmethod
    def stylesheet(
        theme: str,
        *,
        chevron_down: str | None = None,
        chevron_up: str | None = None,
    ) -> str:
        palette = _PALETTES.get(theme, _PALETTES["dark"])
        combo_down_arrow = _arrow_css("QComboBox::down-arrow", chevron_down, 12)
        spin_up_arrow = _arrow_css("QSpinBox::up-arrow", chevron_up, 11)
        spin_down_arrow = _arrow_css("QSpinBox::down-arrow", chevron_down, 11)
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
            QLabel#output-path {{ background: {palette.surface_raised}; border-radius: 8px; color: {palette.text_muted}; padding: 9px 11px; }}

            QGroupBox {{
                background: {palette.surface};
                border: 1px solid {palette.border};
                border-radius: 12px;
                margin-top: 0;
                padding: 30px 14px 12px 14px;
                font-weight: 700;
            }}
            QGroupBox::title {{
                subcontrol-origin: border;
                subcontrol-position: top left;
                left: 14px;
                top: 9px;
                background: transparent;
                padding: 0;
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
                background: transparent;
                border: none;
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 28px;
            }}
            QComboBox::drop-down:hover {{ background: {palette.surface_raised}; }}
            {combo_down_arrow}

            QSpinBox::up-button {{
                background: transparent;
                border: none;
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 22px;
            }}
            QSpinBox::up-button:hover {{ background: {palette.surface_raised}; }}
            QSpinBox::down-button {{
                background: transparent;
                border: none;
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 22px;
            }}
            QSpinBox::down-button:hover {{ background: {palette.surface_raised}; }}
            {spin_up_arrow}
            {spin_down_arrow}

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
                color: {palette.accent};
                font-size: 15px;
                font-weight: 600;
                padding: 12px 16px;
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
