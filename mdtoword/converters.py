"""Ядро конвертации Markdown ↔ Word.

Модуль намеренно свободен от PyQt6: им пользуются и GUI (``mdtoword.app``),
и MCP-сервер (``mdtoword.mcp_server``), а последний обязан работать без
графической подсистемы.

Контракт: успешная конвертация возвращает список необязательных предупреждений,
неуспешная — бросает :class:`ConversionError`. Сообщения об ошибках намеренно
не локализованы: язык подставляет потребитель, у GUI для этого есть словарь
переводов, а агенту локаль не нужна вовсе.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

from docx.shared import Pt

from .errors import ConversionError, ConversionWarning
from .gfm_renderer import GfmDocxRenderer
from .options import DocumentOptions

# Обратное направление живёт в отдельном модуле; импорт здесь сохраняет
# публичный адрес ``mdtoword.converters.WordToMarkdownConverter``.
from .docx_to_markdown import WordToMarkdownConverter

__all__ = [
    "ConversionError",
    "ConversionWarning",
    "DocumentOptions",
    "MarkdownToWordConverter",
    "WordToMarkdownConverter",
]


class MarkdownToWordConverter:
    """Конвертирует GFM-разметку в документ Word."""

    def __init__(
        self,
        font_name: str = "Times New Roman",
        font_size: Pt = Pt(12),
        footnotes_heading: str = "Footnotes",
        allow_remote_images: bool = True,
        image_roots: Sequence[Path] | None = None,
        *,
        document_options: DocumentOptions | None = None,
    ) -> None:
        self.default_font_name = font_name
        self.default_font_size = font_size
        self.footnotes_heading = footnotes_heading
        self.allow_remote_images = allow_remote_images
        self.image_roots = image_roots
        self.document_options = document_options or DocumentOptions()

    def _render(self, content: str, source_path: Path | None) -> tuple[Any, list[str]]:
        """Отрендерить Markdown, переведя любой сбой рендеринга в ConversionError."""
        try:
            return GfmDocxRenderer(
                self.default_font_name,
                self.default_font_size,
                self.footnotes_heading,
                self.allow_remote_images,
                self.image_roots,
                document_options=self.document_options,
            ).render(content, source_path=source_path)
        except Exception as error:
            raise ConversionError(str(error)) from error

    @staticmethod
    def _read_source(input_path: str | Path) -> tuple[Path, str]:
        """Прочитать исходник, переведя сбой чтения в ConversionError."""
        source_path = Path(input_path)
        try:
            return source_path, source_path.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError) as error:
            # UnicodeDecodeError — подкласс ValueError, а не OSError:
            # файл в CP1251 иначе улетел бы мимо контракта ConversionError.
            raise ConversionError(str(error)) from error

    def convert_content(
        self, content: str, output_path: str | Path, source_path: Path | None = None
    ) -> list[str]:
        """Отрендерить Markdown и сохранить результат в *output_path*."""
        document, warnings = self._render(content, source_path)
        try:
            document.save(str(output_path))
        except Exception as error:
            raise ConversionError(str(error)) from error
        return warnings

    def convert_file(
        self, input_path: str | Path, output_path: str | Path
    ) -> list[str]:
        """Прочитать Markdown-файл и сконвертировать его."""
        source_path, content = self._read_source(input_path)
        return self.convert_content(content, output_path, source_path)

    def preview_content(
        self, content: str, source_path: Path | None = None
    ) -> list[str]:
        """Отрендерить Markdown в память и вернуть варнинги, ничего не сохраняя."""
        _, warnings = self._render(content, source_path)
        return warnings

    def preview_file(self, input_path: str | Path) -> list[str]:
        """Прочитать Markdown-файл и отрендерить его вхолостую."""
        source_path, content = self._read_source(input_path)
        return self.preview_content(content, source_path)

