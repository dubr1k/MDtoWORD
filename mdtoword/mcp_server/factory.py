"""Единственное место, где сервер создаёт конвертеры.

Инструменты получают конвертеры только через функции этого модуля, а
классы ищутся в его пространстве имён в момент вызова: тесты подменяют
``mdtoword.mcp_server.factory.MarkdownToWordConverter`` и
``...WordToMarkdownConverter``, и подмена действует на все инструменты.
"""

from __future__ import annotations

from contextlib import ExitStack
from pathlib import Path

from docx.shared import Pt

from ..converters import MarkdownToWordConverter, WordToMarkdownConverter
from ..options import PRESET_FONT_SIZE, DocumentOptions
from .paths import resolve_image_roots, resolve_template


def markdown_converter(
    inputs: list[str],
    scratch: ExitStack,
    *,
    font_name: str,
    font_size: float | None,
    footnotes_heading: str,
    fetch_remote_images: bool,
    image_root: str | None,
    preset: str,
    page_size: str | None,
    language: str,
    line_breaks: str,
    template: str | None,
    toc: bool,
    footnotes: str,
) -> MarkdownToWordConverter:
    """Собрать конвертер с параметрами инструмента; ошибки — ValueError."""
    options = DocumentOptions(
        preset=preset,
        page_size=page_size,
        language=language,
        line_breaks=line_breaks,
        template=resolve_template(template, scratch),
        toc=toc,
        footnotes=footnotes,
    )
    image_roots = (
        [Path(image_root).expanduser().resolve()]
        if image_root is not None
        else resolve_image_roots(inputs)
    )
    size = font_size if font_size is not None else PRESET_FONT_SIZE[preset]
    return MarkdownToWordConverter(
        font_name=font_name,
        font_size=Pt(size),
        footnotes_heading=footnotes_heading,
        allow_remote_images=fetch_remote_images,
        image_roots=image_roots,
        document_options=options,
    )


def word_converter(extract_media: bool) -> WordToMarkdownConverter:
    """Конвертер Word → Markdown для ``word_to_markdown``."""
    return WordToMarkdownConverter(extract_media=extract_media)


def formula_converter() -> MarkdownToWordConverter:
    """Конвертер для проверки одной формулы в ``check_latex``.

    Документ из одной формулы не ссылается на изображения, поэтому ни
    сеть, ни локальные файлы ему не разрешены.
    """
    return MarkdownToWordConverter(allow_remote_images=False, image_roots=[])
