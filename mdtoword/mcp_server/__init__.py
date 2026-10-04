"""MCP-сервер: конвертация Markdown ↔ Word для агентов.

Транспорт — stdio, поэтому stdout занят самим протоколом: печатать в него
нельзя ни при импорте, ни во время работы. Все диагностические сообщения
должны идти в stderr.

Контракт инструментов — только пути к файлам. Содержимое документов через
границу MCP не передаётся: docx — бинарный формат, а Markdown ссылается на
изображения относительными путями, которые вне файловой системы теряют смысл.
Исключение — ``check_latex``: формулы приходят строками, файлов у них нет.

Инструменты асинхронные: FastMCP вызывает синхронный инструмент прямо в
цикле событий, и длинный батч блокировал бы пинги и отмену. Поэтому каждый
файл конвертируется в рабочем потоке (``anyio.to_thread.run_sync``), а после
каждого файла клиенту уходит уведомление о прогрессе. Отмена срабатывает
между файлами: поток с текущим файлом дорабатывает, следующий не начинается.

Помимо инструментов сервер отдаёт ресурс ``mdtoword://guide/markdown`` —
руководство для агента по Markdown, который конвертируется без потерь, — и
промпт ``prepare_markdown_for_word`` с порядком работы над документом.

Модули пакета:

- ``server`` — экземпляр FastMCP, инструкции, ``main``;
- ``params`` — типы параметров инструментов (enum и описания полей);
- ``models`` — модели отчётов (outputSchema);
- ``paths`` — входы, выходы, песочница изображений, шаблон (.dotx);
- ``factory`` — создание конвертеров (точка подмены для тестов);
- ``batch`` — обработка файлов в рабочем потоке, прогресс, диагностика;
- ``tools_convert`` — markdown_to_word, word_to_markdown, preview_markdown;
- ``tools_check`` — check_latex;
- ``tools_inspect`` — inspect_docx;
- ``guide`` — ресурс с руководством и промпт.
"""

from __future__ import annotations

# Context доступен как mcp_server.Context: тесты подменяют его report_progress.
from mcp.server.fastmcp import Context as Context

from .models import (
    ConversionReport,
    ConvertedFile,
    Diagnostic,
    DocxCounts,
    DocxInspection,
    DocxProperties,
    FailedFile,
    LatexCheckReport,
    LatexCheckResult,
    OutlineEntry,
    PageSetup,
    PreviewedFile,
    PreviewReport,
)
from .params import FootnoteMode, LineBreaks, PageSize, Preset
from .server import GUIDE_URI, main, mcp

# isort: off
# Порядок импорта — порядок регистрации на mcp, а значит и порядок в
# list_tools: markdown_to_word, word_to_markdown, preview_markdown,
# check_latex, inspect_docx; затем ресурс и промпт.
from .tools_convert import markdown_to_word, word_to_markdown, preview_markdown
from .tools_check import check_latex
from .tools_inspect import inspect_docx
from .guide import markdown_guide, prepare_markdown_for_word
# isort: on

__all__ = [
    "GUIDE_URI",
    "ConversionReport",
    "ConvertedFile",
    "Diagnostic",
    "DocxCounts",
    "DocxInspection",
    "DocxProperties",
    "FailedFile",
    "FootnoteMode",
    "LatexCheckReport",
    "LatexCheckResult",
    "LineBreaks",
    "OutlineEntry",
    "PageSetup",
    "PageSize",
    "Preset",
    "PreviewReport",
    "PreviewedFile",
    "check_latex",
    "inspect_docx",
    "main",
    "markdown_guide",
    "markdown_to_word",
    "mcp",
    "prepare_markdown_for_word",
    "preview_markdown",
    "word_to_markdown",
]
