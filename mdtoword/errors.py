"""Общий контракт ошибок и предупреждений конвертации.

Вынесено в отдельный модуль, чтобы им могли пользоваться оба направления
конвертации (``converters``, ``docx_to_markdown``) и рендерер, не импортируя
друг друга по кругу.
"""

from __future__ import annotations


class ConversionError(Exception):
    """Конвертация не удалась. Сообщение пригодно для показа без перевода."""


class ConversionWarning(str):
    """Предупреждение конвертации: обычная строка плюс машиночитаемые поля.

    Подкласс ``str`` намеренно: GUI, тесты и любой код, который работает с
    предупреждениями как со строками, продолжают работать без изменений, а
    MCP-сервер достаёт из того же объекта ``code`` и ``line``, чтобы агент
    мог исправить исходник по номеру строки, не ища конструкцию по тексту.

    ``code`` — стабильный идентификатор вида ``formula_unsupported``;
    ``line`` — номер строки исходного Markdown, начиная с 1, или ``None``,
    если строку установить нельзя.
    """

    code: str
    line: int | None

    def __new__(
        cls, message: str, code: str = "general", line: int | None = None
    ) -> "ConversionWarning":
        warning = super().__new__(cls, message)
        warning.code = code
        warning.line = line
        return warning

    def __reduce__(self):  # pragma: no cover - pickling support for workers
        return (ConversionWarning, (str(self), self.code, self.line))
