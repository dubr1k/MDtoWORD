"""Параметры оформления документа Word, не сводящиеся к шрифту.

Шрифт и кегль исторически передаются в конвертер отдельными аргументами
(их меняет GUI на лету), а всё, что описывает документ целиком — пресет,
формат страницы, язык, переносы строк, шаблон, оглавление, сноски, —
собрано здесь, чтобы GUI, MCP-сервер и тесты описывали одно и то же одним
объектом, а не растущим списком позиционных аргументов.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

PRESETS = ("default", "gost", "gost_user")
PAGE_SIZES = ("A4", "Letter")
LINE_BREAK_MODES = ("soft", "preserve")
FOOTNOTE_MODES = ("native", "section")

# Кегль по умолчанию для каждого пресета: ГОСТ 7.32-2017 требует не меньше
# 12 pt, а на практике кафедры и журналы ждут 14 pt Times New Roman.
PRESET_FONT_SIZE = {"default": 12.0, "gost": 14.0, "gost_user": 12.0}


@dataclass(frozen=True)
class DocumentOptions:
    """Оформление документа целиком.

    ``preset``
        ``"default"`` — нейтральное оформление; ``"gost"`` — ГОСТ 7.32-2017:
        поля 30/15/20/20 мм, полуторный интервал, абзацный отступ 1,25 см,
        подписи «Рисунок N — …» и «Таблица N — …», номер страницы внизу по
        центру.
        ``"gost_user"`` — пользовательская адаптация: A4, 12 pt, поля
        30/15/15/15 мм, интервал 1,5, по ширине, без колонтитулов.
        Это не заявление о полном соответствии ГОСТ 7.32. Явный шаблон
        сохраняет приоритет над оформлением пресета.
    ``page_size``
        ``"A4"`` или ``"Letter"``; ``None`` — A4.
    ``language``
        Язык текста для проверки орфографии и переносов в Word: тег вида
        ``"ru-RU"`` или ``"auto"`` — определить по доле кириллицы.
    ``line_breaks``
        ``"soft"`` — одиночный перенос строки внутри абзаца становится
        пробелом, как в CommonMark/GFM; ``"preserve"`` — разрывом строки.
    ``template``
        Путь к .docx, из которого берутся стили, поля и колонтитулы (как
        ``--reference-doc`` у pandoc). Содержимое шаблона не копируется.
    ``toc``
        Вставить оглавление (поле TOC) в начало документа. Маркер ``[TOC]``
        отдельным абзацем вставляет оглавление в своё место независимо от
        этого флага.
    ``footnotes``
        ``"native"`` — настоящие сноски Word внизу страницы; ``"section"`` —
        нумерованный раздел в конце документа под ``footnotes_heading``.
    """

    preset: str = "default"
    page_size: str | None = None
    language: str = "auto"
    line_breaks: str = "soft"
    template: Path | None = None
    toc: bool = False
    footnotes: str = "native"

    def __post_init__(self) -> None:
        _require(self.preset, PRESETS, "preset")
        if self.page_size is not None:
            _require(self.page_size, PAGE_SIZES, "page_size")
        _require(self.line_breaks, LINE_BREAK_MODES, "line_breaks")
        _require(self.footnotes, FOOTNOTE_MODES, "footnotes")
        if not self.language or not self.language.strip():
            raise ValueError("language must be 'auto' or a language tag such as 'ru-RU'")
        if self.template is not None:
            object.__setattr__(self, "template", Path(self.template))


def _require(value: str, allowed: tuple[str, ...], name: str) -> None:
    if value not in allowed:
        choices = ", ".join(repr(item) for item in allowed)
        raise ValueError(f"{name} must be one of {choices}, got {value!r}")
