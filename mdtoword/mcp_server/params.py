"""Типы параметров инструментов: перечисления и поля с описаниями для JSON-схемы."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field

# Допустимые значения дублируют кортежи из options.py: Literal нужен, чтобы
# JSON-схема инструмента показывала enum. Расхождение ловит тест.
Preset = Literal["default", "gost"]
PageSize = Literal["A4", "Letter"]
LineBreaks = Literal["soft", "preserve"]
FootnoteMode = Literal["native", "section"]

InputPaths = Annotated[
    list[str],
    Field(
        min_length=1,
        description="Files and/or directories, mixed; directories are scanned recursively",
    ),
]
OutputDir = Annotated[
    str | None,
    Field(
        description=(
            "Absolute directory for all outputs, created if missing; "
            "null writes each output next to its source"
        )
    ),
]
FontName = Annotated[
    str,
    Field(min_length=1, pattern=r"\S", description="Font family for body text"),
]
FontSize = Annotated[
    float | None,
    Field(
        gt=0,
        le=400,
        description="Body font size in points; null uses the preset default (12 default, 14 gost)",
    ),
]
FootnotesHeading = Annotated[
    str,
    Field(description="Heading of the end-of-document notes section (footnotes='section' only)"),
]
FetchRemoteImages = Annotated[
    bool,
    Field(description="Download http(s) images; only for Markdown from a trusted source"),
]
ImageRoot = Annotated[
    str | None,
    Field(description="Directory to allow local images from instead of the inputs' own directories"),
]
PresetParam = Annotated[
    Preset,
    Field(description="'default' — neutral styling; 'gost' — GOST 7.32-2017 layout"),
]
PageSizeParam = Annotated[
    PageSize | None,
    Field(description="Paper size; null means A4"),
]
LanguageParam = Annotated[
    str,
    Field(
        pattern=r"^(auto|[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*)$",
        description="Proofing language tag such as 'ru-RU' or 'en-US'; 'auto' detects it from the text",
    ),
]
LineBreaksParam = Annotated[
    LineBreaks,
    Field(description="'soft' — a single newline inside a paragraph is a space; 'preserve' — a line break"),
]
TemplateParam = Annotated[
    str | None,
    Field(description="Path of an existing reference .docx or .dotx whose styles, margins and headers are reused"),
]
TocParam = Annotated[
    bool,
    Field(description="Insert a table of contents at the start of the document"),
]
FootnotesParam = Annotated[
    FootnoteMode,
    Field(description="'native' — Word footnotes at the page bottom; 'section' — numbered notes at the end"),
]
