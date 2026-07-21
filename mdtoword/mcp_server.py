"""MCP-сервер: конвертация Markdown ↔ Word для агентов.

Транспорт — stdio, поэтому stdout занят самим протоколом: печатать в него
нельзя ни при импорте, ни во время работы. Все диагностические сообщения
должны идти в stderr.

Контракт инструментов — только пути к файлам. Содержимое документов через
границу MCP не передаётся: docx — бинарный формат, а Markdown ссылается на
изображения относительными путями, которые вне файловой системы теряют смысл.
"""

from __future__ import annotations

from pathlib import Path

from docx.shared import Pt
from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field

from .converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)
from .workflow import discover_sources, resolve_output_paths

mcp = FastMCP("mdtoword")


class ConvertedFile(BaseModel):
    """One successfully converted file."""

    source: str = Field(description="Absolute path of the input file")
    output: str = Field(description="Absolute path of the file that was written")
    warnings: list[str] = Field(
        default_factory=list,
        description="Non-fatal issues; the output was still written",
    )


class FailedFile(BaseModel):
    """One file that could not be converted."""

    source: str = Field(description="Absolute path of the input file")
    error: str = Field(description="Why this file could not be converted")


class ConversionReport(BaseModel):
    """The result of a batch conversion."""

    sources_found: int = Field(
        description=(
            "How many supported files the inputs resolved to. "
            "0 means the paths matched nothing — check the paths rather than "
            "assuming there was nothing to do."
        )
    )
    converted: list[ConvertedFile] = Field(default_factory=list)
    failed: list[FailedFile] = Field(default_factory=list)


class PreviewedFile(BaseModel):
    """One Markdown file rendered without writing anything."""

    source: str = Field(description="Absolute path of the input file")
    warnings: list[str] = Field(
        default_factory=list,
        description="What would not survive the conversion to Word",
    )


class PreviewReport(BaseModel):
    """Result of previewing a batch of Markdown files."""

    sources_found: int = Field(
        description=(
            "How many Markdown files the inputs resolved to. "
            "0 means the paths matched nothing."
        )
    )
    previews: list[PreviewedFile] = Field(default_factory=list)
    failed: list[FailedFile] = Field(default_factory=list)


def _resolve_inputs(inputs: list[str], mode: str) -> list[Path]:
    """Развернуть переданные пути в отсортированный список исходных файлов."""
    if not inputs:
        raise ValueError(
            "inputs must not be empty; pass at least one file or directory path"
        )
    return discover_sources([Path(item).expanduser() for item in inputs], mode)


def _prepare_output_dir(output_dir: str | None) -> Path | None:
    """Подготовить каталог назначения; None означает «рядом с исходником»."""
    if output_dir is None:
        return None
    # resolve() — иначе относительный output_dir оставит ConvertedFile.output
    # относительным, вопреки его же Field(description="Absolute path ...")
    directory = Path(output_dir).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory


@mcp.tool()
def markdown_to_word(
    inputs: list[str],
    output_dir: str | None = None,
    font_name: str = "Times New Roman",
    font_size: float = 12,
    footnotes_heading: str = "Footnotes",
) -> ConversionReport:
    """Convert Markdown files to Word .docx documents.

    `inputs` accepts files and directories mixed together; directories are
    scanned recursively for .md and .markdown files.

    Supports GitHub Flavored Markdown: headings, emphasis, lists, task lists,
    tables, blockquotes, code blocks, footnotes, links and images. LaTeX math
    (`$inline$`, `$$display$$`, and amsmath environments) becomes native Word
    OMML equations rather than an image or plain text.

    Each output is written next to its source unless `output_dir` is given.
    Existing files at the target paths are overwritten without warning.

    Check `sources_found` in the result: 0 means the paths matched no
    Markdown files at all.
    """
    sources = _resolve_inputs(inputs, "md_to_word")
    outputs = resolve_output_paths(sources, _prepare_output_dir(output_dir), ".docx")
    converter = MarkdownToWordConverter(font_name, Pt(font_size), footnotes_heading)
    return _run_batch(sources, outputs, converter)


@mcp.tool()
def word_to_markdown(
    inputs: list[str],
    output_dir: str | None = None,
) -> ConversionReport:
    """Convert Word .docx documents to Markdown files.

    `inputs` accepts files and directories mixed together; directories are
    scanned recursively for .docx files.

    This direction is LOSSY. It extracts headings (from `Heading N` styles),
    bold and italic runs, and tables. Everything else — lists, images,
    equations, footnotes, colours, and other styling — is flattened to plain
    text. Do not round-trip documents through this tool expecting to get the
    original back.

    Each output is written next to its source unless `output_dir` is given.
    Existing files at the target paths are overwritten without warning.

    Check `sources_found` in the result: 0 means the paths matched no
    .docx files at all.
    """
    sources = _resolve_inputs(inputs, "word_to_md")
    outputs = resolve_output_paths(sources, _prepare_output_dir(output_dir), ".md")
    return _run_batch(sources, outputs, WordToMarkdownConverter())


@mcp.tool()
def preview_markdown(
    inputs: list[str],
    font_name: str = "Times New Roman",
    font_size: float = 12,
    footnotes_heading: str = "Footnotes",
) -> PreviewReport:
    """Check what Markdown would lose in Word, without writing any file.

    Runs the full conversion in memory and discards the result, reporting only
    the warnings: missing images, LaTeX that cannot become an OMML equation,
    math inside table cells. Use this before `markdown_to_word` when you want
    to fix the source first, or to inspect a document you must not overwrite.

    Nothing is written to disk by this tool.
    """
    sources = _resolve_inputs(inputs, "md_to_word")
    converter = MarkdownToWordConverter(font_name, Pt(font_size), footnotes_heading)
    report = PreviewReport(sources_found=len(sources))
    for source in sources:
        try:
            warnings = converter.preview_file(source)
        except ConversionError as error:
            report.failed.append(FailedFile(source=str(source), error=str(error)))
        else:
            report.previews.append(
                PreviewedFile(source=str(source), warnings=warnings)
            )
    return report


def _run_batch(
    sources: list[Path],
    outputs: dict[Path, Path],
    converter: MarkdownToWordConverter | WordToMarkdownConverter,
) -> ConversionReport:
    """Сконвертировать все файлы, не прерываясь на отдельных отказах."""
    report = ConversionReport(sources_found=len(sources))
    for source in sources:
        try:
            warnings = converter.convert_file(source, outputs[source])
        except ConversionError as error:
            report.failed.append(FailedFile(source=str(source), error=str(error)))
        else:
            report.converted.append(
                ConvertedFile(
                    source=str(source), output=str(outputs[source]), warnings=warnings
                )
            )
    return report


def main() -> None:
    """Запустить сервер на транспорте stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
