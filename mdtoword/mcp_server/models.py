"""Модели отчётов инструментов (Pydantic): из них строится outputSchema."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class Diagnostic(BaseModel):
    """A non-fatal issue found while rendering one Markdown file."""

    message: str = Field(description="Human-readable description of the issue")
    code: str = Field(
        description=(
            "Stable machine-readable category, e.g. 'formula_unsupported' or "
            "'image_not_found'; 'general' when the issue has no specific code"
        )
    )
    line: int | None = Field(
        default=None,
        description=(
            "1-based line in the Markdown source where the offending construct "
            "starts; null when it cannot be located. Edit the source at this "
            "line and run preview_markdown again to confirm the fix"
        ),
    )


class ConvertedFile(BaseModel):
    """One successfully converted file."""

    source: str = Field(description="Absolute path of the input file")
    output: str = Field(description="Absolute path of the file that was written")
    warnings: list[Diagnostic] = Field(
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
    converted: list[ConvertedFile] = Field(
        default_factory=list,
        description="Files that were converted successfully",
    )
    failed: list[FailedFile] = Field(
        default_factory=list,
        description="Files that could not be converted, each with the reason",
    )


class PreviewedFile(BaseModel):
    """One Markdown file rendered without writing anything."""

    source: str = Field(description="Absolute path of the input file")
    warnings: list[Diagnostic] = Field(
        default_factory=list,
        description="What would not survive the conversion to Word",
    )


class PreviewReport(BaseModel):
    """Result of previewing a batch of Markdown files."""

    sources_found: int = Field(
        description=(
            "How many Markdown files the inputs resolved to. "
            "0 means the paths matched nothing — check the paths rather than "
            "assuming there was nothing to do."
        )
    )
    previews: list[PreviewedFile] = Field(
        default_factory=list,
        description="Files that were rendered in memory, each with its warnings",
    )
    failed: list[FailedFile] = Field(
        default_factory=list,
        description="Files that could not be previewed, each with the reason",
    )


class LatexCheckResult(BaseModel):
    """Verdict for one formula."""

    formula: str = Field(description="The formula exactly as it was passed in")
    ok: bool = Field(
        description="True if the formula becomes a native, editable Word equation"
    )
    error: str | None = Field(
        default=None,
        description=(
            "Why the formula would be kept as verbatim text instead; "
            "null when ok is true"
        ),
    )


class LatexCheckReport(BaseModel):
    """Result of checking a list of formulas."""

    results: list[LatexCheckResult] = Field(
        description="One entry per input formula, in the same order"
    )
    all_ok: bool = Field(description="True if every formula converts")


class PageSetup(BaseModel):
    """Page geometry of the first section."""

    width_mm: float | None = Field(description="Page width in millimetres")
    height_mm: float | None = Field(description="Page height in millimetres")
    paper: str | None = Field(
        description=(
            "Recognised paper size ('A4', 'Letter', 'Legal', 'A3', 'A5') in "
            "either orientation, 'custom' otherwise, null if the size is unset"
        )
    )
    orientation: Literal["portrait", "landscape"] | None = Field(
        description="Derived from the page dimensions; null if they are unset"
    )
    margin_top_mm: float | None = Field(description="Top margin in millimetres")
    margin_bottom_mm: float | None = Field(description="Bottom margin in millimetres")
    margin_left_mm: float | None = Field(description="Left margin in millimetres")
    margin_right_mm: float | None = Field(description="Right margin in millimetres")


class DocxCounts(BaseModel):
    """Element counts over the main document body, table cells included."""

    paragraphs: int = Field(
        description="Paragraphs with visible content (text, equation or image), headings included"
    )
    headings: int = Field(description="Title and Heading 1-9 paragraphs")
    tables: int = Field(description="Top-level tables")
    images: int = Field(description="Pictures, inline and floating (anchored)")
    equations: int = Field(
        description="Native Word equations (m:oMath), nested ones not counted twice"
    )
    footnotes: int = Field(description="Native footnotes, separators excluded")
    hyperlinks: int = Field(
        description="Hyperlinks, both external URLs and internal bookmark links"
    )
    list_paragraphs: int = Field(
        description="Numbered or bulleted list items (numbering set directly or by style), headings excluded"
    )


class OutlineEntry(BaseModel):
    """One heading in document order."""

    level: int = Field(description="0 for the Title style, 1-9 for Heading 1-9")
    text: str = Field(description="Heading text")


class DocxProperties(BaseModel):
    """Core document properties (File > Info in Word); null when empty."""

    title: str | None = None
    author: str | None = None
    subject: str | None = None
    keywords: str | None = None
    language: str | None = None


class DocxInspection(BaseModel):
    """Structure summary of a .docx file."""

    path: str = Field(description="Absolute path of the inspected file")
    page: PageSetup = Field(description="Page size, orientation and margins of the first section")
    section_count: int = Field(description="Number of sections in the document")
    counts: DocxCounts
    has_toc: bool = Field(description="True if the body contains a table-of-contents field")
    outline: list[OutlineEntry] = Field(description="Headings in document order")
    properties: DocxProperties
    language: str | None = Field(
        description=(
            "Default proofing language from the document's style defaults "
            "(e.g. 'ru-RU'); null if not set"
        )
    )
    text_preview: str = Field(
        description="The first ~500 characters of body text, paragraphs separated by newlines"
    )
