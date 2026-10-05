"""Инструменты конвертации: markdown_to_word, word_to_markdown, preview_markdown."""

from __future__ import annotations

from contextlib import ExitStack
from pathlib import Path
from typing import TYPE_CHECKING, Annotated

import anyio
from mcp.server.fastmcp import Context
from mcp.types import ToolAnnotations
from pydantic import Field

from ..workflow import resolve_output_paths
from .batch import diagnostics, for_each_file, report_conversions
from .factory import markdown_converter, word_converter
from .models import ConversionReport, FailedFile, PreviewedFile, PreviewReport
from .params import (
    FetchRemoteImages,
    FontName,
    FontSize,
    FootnotesHeading,
    FootnotesParam,
    ImageRoot,
    InputPaths,
    LanguageParam,
    LineBreaksParam,
    OutputDir,
    PageSizeParam,
    PresetParam,
    TemplateParam,
    TocParam,
)
from .paths import prepare_output_dir, resolve_inputs
from .server import mcp

if TYPE_CHECKING:
    from ..converters import MarkdownToWordConverter

_MARKDOWN_TO_WORD_TITLE = "Convert Markdown to Word"


@mcp.tool(
    title=_MARKDOWN_TO_WORD_TITLE,
    annotations=ToolAnnotations(
        title=_MARKDOWN_TO_WORD_TITLE,
        readOnlyHint=False,
        destructiveHint=True,
        idempotentHint=True,
        openWorldHint=True,
    ),
)
async def markdown_to_word(
    inputs: InputPaths,
    output_dir: OutputDir = None,
    font_name: FontName = "Times New Roman",
    font_size: FontSize = None,
    footnotes_heading: FootnotesHeading = "Footnotes",
    fetch_remote_images: FetchRemoteImages = False,
    image_root: ImageRoot = None,
    preset: PresetParam = "default",
    page_size: PageSizeParam = None,
    language: LanguageParam = "auto",
    line_breaks: LineBreaksParam = "soft",
    template: TemplateParam = None,
    toc: TocParam = False,
    footnotes: FootnotesParam = "native",
    ctx: Context | None = None,
) -> ConversionReport:
    """Convert Markdown files to Word .docx documents.

    `inputs` accepts files and directories mixed together; directories are
    scanned recursively for .md and .markdown files. For the Markdown
    features that convert cleanly, read the resource
    `mdtoword://guide/markdown`; run `preview_markdown` first to see what
    would not convert, with line numbers.

    Supports CommonMark + GitHub Flavored Markdown: headings (bookmarked, so
    `[text](#slug)` links work), emphasis, strikethrough, lists, task lists,
    tables, blockquotes, GitHub alerts, code blocks, footnotes, links and
    images, YAML front matter (title, author, date, keywords, lang → title
    block and document properties), a `[TOC]` paragraph (table of contents
    at that spot), definition lists, `~sub~`, `^sup^`, `==highlight==`. LaTeX
    math (`$inline$`, `$$display$$`, amsmath environments, `\\tag{n}`)
    becomes native, editable Word equations (OMML), not images or text —
    check doubtful formulas with `check_latex` first.

    Document options:
    - `preset`: "default" is neutral styling. "gost" follows GOST 7.32-2017:
      A4, margins 30 mm left / 15 mm right / 20 mm top / 20 mm bottom, 1.5
      line spacing, 1.25 cm first-line indent, captions "Рисунок N — …" and
      "Таблица N — …", page numbers centred at the bottom, 14 pt by default.
      "gost_user" is a custom adaptation (not full GOST 7.32 compliance):
      A4, Times New Roman 12 pt black, margins 30/15/15/15 mm, 1.5 spacing,
      justified body, no headers or footers; native footnotes by default.
      Explicit font/page/footnote options win; a template keeps its own
      styles, margins, headers and footers for either preset.
      Supply GOST-formatted bibliography text yourself; no automatic
      bibliography generation or validation is performed.
    - `font_size`: null means the preset default (12 pt, or 14 pt for gost).
    - `page_size`: "A4" or "Letter"; null means A4.
    - `language`: proofing/hyphenation language for Word, e.g. "ru-RU";
      "auto" picks it from the share of Cyrillic text.
    - `line_breaks`: "soft" (CommonMark behaviour) turns a single newline
      inside a paragraph into a space — end a line with a backslash or two
      spaces for a hard break; "preserve" keeps every newline as a line
      break.
    - `template`: path to a reference .docx or .dotx (like pandoc's
      `--reference-doc`); its styles, margins, headers and footers are
      reused, its body text is not copied. Must exist and open as a Word
      document; macro-enabled .dotm/.docm files are rejected.
    - `toc`: insert a table of contents at the start. A `[TOC]` paragraph
      in the source places one there regardless of this flag. Word fills
      it in when the document is opened (update fields).
    - `footnotes`: "native" makes real Word footnotes at the bottom of the
      page; "section" collects them as a numbered list at the end under
      `footnotes_heading`.

    Images referenced by an `http(s)` URL are not fetched by default, and UNC
    (`\\\\host\\share`) or protocol-relative (`//host/...`) paths are not
    touched either; such an image becomes its alt text plus a warning
    instead. This tool reaches the network only when you pass
    `fetch_remote_images=true` — do this only for Markdown from a source you
    trust, since the server fetches using its own network access.

    Images referenced by a local filesystem path are only read from within
    the paths passed in `inputs`: a directory input allows images anywhere
    under it, a file input allows images only next to it (in its parent
    directory), not in sibling directories — naming one file should not
    license wandering elsewhere on disk. An image outside these roots
    becomes its alt text plus a warning, same as a missing file. If your
    images live outside `inputs` (e.g. a shared-assets layout rooted
    elsewhere), pass `image_root` to widen the allowed root to that one
    directory.

    Each output is written next to its source unless `output_dir` is given.
    Existing files at the target paths are overwritten without warning. A
    relative `output_dir` resolves against the server process's working
    directory, not the caller's — pass an absolute path.

    Files are converted one at a time with a progress notification after
    each; one failing file lands in `failed` and does not stop the batch.
    Each warning carries `code` and the 1-based source `line` when known.
    Check `sources_found` in the result: 0 means the paths matched no
    Markdown files at all. Verify an output with `inspect_docx`.
    """

    def prepare() -> tuple[list[Path], dict[Path, Path], MarkdownToWordConverter]:
        # Конвертер (а с ним проверка шаблона) — до создания output_dir:
        # неверный аргумент не должен оставлять следов на диске.
        converter = markdown_converter(
            inputs,
            scratch,
            font_name=font_name,
            font_size=font_size,
            footnotes_heading=footnotes_heading,
            fetch_remote_images=fetch_remote_images,
            image_root=image_root,
            preset=preset,
            page_size=page_size,
            language=language,
            line_breaks=line_breaks,
            template=template,
            toc=toc,
            footnotes=footnotes,
        )
        sources = resolve_inputs(inputs, "md_to_word")
        output_directory = prepare_output_dir(output_dir) if sources else None
        outputs = resolve_output_paths(sources, output_directory, ".docx")
        return sources, outputs, converter

    with ExitStack() as scratch:
        sources, outputs, converter = await anyio.to_thread.run_sync(prepare)
        outcomes = await for_each_file(
            sources, lambda source: converter.convert_file(source, outputs[source]), ctx
        )
    return report_conversions(sources, outputs, outcomes)


_WORD_TO_MARKDOWN_TITLE = "Convert Word to Markdown"


@mcp.tool(
    title=_WORD_TO_MARKDOWN_TITLE,
    annotations=ToolAnnotations(
        title=_WORD_TO_MARKDOWN_TITLE,
        readOnlyHint=False,
        destructiveHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def word_to_markdown(
    inputs: InputPaths,
    output_dir: OutputDir = None,
    extract_media: Annotated[
        bool,
        Field(
            description=(
                "Save pictures to a `<output name>_media/` folder next to each "
                "output and link them from the Markdown; false keeps only their "
                "alt text (with a warning)"
            )
        ),
    ] = True,
    ctx: Context | None = None,
) -> ConversionReport:
    """Convert Word .docx documents to Markdown files.

    `inputs` accepts files and directories mixed together; directories are
    scanned recursively for .docx files.

    Kept: headings (also numbered ones), inline formatting, links (internal
    ones as `#slug` anchors), real Word lists with their numbering and
    nesting, code blocks, quotes and callouts, tables in their original
    position, footnotes and endnotes as `[^n]`, equations as LaTeX
    (`$…$` / `$$…$$`), pictures, and the title/author/keywords as YAML
    front matter. Text is escaped so it reads back as the same text.

    Not representable in Markdown and reported as warnings (one per kind,
    with a count): merged table cells, nested tables, text boxes, OLE
    objects, charts and shapes, the table of contents. Page layout, headers
    and footers are dropped silently.

    With `extract_media` (the default) pictures are written to a
    `<output name>_media/` folder next to each output — created only when
    the document has pictures, overwritten like the output itself.

    Each output is written next to its source unless `output_dir` is given.
    Existing files at the target paths are overwritten without warning. A
    relative `output_dir` resolves against the server process's working
    directory, not the caller's — pass an absolute path.

    Files are converted one at a time with a progress notification after
    each; one failing file lands in `failed` and does not stop the batch.
    Check `sources_found` in the result: 0 means the paths matched no
    .docx files at all.
    """

    def prepare() -> tuple[list[Path], dict[Path, Path]]:
        sources = resolve_inputs(inputs, "word_to_md")
        output_directory = prepare_output_dir(output_dir) if sources else None
        return sources, resolve_output_paths(sources, output_directory, ".md")

    sources, outputs = await anyio.to_thread.run_sync(prepare)
    converter = word_converter(extract_media)
    outcomes = await for_each_file(
        sources, lambda source: converter.convert_file(source, outputs[source]), ctx
    )
    return report_conversions(sources, outputs, outcomes)


_PREVIEW_TITLE = "Preview Markdown → Word conversion"


@mcp.tool(
    title=_PREVIEW_TITLE,
    annotations=ToolAnnotations(
        title=_PREVIEW_TITLE,
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=True,
    ),
)
async def preview_markdown(
    inputs: InputPaths,
    font_name: FontName = "Times New Roman",
    font_size: FontSize = None,
    footnotes_heading: FootnotesHeading = "Footnotes",
    fetch_remote_images: FetchRemoteImages = False,
    image_root: ImageRoot = None,
    preset: PresetParam = "default",
    page_size: PageSizeParam = None,
    language: LanguageParam = "auto",
    line_breaks: LineBreaksParam = "soft",
    template: TemplateParam = None,
    toc: TocParam = False,
    footnotes: FootnotesParam = "native",
    ctx: Context | None = None,
) -> PreviewReport:
    """Check what Markdown would lose in Word, without writing any file.

    Runs the full conversion in memory with exactly the options
    `markdown_to_word` accepts (see its description for what each one
    means) and discards the result, reporting only the warnings: missing or
    refused images, LaTeX that cannot become a Word equation, stripped raw
    HTML, and similar. Each warning has a `code` and, when known, the
    1-based `line` in the Markdown source, so you can fix the source and
    preview again until the list is empty. Use this before
    `markdown_to_word`, or to inspect a document you must not overwrite.

    Nothing is written to disk by this tool. By default it does not fetch
    images referenced by an `http(s)` URL, and does not touch UNC
    (`\\\\host\\share`) or protocol-relative (`//host/...`) paths either --
    such an image becomes its alt text plus a warning instead. This tool
    reaches the network only when you pass `fetch_remote_images=true` — do
    this only for Markdown from a source you trust, since the server fetches
    using its own network access.

    Images referenced by a local filesystem path are only read from within
    the paths passed in `inputs`, with the same rule as `markdown_to_word`:
    a directory input allows images anywhere under it, a file input allows
    images only next to it, not in sibling directories. Pass `image_root`
    to widen the allowed root when your images live elsewhere.
    """

    def prepare() -> tuple[list[Path], MarkdownToWordConverter]:
        converter = markdown_converter(
            inputs,
            scratch,
            font_name=font_name,
            font_size=font_size,
            footnotes_heading=footnotes_heading,
            fetch_remote_images=fetch_remote_images,
            image_root=image_root,
            preset=preset,
            page_size=page_size,
            language=language,
            line_breaks=line_breaks,
            template=template,
            toc=toc,
            footnotes=footnotes,
        )
        return resolve_inputs(inputs, "md_to_word"), converter

    with ExitStack() as scratch:
        sources, converter = await anyio.to_thread.run_sync(prepare)
        outcomes = await for_each_file(sources, converter.preview_file, ctx)
    report = PreviewReport(sources_found=len(sources))
    for source, warnings, error_text in outcomes:
        if error_text is not None:
            report.failed.append(FailedFile(source=str(source), error=error_text))
        else:
            report.previews.append(
                PreviewedFile(source=str(source), warnings=diagnostics(warnings or []))
            )
    return report
