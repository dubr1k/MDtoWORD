# MDtoWord GUI and GFM Conversion Design

## Goal

Replace the line-and-regex Markdown converter with a GFM AST-to-DOCX renderer and modernize the PyQt6 desktop flow around drag-and-drop batches.

## Constraints

- Preserve the existing Markdown-to-Word and Word-to-Markdown mode switch.
- Accept Markdown input in the Markdown-to-Word mode and DOCX input in the reverse mode.
- A missing output-directory selection means that each result is saved beside its own source file.
- When an output directory is selected, all results are written directly there; name collisions receive deterministic numeric suffixes rather than overwriting another pending result.
- Support GFM semantics, not arbitrary vendor-specific Markdown extensions.
- Keep raw HTML and unsupported extensions as literal text; never attempt to execute or interpret them.
- Keep the project distributable through its existing macOS and Windows build flows.

## Architecture

### Markdown conversion

`MarkdownToWordConverter` will become a small orchestration façade:

1. Parse source text through a tokenizing GFM parser.
2. Pass the token stream to a focused DOCX renderer.
3. Return a conversion result that includes success state, output path, and non-fatal rendering warnings.

The renderer maps blocks and inline tokens independently instead of detecting syntax from individual input lines. It renders headings, paragraphs and line breaks, emphasis, strong emphasis, strikethrough, inline code, links, images, quotations, thematic breaks, fenced code blocks, nested ordered and unordered lists, task lists, tables, and footnotes.

The document model cannot represent every Markdown feature directly. Fenced-block language identifiers become captions, and footnotes become a labelled `Сноски` / `Footnotes` section at the end of the document. A failed image download or unsupported media reference produces readable fallback text and a per-file warning without aborting the conversion. Raw HTML is emitted as literal text.

### File discovery and output resolution

The GUI owns a normalized queue of canonical source paths. Files and recursively discovered directory contents are filtered by the active mode, deduplicated, and shown with their source path and queue status. Switching modes removes entries incompatible with the new mode.

The output resolver is shared by both conversions:

- Auto mode: replace the input suffix next to each input file.
- Selected-directory mode: use the selected directory and the converted basename.
- If selected-directory-mode inputs produce duplicate output basenames, append ` (2)`, ` (3)`, and so on before the extension.

### Interface

Use the selected compact single-column layout:

1. Header: app identity and current conversion mode.
2. A prominent drop area that accepts files and directories, with buttons to add files or folders.
3. A queue with filename, original path, and conversion eligibility; users can remove selected entries or clear the queue.
4. Existing font and size controls, visible only in Markdown-to-Word mode.
5. An output card reading `Рядом с исходным файлом` / `Next to each source file` by default, with controls to choose or reset an output directory.
6. Progress, actionable status, and a primary conversion button labelled with the queue count.

Directory drops scan recursively. Without a selected output directory, each result stays beside the file from which it was derived, thereby preserving the source directory structure. With an explicit output directory, results are intentionally flat and collision-safe.

### Errors and reporting

Incompatible drag-and-drop items and directories without matching files are reported in the status area and are not added to the queue. A conversion batch continues after individual-file failures. Its final dialog distinguishes completed files, failures, and rendering warnings.

## Validation

Add deterministic tests for:

- AST rendering of every supported GFM block and inline feature, including nested lists, table alignment, links, images, task-list state, and fallback handling.
- Recursive directory discovery, mode filtering, deduplication, auto output paths, selected-directory output paths, and numeric collision resolution.
- Drag-enter/drop acceptance for supported file and directory URLs, with unsupported items rejected.
- An end-to-end batch that converts a GFM fixture to DOCX and inspects resulting document structure and output paths.

Manually smoke-test the compiled GUI by dragging a directory, leaving the output directory unset, converting it, and confirming each generated file appears beside its source.

## Non-goals

- Parsing, executing, or styling raw HTML embedded in Markdown.
- Supporting every vendor-specific extension outside GFM.
- Changing the existing Word-to-Markdown feature beyond folder-aware input selection and the shared output resolver.
