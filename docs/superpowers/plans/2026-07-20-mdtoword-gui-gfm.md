# MDtoWord GUI and GFM Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Deliver a GFM-to-DOCX conversion pipeline, recursive drag-and-drop input, output-path resolution, and a rebuilt macOS distributable.

**Architecture:** Parse Markdown once into a `markdown-it-py` token stream and render that stream to DOCX through dedicated renderer helpers. Keep queue discovery and output-path allocation independent from PyQt widgets, then connect both to a compact `ConverterGUI` through a drop-aware queue widget.

**Tech Stack:** Python 3.12, PyQt6, python-docx, markdown-it-py, mdit-py-plugins, unittest, PyInstaller.

## Global Constraints

- Markdown contract: GFM; raw HTML and unsupported vendor extensions are literal text.
- Auto output: replace the input suffix beside each source file.
- Explicit output directory: one flat directory with ` (N)` collision suffixes.
- Folder discovery: recursive and filtered by active conversion mode.
- No commit or push; the user will inspect the locally-built app first.

---

### Task 1: Establish testable conversion and queue primitives

**Files:**
- Create: `tests/test_conversion_workflow.py`
- Create: `conversion_workflow.py`
- Modify: `requirements.txt`

**Interfaces:**
- Produces `supported_suffixes(mode: str) -> frozenset[str]`.
- Produces `discover_sources(paths: Iterable[Path], mode: str) -> list[Path]`.
- Produces `resolve_output_paths(inputs: Sequence[Path], output_directory: Path | None, suffix: str) -> dict[Path, Path]`.

- [ ] **Step 1: Write failing workflow tests**

```python
from pathlib import Path
from conversion_workflow import discover_sources, resolve_output_paths


def test_discovery_recurses_and_deduplicates_md_sources(tmp_path: Path):
    nested = tmp_path / "nested"
    nested.mkdir()
    first = tmp_path / "first.md"
    second = nested / "second.MD"
    ignored = nested / "ignored.txt"
    for path in (first, second, ignored):
        path.write_text("# test", encoding="utf-8")

    assert discover_sources([tmp_path, first], "md_to_word") == [first.resolve(), second.resolve()]


def test_selected_output_directory_allocates_collision_suffixes(tmp_path: Path):
    left = tmp_path / "left" / "report.md"
    right = tmp_path / "right" / "report.md"
    output = tmp_path / "output"
    left.parent.mkdir(parents=True)
    right.parent.mkdir(parents=True)
    output.mkdir()

    paths = resolve_output_paths([left, right], output, ".docx")

    assert paths[left] == output / "report.docx"
    assert paths[right] == output / "report (2).docx"
```

- [ ] **Step 2: Run the new tests and verify RED**

Run: `python -m unittest tests.test_conversion_workflow -v`

Expected: import failure because `conversion_workflow` does not yet exist.

- [ ] **Step 3: Implement pure discovery and output resolution**

```python
from collections.abc import Iterable, Sequence
from pathlib import Path

_SUFFIXES = {"md_to_word": frozenset({".md", ".markdown"}), "word_to_md": frozenset({".docx"})}


def supported_suffixes(mode: str) -> frozenset[str]:
    return _SUFFIXES[mode]


def discover_sources(paths: Iterable[Path], mode: str) -> list[Path]:
    suffixes = supported_suffixes(mode)
    candidates = (child for path in paths for child in (path.rglob("*") if path.is_dir() else (path,)))
    return sorted({path.resolve() for path in candidates if path.is_file() and path.suffix.lower() in suffixes})


def resolve_output_paths(inputs: Sequence[Path], output_directory: Path | None, suffix: str) -> dict[Path, Path]:
    allocated: set[Path] = set()
    outputs: dict[Path, Path] = {}
    for source in inputs:
        directory = output_directory or source.parent
        candidate = directory / f"{source.stem}{suffix}"
        index = 2
        while candidate in allocated:
            candidate = directory / f"{source.stem} ({index}){suffix}"
            index += 1
        allocated.add(candidate)
        outputs[source] = candidate
    return outputs
```

- [ ] **Step 4: Run workflow tests and verify GREEN**

Run: `python -m unittest tests.test_conversion_workflow -v`

Expected: 2 tests pass.

### Task 2: Replace the regex parser with a GFM token renderer

**Files:**
- Create: `gfm_docx_renderer.py`
- Create: `tests/test_gfm_docx_renderer.py`
- Modify: `md_to_word_converter.py:22-298`
- Modify: `requirements.txt`

**Interfaces:**
- Produces `GfmDocxRenderer(font_name: str, font_size: Pt)`.
- Produces `GfmDocxRenderer.render(markdown: str, source_path: Path | None = None) -> tuple[Document, list[str]]`.
- `MarkdownToWordConverter.convert_content` saves the returned document and reports renderer warnings.

- [ ] **Step 1: Write a failing renderer contract test**

```python
from docx.shared import Pt
from gfm_docx_renderer import GfmDocxRenderer


def test_renderer_preserves_gfm_blocks_and_inline_semantics(tmp_path):
    document, warnings = GfmDocxRenderer("Arial", Pt(12)).render(
        "# Heading\n\n**bold** and ~~old~~ [link](https://example.com)\n\n- [x] done\n  - child\n\n| a | b |\n| :- | -: |\n| 1 | 2 |\n\n> quote\n\n```python\nprint('x')\n```"
    )

    assert document.paragraphs[0].style.name == "Heading 1"
    assert any(run.bold for run in document.paragraphs[1].runs)
    assert "☒ done" in "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert len(document.tables) == 1
    assert warnings == []
```

- [ ] **Step 2: Run renderer test and verify RED**

Run: `python -m unittest tests.test_gfm_docx_renderer -v`

Expected: import failure because `gfm_docx_renderer` does not yet exist.

- [ ] **Step 3: Implement parser setup and renderer visitor**

Use `MarkdownIt("gfm-like2", {"breaks": True, "html": False})`, add `footnote_plugin`, and visit open/close tokens with an explicit context stack for list level, block quote level, tables, and inline runs. Inline handling maps `text`, `softbreak`, `hardbreak`, `em`, `strong`, `s`, `code_inline`, `link_open`/`link_close`, and `image`; unsupported token types append their content as literal text. Fenced code emits an optional language caption and a Courier New paragraph. Image retrieval accepts local paths relative to `source_path` and HTTP(S) URLs, reporting a warning plus alt text on failure.

- [ ] **Step 4: Route MarkdownToWordConverter through the renderer**

Replace `parse_inline_formatting`, `add_formatted_text`, `process_table`, `process_list`, and the line scanner in `convert_content` with a renderer call. Preserve the public `convert_content` and `convert_file` return contract; append non-fatal warnings to its success message.

- [ ] **Step 5: Run renderer tests and verify GREEN**

Run: `python -m unittest tests.test_gfm_docx_renderer -v`

Expected: all renderer tests pass.

### Task 3: Add a drop-aware compact file queue

**Files:**
- Modify: `md_to_word_converter.py:378-1015`
- Create: `tests/test_drop_queue.py`

**Interfaces:**
- Produces `DropFileList(QListWidget)` with `paths_dropped: pyqtSignal(list)`.
- `ConverterGUI._add_sources(paths: Iterable[Path]) -> None` updates the canonical queue.
- `ConverterGUI._set_output_directory(directory: Path | None) -> None` updates the output card.

- [ ] **Step 1: Write failing drag/drop acceptance test**

```python
from PyQt6.QtCore import QMimeData, QPoint, QUrl, Qt
from PyQt6.QtGui import QDragEnterEvent
from PyQt6.QtWidgets import QApplication
from md_to_word_converter import DropFileList


APP = QApplication.instance() or QApplication([])


def test_drop_file_list_accepts_local_file_urls():
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile("/tmp/input.md")])
    event = QDragEnterEvent(
        QPoint(0, 0), Qt.DropAction.CopyAction, mime,
        Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
    )

    widget = DropFileList()
    widget.dragEnterEvent(event)

    assert event.isAccepted()
```

- [ ] **Step 2: Run test and verify RED**

Run: `python -m unittest tests.test_drop_queue -v`

Expected: import failure because `DropFileList` does not yet exist.

- [ ] **Step 3: Implement the focused GUI flow**

Replace the generic file-list widget with `DropFileList`; add `Добавить файлы` and `Добавить папку` buttons; route their selections and dropped URLs into `discover_sources`. Render each queue entry with filename and canonical source path. The default output card reads `Рядом с исходным файлом`; choosing an output directory adds a reset action. Keep existing mode and language controls, but discard incompatible queued sources whenever mode changes.

- [ ] **Step 4: Use resolved paths during batch conversion**

At conversion start, call `resolve_output_paths` once for the entire queue. Continue each batch after individual-file failures, show progress, and include warnings in the final result dialog.

- [ ] **Step 5: Run all focused tests and verify GREEN**

Run: `python -m unittest tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_drop_queue -v`

Expected: all tests pass.

### Task 4: Validate, package, and smoke-test macOS app

**Files:**
- Modify: `MDtoWORD.spec` only if PyInstaller needs hidden imports for installed Markdown packages.
- Build: `dist/MDtoWORD.app`

- [ ] **Step 1: Run the complete automated suite**

Run: `python -m unittest discover -s tests -v`

Expected: 0 failures and 0 errors.

- [ ] **Step 2: Run compile diagnostics**

Run: `python -m compileall -q md_to_word_converter.py conversion_workflow.py gfm_docx_renderer.py`

Expected: exit code 0.

- [ ] **Step 3: Build the signed local macOS app**

Run: `./scripts/build_macos.sh`

Expected: `codesign --verify --deep --strict` succeeds and `dist/MDtoWORD.app` exists.

- [ ] **Step 4: Smoke-test the application bundle**

Launch `dist/MDtoWORD.app`, drag a fixture directory containing nested Markdown files, leave output-directory mode automatic, convert, and verify that every generated DOCX appears beside its source.
