# MCP-сервер MDtoWORD — план реализации

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Дать MCP-агентам конвертацию Markdown ↔ Word через три инструмента поверх существующего движка рендеринга.

**Architecture:** Ядро конвертации извлекается из `mdtoword/app.py` (который импортирует PyQt6) в Qt-free модуль `mdtoword/converters.py` с новым контрактом: успех возвращает `list[str]` варнингов, отказ бросает `ConversionError`. GUI и MCP-сервер становятся двумя равноправными потребителями этого ядра. Сервер — stdio, на официальном SDK `mcp`.

**Tech Stack:** Python 3.12, `mcp==1.28.1` (`mcp.server.fastmcp.FastMCP`), Pydantic v2, python-docx, markdown-it-py, unittest (стиль репозитория).

**Спек:** `docs/design/specs/2026-07-21-mcp-server-design.md`

## Global Constraints

- Никакой код в цепочке `mcp_server → converters → gfm_renderer → latex_omml` не смеет импортировать PyQt6.
- stdout — транспорт stdio. Ни `print`, ни вывод библиотек в stdout ни при импорте, ни при выполнении. Логи только в stderr.
- Тесты — `unittest.TestCase` / `unittest.IsolatedAsyncioTestCase`, без pytest-плагинов: в репозитории нет ни `conftest.py`, ни `pytest.ini`.
- `mcp` идёт **только** в новый `requirements-mcp.txt`. `requirements.txt` и `environment.yml` не трогать — их попарное соответствие охраняет `tests/test_packaging.py`.
- Существующие имена атрибутов `default_font_name`, `default_font_size`, `footnotes_heading` сохраняются: GUI присваивает их напрямую после конструирования (`app.py:436-438`, `510`, `514`).
- Перезапись выходных файлов — молча, по решению из спека. Флаг `overwrite` не вводить.
- Команда прогона тестов: `.venv-macos-build/bin/python -m pytest tests/ -q` из корня репозитория. Именно этот интерпретатор, а не `python` из PATH: под anaconda-сборкой Qt GUI-тесты роняют интерпретатор segfault'ом, и на голом `python` набор выглядит «зелёным» лишь потому, что `test_drop_queue.py` и `test_gui_theme.py` до утверждений не доживают. В `.venv-macos-build` набор даёт 177 passed.

---

### Task 1: Qt-free ядро конвертации

Извлечение `MarkdownToWordConverter` и `WordToMarkdownConverter` из `app.py` в отдельный модуль с новым контрактом ошибок. Поведение рендеринга не меняется — меняется только форма возврата.

**Files:**
- Create: `mdtoword/converters.py`
- Create: `tests/test_converters.py`
- Delete: `tests/test_markdown_converter.py`

**Interfaces:**
- Consumes: `mdtoword.gfm_renderer.GfmDocxRenderer(font_name: str, font_size: Pt, footnotes_heading: str)` с методом `.render(markdown: str, source_path: Path | None) -> tuple[Document, list[str]]`.
- Produces:
  - `ConversionError(Exception)`
  - `MarkdownToWordConverter(font_name: str = "Times New Roman", font_size: Pt = Pt(12), footnotes_heading: str = "Footnotes")`
    - `.convert_content(content: str, output_path: str | Path, source_path: Path | None = None) -> list[str]`
    - `.convert_file(input_path: str | Path, output_path: str | Path) -> list[str]`
    - публичные атрибуты `default_font_name`, `default_font_size`, `footnotes_heading`
  - `WordToMarkdownConverter()` с `.convert_file(input_path: str | Path, output_path: str | Path) -> list[str]` (всегда возвращает `[]`)

- [ ] **Step 1: Написать падающий тест**

Создать `tests/test_converters.py`:

```python
"""Тесты ядра конвертации.

Модуль намеренно не импортирует ``mdtoword.app``: ядро не должно зависеть
от PyQt6, и этот файл — исполняемое доказательство того, что не зависит.
"""

from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from docx import Document

from mdtoword.converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)


class MarkdownToWordConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    def test_convert_content_writes_the_document_and_returns_no_warnings(self) -> None:
        output = self.root / "out.docx"

        warnings = MarkdownToWordConverter().convert_content("# Заголовок", output)

        self.assertEqual(warnings, [])
        self.assertEqual(Document(str(output)).paragraphs[0].text, "Заголовок")

    def test_convert_file_returns_nonfatal_renderer_warnings_as_a_list(self) -> None:
        source = self.root / "source.md"
        source.write_text("![diagram](missing.png)", encoding="utf-8")
        output = self.root / "source.docx"

        warnings = MarkdownToWordConverter().convert_file(source, output)

        self.assertEqual(warnings, ["Image not found: missing.png"])
        self.assertTrue(output.is_file())

    def test_missing_source_raises_conversion_error(self) -> None:
        with self.assertRaises(ConversionError):
            MarkdownToWordConverter().convert_file(
                self.root / "нет-такого.md", self.root / "out.docx"
            )

    def test_unwritable_output_raises_conversion_error(self) -> None:
        source = self.root / "source.md"
        source.write_text("# Заголовок", encoding="utf-8")
        # Каталог вместо файла: python-docx не может писать поверх директории.
        blocked = self.root / "blocked.docx"
        blocked.mkdir()

        with self.assertRaises(ConversionError):
            MarkdownToWordConverter().convert_file(source, blocked)

    def test_font_settings_reach_the_rendered_document(self) -> None:
        output = self.root / "styled.docx"

        MarkdownToWordConverter(font_name="Georgia").convert_content("Текст", output)

        document = Document(str(output))
        self.assertEqual(document.styles["Normal"].font.name, "Georgia")


class WordToMarkdownConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    def test_headings_and_bold_runs_round_trip_into_markdown(self) -> None:
        docx_path = self.root / "source.docx"
        document = Document()
        document.add_heading("Раздел", level=2)
        paragraph = document.add_paragraph()
        paragraph.add_run("жирный").bold = True
        document.save(str(docx_path))
        output = self.root / "source.md"

        warnings = WordToMarkdownConverter().convert_file(docx_path, output)

        text = output.read_text(encoding="utf-8")
        self.assertEqual(warnings, [])
        self.assertIn("## Раздел", text)
        self.assertIn("**жирный**", text)

    def test_missing_source_raises_conversion_error(self) -> None:
        with self.assertRaises(ConversionError):
            WordToMarkdownConverter().convert_file(
                self.root / "нет-такого.docx", self.root / "out.md"
            )


class CoreIsolationTests(unittest.TestCase):
    def test_importing_the_core_does_not_pull_in_pyqt(self) -> None:
        # Обязательно в подпроцессе: в общем прогоне tests/test_drop_queue.py
        # импортирует PyQt6, и проверка sys.modules в текущем процессе
        # проходила бы или падала в зависимости от порядка сбора тестов.
        repo_root = Path(__file__).resolve().parent.parent
        result = subprocess.run(
            [
                sys.executable,
                "-c",
                "import mdtoword.converters, sys; "
                "sys.exit(1 if 'PyQt6' in sys.modules else 0)",
            ],
            capture_output=True,
            text=True,
            cwd=repo_root,
        )

        self.assertEqual(result.returncode, 0, "ядро подтянуло PyQt6:\n" + result.stderr)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест и убедиться, что он падает**

Run: `python -m pytest tests/test_converters.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'mdtoword.converters'`

- [ ] **Step 3: Реализовать модуль**

Создать `mdtoword/converters.py`:

```python
"""Ядро конвертации Markdown ↔ Word.

Модуль намеренно свободен от PyQt6: им пользуются и GUI (``mdtoword.app``),
и MCP-сервер (``mdtoword.mcp_server``), а последний обязан работать без
графической подсистемы.

Контракт: успешная конвертация возвращает список необязательных предупреждений,
неуспешная — бросает :class:`ConversionError`. Сообщения об ошибках намеренно
не локализованы: язык подставляет потребитель, у GUI для этого есть словарь
переводов, а агенту локаль не нужна вовсе.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.shared import Pt

from .gfm_renderer import GfmDocxRenderer


class ConversionError(Exception):
    """Конвертация не удалась. Сообщение пригодно для показа без перевода."""


class MarkdownToWordConverter:
    """Конвертирует GFM-разметку в документ Word."""

    def __init__(
        self,
        font_name: str = "Times New Roman",
        font_size: Pt = Pt(12),
        footnotes_heading: str = "Footnotes",
    ) -> None:
        self.default_font_name = font_name
        self.default_font_size = font_size
        self.footnotes_heading = footnotes_heading

    def convert_content(
        self, content: str, output_path: str | Path, source_path: Path | None = None
    ) -> list[str]:
        """Отрендерить Markdown и сохранить результат в *output_path*."""
        try:
            document, warnings = GfmDocxRenderer(
                self.default_font_name, self.default_font_size, self.footnotes_heading
            ).render(content, source_path=source_path)
            document.save(str(output_path))
        except Exception as error:
            raise ConversionError(str(error)) from error
        return warnings

    def convert_file(
        self, input_path: str | Path, output_path: str | Path
    ) -> list[str]:
        """Прочитать Markdown-файл и сконвертировать его."""
        source_path = Path(input_path)
        try:
            content = source_path.read_text(encoding="utf-8")
        except OSError as error:
            raise ConversionError(str(error)) from error
        return self.convert_content(content, output_path, source_path)


class WordToMarkdownConverter:
    """Извлекает Markdown из документа Word.

    Преобразование лоссовое: распознаются заголовки по стилям ``Heading N``,
    жирный и курсивный текст по run'ам и таблицы. Списки, изображения,
    формулы и сноски схлопываются в плоский текст.
    """

    def convert_file(
        self, input_path: str | Path, output_path: str | Path
    ) -> list[str]:
        """Сконвертировать документ Word в Markdown-файл."""
        try:
            document = Document(str(input_path))
            lines = self._paragraph_lines(document)
            lines.extend(self._table_lines(document))
            Path(output_path).write_text("\n".join(lines), encoding="utf-8")
        except Exception as error:
            raise ConversionError(str(error)) from error
        return []

    def _paragraph_lines(self, document: Any) -> list[str]:
        lines: list[str] = []
        for paragraph in document.paragraphs:
            if not paragraph.text.strip():
                lines.append("")
                continue
            heading_level = self._heading_level(paragraph)
            if heading_level:
                lines.append("#" * heading_level + " " + paragraph.text)
            else:
                lines.append(self._inline_markup(paragraph))
        return lines

    @staticmethod
    def _heading_level(paragraph: Any) -> int:
        """Вернуть уровень заголовка 1..6 или 0, если это не заголовок."""
        style = paragraph.style
        style_name = (style.name or "") if style is not None else ""
        if not style_name.startswith("Heading "):
            return 0
        try:
            level = int(style_name.split()[-1])
        except ValueError:
            return 0
        return level if 1 <= level <= 6 else 0

    @staticmethod
    def _inline_markup(paragraph: Any) -> str:
        parts: list[str] = []
        for run in paragraph.runs:
            text = run.text
            if run.bold and run.italic:
                parts.append(f"***{text}***")
            elif run.bold:
                parts.append(f"**{text}**")
            elif run.italic:
                parts.append(f"*{text}*")
            else:
                parts.append(text)
        return "".join(parts)

    @staticmethod
    def _table_lines(document: Any) -> list[str]:
        lines: list[str] = []
        for table in document.tables:
            if not table.rows:
                continue
            lines.append("")
            header = [cell.text.strip() for cell in table.rows[0].cells]
            lines.append("| " + " | ".join(header) + " |")
            lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            for row in table.rows[1:]:
                lines.append("| " + " | ".join(cell.text.strip() for cell in row.cells) + " |")
            lines.append("")
        return lines
```

Разница с оригиналом из `app.py:58-133`, помимо контракта ошибок: тело `convert_file` разбито на четыре именованных метода вместо одного блока с цикломатической сложностью 52 (`health://complexity` помечает его как hotspot), и выброшена мёртвая ветка `run_text.strip().startswith('`')` — она недостижима, поскольку проверялась после `run.bold`/`run.italic` и никогда не срабатывала для обычного текста с backtick'ами, а python-docx не выставляет для inline-кода ни один из этих флагов.

- [ ] **Step 4: Запустить тест и убедиться, что он проходит**

Run: `python -m pytest tests/test_converters.py -q`
Expected: PASS, 8 passed

- [ ] **Step 5: Удалить заменённый тест**

Файл `tests/test_markdown_converter.py` целиком поглощён новым: его единственная проверка (варнинг про `missing.png`) воспроизведена в `test_convert_file_returns_nonfatal_renderer_warnings_as_a_list` уже без импорта PyQt6.

```bash
git rm tests/test_markdown_converter.py
```

- [ ] **Step 6: Убедиться, что остальной набор цел**

Run: `python -m pytest tests/ -q`
Expected: PASS. `mdtoword/app.py` пока держит собственные копии классов — это ожидаемо и снимается Задачей 2.

- [ ] **Step 7: Коммит**

```bash
git add mdtoword/converters.py tests/test_converters.py
git commit -m "refactor: extract a Qt-free conversion core with structured errors"
```

---

### Task 2: Перевести GUI на новое ядро

`app.py` перестаёт объявлять конвертеры и начинает их импортировать. Локализация сообщений переезжает из ядра в GUI.

**Files:**
- Modify: `mdtoword/app.py:1-19` (импорты), `21-133` (удаление классов), `241-242` и `262-263` (ключи переводов), `585-598` (цикл батча), `614-629` (конвертация текста)
- Test: `tests/test_drop_queue.py` (существующий, должен остаться зелёным)

**Interfaces:**
- Consumes: `ConversionError`, `MarkdownToWordConverter`, `WordToMarkdownConverter` из Задачи 1.
- Produces: ничего для последующих задач. `mdtoword.app` остаётся импортируемым как `from mdtoword.app import ConverterGUI, DropFileList` (`tests/test_drop_queue.py:13`) и реэкспортирует имена конвертеров для обратной совместимости этого импорта.

- [ ] **Step 1: Удалить классы конвертеров и импортировать их из ядра**

В `mdtoword/app.py` удалить строки 21–133 целиком (оба класса `MarkdownToWordConverter` и `WordToMarkdownConverter` вместе с пустой строкой 134-135 перед `_dropped_local_paths`) и заменить блок импортов строк 16–18 на:

```python
from .converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)
from .workflow import discover_sources, resolve_output_paths
from .gfm_renderer import GfmDocxRenderer
from .theme import ThemeManager
```

Импорт `from docx import Document` (строка 5) больше не нужен — `Document` использовался только внутри `WordToMarkdownConverter`. Удалить его. `from docx.shared import Pt` (строка 6) остаётся: он нужен в `_toggle_converter_type` и `_on_size_change`. Импорт `GfmDocxRenderer` тоже больше не используется напрямую — удалить и его.

Итоговый блок импортов строк 1–19:

```python
import sys
from typing import Any, cast
from pathlib import Path

from docx.shared import Pt
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent, QIcon, QMouseEvent
from PyQt6.QtWidgets import (
    QApplication, QAbstractItemView, QComboBox, QFileDialog, QGroupBox,
    QHBoxLayout, QLabel, QListWidget, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QSpinBox, QTabWidget,
    QVBoxLayout, QWidget,
)

from .converters import (
    ConversionError,
    MarkdownToWordConverter,
    WordToMarkdownConverter,
)
from .workflow import discover_sources, resolve_output_paths
from .theme import ThemeManager
```

- [ ] **Step 2: Добавить ключи локализации для результата конвертации**

Раньше строки «Успешно конвертировано» и «Ошибка при конвертации: …» приходили из ядра и всегда были русскими — даже когда интерфейс переключён на английский. Теперь их формирует GUI.

В русский словарь после строки `"footnotes_heading": "Сноски",` (`app.py:242`) добавить:

```python
                "converted_ok": "Успешно конвертировано",
                "convert_failed": "Ошибка при конвертации: {error}",
```

В английский словарь после `"footnotes_heading": "Footnotes",` (`app.py:263`) добавить:

```python
                "converted_ok": "Converted successfully",
                "convert_failed": "Conversion failed: {error}",
```

- [ ] **Step 3: Переписать цикл батча на структурированные варнинги**

Заменить строки 585–598 (`for index, source in enumerate(...)` до `details = errors + warnings` включительно) на:

```python
            for index, source in enumerate(queue, start=1):
                self.status_label.setText(self._text["converting"].format(filename=source.name))
                QApplication.processEvents()
                try:
                    file_warnings = self.converter.convert_file(source, outputs[source])
                except ConversionError as error:
                    errors.append(
                        f"{source.name}: " + self._text["convert_failed"].format(error=error)
                    )
                else:
                    success_count += 1
                    warnings.extend(f"{source.name}: {warning}" for warning in file_warnings)
                self.progress.setValue(index)
                QApplication.processEvents()

            details = errors + warnings
```

Это снимает главную хрупкость старого кода: он определял наличие варнингов подстрокой `"Warnings:" in message` и вытаскивал их через `message.split("Warnings:", 1)` — то есть парсил обратно текст, который сам же и склеил, и разваливался бы от любого исходника, где слово «Warnings:» встречается в теле документа.

- [ ] **Step 4: Переписать конвертацию из текстовой вкладки**

Заменить строки 625–629 (от `success, message = cast(...)` до конца метода) на:

```python
        try:
            warnings = cast(MarkdownToWordConverter, self.converter).convert_content(content, output)
        except ConversionError as error:
            QMessageBox.critical(
                self, self._text["errors"], self._text["convert_failed"].format(error=error)
            )
            return
        message = self._text["converted_ok"]
        if warnings:
            message += "\n\n" + "\n".join(warnings)
        QMessageBox.information(self, self.windowTitle(), message)
```

- [ ] **Step 5: Прогнать весь набор**

Run: `python -m pytest tests/ -q`
Expected: PASS. `tests/test_drop_queue.py` поднимает `ConverterGUI` — он проверяет, что импорты и конструктор целы.

- [ ] **Step 6: Проверить, что GUI запускается**

Run: `python -m mdtoword`
Expected: открывается окно. Сконвертировать один `.md` через вкладку «Текст» и убедиться, что показывается «Успешно конвертировано», а при переключении языка — «Converted successfully». Закрыть окно.

- [ ] **Step 7: Коммит**

```bash
git add mdtoword/app.py
git commit -m "refactor: point the GUI at the shared core and localise its own messages"
```

---

### Task 3: MCP-сервер с двумя инструментами конвертации

**Files:**
- Create: `requirements-mcp.txt`
- Create: `mdtoword/mcp_server.py`
- Create: `tests/test_mcp_server.py`

**Interfaces:**
- Consumes: ядро из Задачи 1; `workflow.discover_sources(paths: Iterable[Path], mode: str) -> list[Path]` и `workflow.resolve_output_paths(inputs: Sequence[Path], output_directory: Path | None, suffix: str) -> dict[Path, Path]`.
- Produces:
  - модульный объект `mcp: FastMCP` (Задача 4 регистрирует на нём третий инструмент);
  - `_resolve_inputs(inputs: list[str], mode: str) -> list[Path]`;
  - `_prepare_output_dir(output_dir: str | None) -> Path | None`;
  - модели `ConvertedFile`, `FailedFile`, `ConversionReport`;
  - `main() -> None`.

- [ ] **Step 1: Зафиксировать зависимость**

Создать `requirements-mcp.txt`:

```
# Зависимости MCP-сервера (mdtoword/mcp_server.py).
# Намеренно отдельно от requirements.txt: тот описывает GUI-приложение,
# сверяется с environment.yml тестом tests/test_packaging.py и попадает
# в сборку PyInstaller, которой серверный SDK не нужен.
-r requirements.txt
mcp==1.28.1
```

Установить: `python -m pip install -r requirements-mcp.txt`

- [ ] **Step 2: Написать падающий тест**

Создать `tests/test_mcp_server.py`:

```python
"""Тесты MCP-сервера через in-memory клиент из SDK.

Клиент и сервер соединяются напрямую в одном процессе, без подпроцесса
и без stdio, — проверяется реальный путь вызова инструмента вместе со
схемами и валидацией аргументов.
"""

from pathlib import Path
import tempfile
import unittest

from docx import Document

try:
    from mcp.shared.memory import (
        create_connected_server_and_client_session as client_session,
    )
except ImportError:  # pragma: no cover
    raise unittest.SkipTest("SDK mcp не установлен; см. requirements-mcp.txt")

from mdtoword.mcp_server import mcp as server


class McpServerTestCase(unittest.IsolatedAsyncioTestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    async def call(self, tool: str, arguments: dict):
        async with client_session(server._mcp_server) as client:
            return await client.call_tool(tool, arguments)


class ToolRegistrationTests(McpServerTestCase):
    async def test_both_conversion_tools_are_advertised_with_descriptions(self) -> None:
        async with client_session(server._mcp_server) as client:
            listed = await client.list_tools()

        tools = {tool.name: tool for tool in listed.tools}
        self.assertIn("markdown_to_word", tools)
        self.assertIn("word_to_markdown", tools)
        for tool in tools.values():
            self.assertTrue(tool.description)

    async def test_the_lossy_direction_says_so_in_its_description(self) -> None:
        async with client_session(server._mcp_server) as client:
            listed = await client.list_tools()

        description = next(t.description for t in listed.tools if t.name == "word_to_markdown")
        self.assertIn("lossy", description.lower())


class MarkdownToWordTests(McpServerTestCase):
    async def test_a_directory_is_converted_recursively_with_one_output_each(self) -> None:
        nested = self.root / "nested"
        nested.mkdir()
        (self.root / "first.md").write_text("# Первый", encoding="utf-8")
        (nested / "second.markdown").write_text("# Второй", encoding="utf-8")
        (self.root / "ignored.txt").write_text("не markdown", encoding="utf-8")

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        self.assertFalse(result.isError)
        report = result.structuredContent
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["converted"]), 2)
        self.assertEqual(report["failed"], [])
        for entry in report["converted"]:
            self.assertTrue(Path(entry["output"]).is_file())

    async def test_output_dir_is_created_and_used(self) -> None:
        (self.root / "doc.md").write_text("# Заголовок", encoding="utf-8")
        destination = self.root / "out" / "deep"

        result = await self.call(
            "markdown_to_word",
            {"inputs": [str(self.root / "doc.md")], "output_dir": str(destination)},
        )

        report = result.structuredContent
        self.assertEqual(Path(report["converted"][0]["output"]).parent, destination)
        self.assertTrue((destination / "doc.docx").is_file())

    async def test_nonfatal_warnings_are_reported_per_file(self) -> None:
        (self.root / "doc.md").write_text("![diagram](missing.png)", encoding="utf-8")

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(
            report["converted"][0]["warnings"], ["Image not found: missing.png"]
        )

    async def test_font_arguments_reach_the_document(self) -> None:
        (self.root / "doc.md").write_text("Текст", encoding="utf-8")

        result = await self.call(
            "markdown_to_word",
            {"inputs": [str(self.root)], "font_name": "Georgia", "font_size": 14},
        )

        output = Path(result.structuredContent["converted"][0]["output"])
        document = Document(str(output))
        self.assertEqual(document.styles["Normal"].font.name, "Georgia")

    async def test_paths_matching_nothing_report_zero_sources_found(self) -> None:
        result = await self.call(
            "markdown_to_word", {"inputs": [str(self.root / "нет-такой-папки")]}
        )

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 0)
        self.assertEqual(report["converted"], [])
        self.assertEqual(report["failed"], [])

    async def test_empty_inputs_is_an_error_not_an_empty_success(self) -> None:
        result = await self.call("markdown_to_word", {"inputs": []})

        self.assertTrue(result.isError)


class WordToMarkdownTests(McpServerTestCase):
    def write_docx(self, name: str) -> Path:
        path = self.root / name
        document = Document()
        document.add_heading("Раздел", level=1)
        document.save(str(path))
        return path

    async def test_documents_are_converted_to_markdown_files(self) -> None:
        self.write_docx("report.docx")

        result = await self.call("word_to_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 1)
        output = Path(report["converted"][0]["output"])
        self.assertEqual(output.suffix, ".md")
        self.assertIn("# Раздел", output.read_text(encoding="utf-8"))

    async def test_a_broken_file_fails_alone_without_stopping_the_batch(self) -> None:
        self.write_docx("good.docx")
        (self.root / "broken.docx").write_text("это не zip-контейнер", encoding="utf-8")

        result = await self.call("word_to_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertFalse(result.isError)
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["converted"]), 1)
        self.assertEqual(len(report["failed"]), 1)
        self.assertTrue(report["failed"][0]["source"].endswith("broken.docx"))
        self.assertTrue(report["failed"][0]["error"])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Запустить тест и убедиться, что он падает**

Run: `python -m pytest tests/test_mcp_server.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'mdtoword.mcp_server'`

- [ ] **Step 4: Реализовать сервер**

Создать `mdtoword/mcp_server.py`:

```python
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
    """Один успешно сконвертированный файл."""

    source: str = Field(description="Absolute path of the input file")
    output: str = Field(description="Absolute path of the file that was written")
    warnings: list[str] = Field(
        default_factory=list,
        description="Non-fatal issues; the output was still written",
    )


class FailedFile(BaseModel):
    """Один файл, который сконвертировать не удалось."""

    source: str = Field(description="Absolute path of the input file")
    error: str = Field(description="Why this file could not be converted")


class ConversionReport(BaseModel):
    """Итог пакетной конвертации."""

    sources_found: int = Field(
        description=(
            "How many supported files the inputs resolved to. "
            "0 means the paths matched nothing — check the paths rather than "
            "assuming there was nothing to do."
        )
    )
    converted: list[ConvertedFile] = Field(default_factory=list)
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
    directory = Path(output_dir).expanduser()
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
```

- [ ] **Step 5: Запустить тест и убедиться, что он проходит**

Run: `python -m pytest tests/test_mcp_server.py -q`
Expected: PASS, 10 passed

- [ ] **Step 6: Прогнать весь набор**

Run: `python -m pytest tests/ -q`
Expected: PASS

- [ ] **Step 7: Коммит**

```bash
git add requirements-mcp.txt mdtoword/mcp_server.py tests/test_mcp_server.py
git commit -m "feat: serve markdown and word conversion over MCP"
```

---

### Task 4: Инструмент preview и защита stdio-протокола

Preview прогоняет полный рендер в память и ничего не пишет — агент видит предупреждения до того, как создаст файл.

**Files:**
- Modify: `mdtoword/converters.py` (добавить `preview_content` и `preview_file` в `MarkdownToWordConverter`)
- Modify: `mdtoword/mcp_server.py` (добавить модели `PreviewedFile`, `PreviewReport` и инструмент `preview_markdown`)
- Modify: `tests/test_converters.py` (тесты preview в ядре)
- Modify: `tests/test_mcp_server.py` (тесты инструмента и чистоты stdout)

**Interfaces:**
- Consumes: всё из Задач 1 и 3.
- Produces: `MarkdownToWordConverter.preview_content(content: str, source_path: Path | None = None) -> list[str]` и `.preview_file(input_path: str | Path) -> list[str]`.

- [ ] **Step 1: Написать падающие тесты ядра**

Добавить в `tests/test_converters.py` внутрь класса `MarkdownToWordConverterTests`:

```python
    def test_preview_reports_warnings_without_writing_anything(self) -> None:
        source = self.root / "source.md"
        source.write_text("![diagram](missing.png)", encoding="utf-8")
        before = sorted(path.name for path in self.root.iterdir())

        warnings = MarkdownToWordConverter().preview_file(source)

        self.assertEqual(warnings, ["Image not found: missing.png"])
        self.assertEqual(sorted(path.name for path in self.root.iterdir()), before)

    def test_preview_of_clean_markdown_returns_no_warnings(self) -> None:
        warnings = MarkdownToWordConverter().preview_content("# Заголовок\n\nТекст.")

        self.assertEqual(warnings, [])

    def test_preview_of_a_missing_file_raises_conversion_error(self) -> None:
        with self.assertRaises(ConversionError):
            MarkdownToWordConverter().preview_file(self.root / "нет-такого.md")
```

- [ ] **Step 2: Запустить и убедиться, что падает**

Run: `python -m pytest tests/test_converters.py -q`
Expected: FAIL — `AttributeError: 'MarkdownToWordConverter' object has no attribute 'preview_file'`

- [ ] **Step 3: Реализовать preview в ядре**

Добавить в `mdtoword/converters.py` в класс `MarkdownToWordConverter` после `convert_file`:

```python
    def preview_content(
        self, content: str, source_path: Path | None = None
    ) -> list[str]:
        """Отрендерить Markdown в память и вернуть варнинги, ничего не сохраняя."""
        try:
            _, warnings = GfmDocxRenderer(
                self.default_font_name, self.default_font_size, self.footnotes_heading
            ).render(content, source_path=source_path)
        except Exception as error:
            raise ConversionError(str(error)) from error
        return warnings

    def preview_file(self, input_path: str | Path) -> list[str]:
        """Прочитать Markdown-файл и отрендерить его вхолостую."""
        source_path = Path(input_path)
        try:
            content = source_path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as error:
            # UnicodeDecodeError — подкласс ValueError, а не OSError:
            # файл в CP1251 иначе улетел бы мимо контракта ConversionError.
            raise ConversionError(str(error)) from error
        return self.preview_content(content, source_path)
```

- [ ] **Step 4: Запустить и убедиться, что проходит**

Run: `python -m pytest tests/test_converters.py -q`
Expected: PASS, 11 passed

- [ ] **Step 5: Написать падающие тесты сервера**

Добавить в `tests/test_mcp_server.py` перед `if __name__ == "__main__":`:

```python
class PreviewTests(McpServerTestCase):
    async def test_preview_reports_warnings_and_writes_no_files(self) -> None:
        (self.root / "doc.md").write_text("![diagram](missing.png)", encoding="utf-8")
        before = sorted(path.name for path in self.root.iterdir())

        result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 1)
        self.assertEqual(
            report["previews"][0]["warnings"], ["Image not found: missing.png"]
        )
        self.assertEqual(sorted(path.name for path in self.root.iterdir()), before)

    async def test_preview_reports_unreadable_files_in_failed(self) -> None:
        (self.root / "good.md").write_text("# Заголовок", encoding="utf-8")
        broken = self.root / "broken.md"
        broken.write_bytes(b"\xff\xfe\x00 invalid utf-8")

        result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["previews"]), 1)
        self.assertEqual(len(report["failed"]), 1)
        self.assertTrue(report["failed"][0]["source"].endswith("broken.md"))


class StdioProtocolTests(unittest.TestCase):
    def test_importing_the_server_writes_nothing_to_stdout(self) -> None:
        # stdout — это канал stdio-протокола: одна лишняя строка при импорте
        # рвёт JSON-RPC сессию, и клиент видит нечитаемую ошибку парсинга.
        import subprocess
        import sys

        repo_root = Path(__file__).resolve().parent.parent
        result = subprocess.run(
            [sys.executable, "-c", "import mdtoword.mcp_server"],
            capture_output=True,
            text=True,
            cwd=repo_root,
        )

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, "")
```

- [ ] **Step 6: Запустить и убедиться, что падает**

Run: `python -m pytest tests/test_mcp_server.py -q`
Expected: FAIL — `Unknown tool: preview_markdown`. Тест `StdioProtocolTests` при этом уже проходит; он остаётся как регрессионная защита.

- [ ] **Step 7: Добавить инструмент preview**

Добавить в `mdtoword/mcp_server.py` модели после `ConversionReport`:

```python
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
```

и сам инструмент после `word_to_markdown`:

```python
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
```

- [ ] **Step 8: Запустить и убедиться, что проходит**

Run: `python -m pytest tests/test_mcp_server.py -q`
Expected: PASS, 13 passed

- [ ] **Step 9: Проверить сервер вживую**

Run: `python -m mdtoword.mcp_server`
Expected: процесс запускается и молча ждёт на stdin, ничего не печатая. Прервать по `Ctrl+C`.

- [ ] **Step 10: Прогнать весь набор**

Run: `python -m pytest tests/ -q`
Expected: PASS

- [ ] **Step 11: Коммит**

```bash
git add mdtoword/converters.py mdtoword/mcp_server.py tests/test_converters.py tests/test_mcp_server.py
git commit -m "feat: preview a markdown conversion without writing the document"
```

---

### Task 5: Документация подключения

**Files:**
- Modify: `README.md`

**Interfaces:**
- Consumes: `main()` и имена трёх инструментов из Задач 3 и 4.
- Produces: ничего.

- [ ] **Step 1: Найти место для раздела**

Run: `grep -n "^## " README.md`
Expected: список разделов. Новый раздел ставится после описания установки и до раздела о сборке — читатель сначала узнаёт, как поставить зависимости, и только потом про интеграцию.

- [ ] **Step 2: Написать раздел**

Добавить в `README.md`:

```markdown
## MCP server

MDtoWORD ships an MCP server so agents can run the same conversions the GUI does.

Install the server dependencies:

```bash
python -m pip install -r requirements-mcp.txt
```

Register it with any MCP client (paths must be absolute):

```json
{
  "mcpServers": {
    "mdtoword": {
      "command": "/path/to/MDtoWord/.venv/bin/python",
      "args": ["-m", "mdtoword.mcp_server"],
      "cwd": "/path/to/MDtoWord"
    }
  }
}
```

For Claude Code:

```bash
claude mcp add mdtoword --scope user \
  -- /path/to/MDtoWord/.venv/bin/python -m mdtoword.mcp_server
```

### Tools

| Tool | What it does |
| --- | --- |
| `markdown_to_word` | Converts `.md` / `.markdown` files and directories to `.docx`, with GFM, footnotes, images and LaTeX → OMML equations. |
| `word_to_markdown` | Converts `.docx` files and directories to Markdown. Lossy: keeps headings, bold, italic and tables; flattens everything else. |
| `preview_markdown` | Renders Markdown in memory and reports only what would not survive the conversion. Writes nothing. |

All three take paths, never file contents, and accept files and directories
mixed together; directories are scanned recursively. Existing output files are
overwritten without warning.
```

- [ ] **Step 3: Проверить, что JSON валиден**

Run: `python -c "import json,re,pathlib; blocks=re.findall(r'\`\`\`json\n(.*?)\`\`\`', pathlib.Path('README.md').read_text(encoding='utf-8'), re.S); [json.loads(b) for b in blocks]; print(f'{len(blocks)} JSON block(s) OK')"`
Expected: `1 JSON block(s) OK` (или больше, если в README уже были json-блоки)

- [ ] **Step 4: Прогнать весь набор в последний раз**

Run: `python -m pytest tests/ -q`
Expected: PASS

- [ ] **Step 5: Коммит**

```bash
git add README.md
git commit -m "docs: document the MCP server and its three tools"
```

---

## Проверка плана против спека

| Требование спека | Задача |
| --- | --- |
| Qt-free `converters.py` | 1 |
| Контракт `list[str]` / `ConversionError` | 1 |
| Локализация переезжает в GUI | 2 |
| `mcp_server.py`, stdio, FastMCP, Pydantic | 3 |
| `markdown_to_word`, `word_to_markdown` | 3 |
| `preview_markdown` | 4 |
| `sources_found` отличает «путь неверный» от «нечего делать» | 3 (тест), 4 (тест) |
| Пустой `inputs` → ошибка | 3 |
| Падение файла не рвёт батч | 3 |
| `output_dir` создаётся с `parents=True` | 3 |
| Молчаливая перезапись | 3 (поведение по умолчанию, задокументировано в 5) |
| Лоссовость `word_to_markdown` в описании инструмента | 3 (реализация + тест) |
| `requirements-mcp.txt` отдельно | 3 |
| Тест чистоты stdout | 4 |
| `test_markdown_converter.py` переезжает | 1 |
| Конфигурация клиента в документации | 5 |
```
