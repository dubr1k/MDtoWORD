# DOCX Formatting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Весь текст в сгенерированном DOCX — чёрный; обычные абзацы выровнены по ширине; таблицы имеют явные границы и уважают выравнивание колонок из Markdown; блоки кода остаются слева и не растягиваются.

**Architecture:** Все правки в `gfm_docx_renderer.py`. Цвет задаётся в `_configure_document` (правка встроенных стилей `Heading N`) и в `_append_hyperlink` (жёсткий синий `0563C1` → чёрный). Выравнивание вводится точечно в `_new_paragraph`, чтобы блоки кода и заголовки его не получили. Таблицы получают прямой `tblBorders` в `_finish_table` плюс выравнивание из токена `table_open`.

**Tech Stack:** Python 3.12, python-docx 1.1.2, markdown-it-py, unittest (offscreen QPA).

## Global Constraints

- Интерпретатор: **только** `/opt/anaconda3/envs/mdtoword/bin/python` (conda env `mdtoword`). Системный `python3` падает — две копии Qt.
- Команда тестов (pytest в env нет, `unittest discover` не работает — нет `tests/__init__.py`):
  ```bash
  cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter
  ```
  Baseline до начала работ: `Ran 26 tests ... OK`.
- Ничего в GUI (`md_to_word_converter.py`, `gui_theme.py`) не трогать — ветка `fix/gui-round-two` уже отревьюена.
- Коммит после каждой задачи, стиль сообщений `fix:` / `feat:` / `chore:`.

## Подтверждённые факты (разбор сгенерированного DOCX, 2026-07-20)

| Дефект | Доказательство |
|---|---|
| Заголовки не чёрные | `Heading 1` → `365F91`, `Heading 2`/`Heading 3` → `4F81BD` в определениях стилей документа |
| Ссылки не чёрные | `_append_hyperlink` жёстко пишет `w:color val="0563C1"` |
| Нет выравнивания по ширине | у всех абзацев `alignment=None` → наследуется LEFT |
| Границы таблиц не гарантированы | таблица ссылается на стиль `TableGrid` (в нём `tblBorders` есть), но собственный `tblPr` таблицы содержит только `tblStyle`/`tblW`/`tblLook` — прямых границ нет, и просмотрщики вроде Pages их не отрисовывают |
| Выравнивание колонок Markdown игнорируется | `_finish_table` жёстко ставит всем ячейкам `WD_ALIGN_PARAGRAPH.LEFT`, хотя в источнике заданы `---:` и `:---:` |
| Перенос строк в коде работает | блок из 14 строк даёт 13 `<w:br/>`; отступы сохранены пробелами; `w:tab` не используется |
| Длинная строка кода | 124 символа в одном run без переносов — Word перенесёт её по ширине в произвольном месте |

## Решения по неоднозначностям

- «Весь шрифт чёрный» трактуется буквально и включает гиперссылки. Подчёркивание у ссылок сохраняется, иначе ссылка перестанет быть отличимой от обычного текста.
- Выравнивание по ширине применяется к обычным абзацам, элементам списков и цитатам. Оно **не** применяется к заголовкам, к подписи языка над блоком кода, к самому блоку кода и к ячейкам таблиц — растянутый по ширине моноширинный код нечитаем, а у ячеек своё выравнивание из Markdown.

---

### Task 1: Весь текст чёрный

**Files:**
- Modify: `gfm_docx_renderer.py` (`_configure_document`, `_append_hyperlink`)
- Test: `tests/test_gfm_docx_renderer.py`

**Interfaces:**
- Produces: `GfmDocxRenderer._BLACK` (модульная или классовая константа `RGBColor(0, 0, 0)`), используемая и в стилях, и в ссылках.

- [ ] **Step 1: Write the failing test**

В `tests/test_gfm_docx_renderer.py` добавить импорт, если его нет:

```python
from docx.shared import RGBColor
```

и тесты:

```python
    def test_headings_are_black(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "# Заголовок\n\n## Второй\n\n### Третий\n"
        )
        for level in (1, 2, 3):
            style = document.styles[f"Heading {level}"]
            self.assertEqual(style.font.color.rgb, RGBColor(0, 0, 0))

    def test_hyperlink_is_black_and_underlined(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "[сайт](https://example.com)\n"
        )
        xml = document.paragraphs[0]._p.xml
        self.assertIn('w:val="000000"', xml)
        self.assertNotIn("0563C1", xml)
        self.assertIn("w:u", xml)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gfm_docx_renderer -v`
Expected: FAIL — заголовки `365F91`/`4F81BD`, ссылка `0563C1`.

- [ ] **Step 3: Write the implementation**

В `gfm_docx_renderer.py` добавить константу рядом с `_TASK_PREFIX`:

```python
_BLACK = RGBColor(0, 0, 0)
```

Заменить `_configure_document`:

```python
    def _configure_document(self) -> None:
        style = cast(ParagraphStyle, self.document.styles["Normal"])
        style.font.name = self.font_name
        style.font.size = self.font_size
        style.font.color.rgb = _BLACK
        for level in range(1, 10):
            try:
                heading = cast(ParagraphStyle, self.document.styles[f"Heading {level}"])
            except KeyError:
                continue
            heading.font.color.rgb = _BLACK
            heading.font.name = self.font_name
```

В `_append_hyperlink` заменить строку с цветом:

```python
        color.set(qn("w:val"), "000000")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: `Ran 28 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add gfm_docx_renderer.py tests/test_gfm_docx_renderer.py
git commit -m "feat: render every DOCX run in black"
```

---

### Task 2: Выравнивание по ширине для текстовых абзацев

**Files:**
- Modify: `gfm_docx_renderer.py` (`_new_paragraph`, `_render_block` для `paragraph_open` в цитате, `_render_code_block`)
- Test: `tests/test_gfm_docx_renderer.py`

**Interfaces:**
- Consumes: `WD_ALIGN_PARAGRAPH` (уже импортирован в модуле).

- [ ] **Step 1: Write the failing test**

```python
    def test_body_paragraphs_are_justified(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "Обычный абзац текста.\n\n- пункт списка\n\n> цитата\n"
        )
        justified = [
            p for p in document.paragraphs
            if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
        ]
        self.assertEqual(len(justified), 3)

    def test_headings_and_code_are_not_justified(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "# Заголовок\n\n```python\nx = 1\n```\n"
        )
        for paragraph in document.paragraphs:
            self.assertNotEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
```

В шапку теста добавить импорт, если его нет:

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gfm_docx_renderer -v`
Expected: FAIL — `test_body_paragraphs_are_justified`, найдено 0 выровненных абзацев.

- [ ] **Step 3: Write the implementation**

Заменить `_new_paragraph` в `gfm_docx_renderer.py`:

```python
    def _new_paragraph(self):
        if self._list_stack:
            style = "List Number" if self._list_stack[-1] == "ordered_list_open" else "List Bullet"
            paragraph = self.document.add_paragraph(style=style)
            paragraph.paragraph_format.left_indent = Pt(18 * (len(self._list_stack) - 1))
        elif self._quote_depth:
            paragraph = self.document.add_paragraph(style="Quote")
        else:
            paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return paragraph
```

Блок кода и подпись языка создаются в `_render_code_block` напрямую через `self.document.add_paragraph()`, минуя `_new_paragraph`, поэтому выравнивание к ним не применяется — менять там ничего не нужно. Заголовки создаются в `_render_block` тоже напрямую.

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: `Ran 30 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add gfm_docx_renderer.py tests/test_gfm_docx_renderer.py
git commit -m "feat: justify body paragraphs, lists and quotes"
```

---

### Task 3: Явные границы таблиц и выравнивание колонок из Markdown

Стиль `Table Grid` описывает границы, но таблица не несёт собственного `tblBorders`, и часть просмотрщиков (Pages, отдельные сборки LibreOffice) стилевые границы не отрисовывает. Прямое форматирование honors везде. Одновременно `_finish_table` затирает выравнивание колонок, заданное в Markdown через `---:` и `:---:`.

**Files:**
- Modify: `gfm_docx_renderer.py` (`_render_block` — ветки `th_open`/`td_open`; `_finish_table`; поле `_table_alignments`)
- Test: `tests/test_gfm_docx_renderer.py`

**Interfaces:**
- Produces: поле `self._table_alignments: list[str | None]`, заполняемое из атрибута `style` токенов `th_open`; статический метод `GfmDocxRenderer._apply_table_borders(table) -> None`.

- [ ] **Step 1: Write the failing tests**

```python
    def test_table_has_explicit_borders(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "| a | b |\n|---|---|\n| 1 | 2 |\n"
        )
        table_xml = document.tables[0]._tbl.tblPr.xml
        self.assertIn("tblBorders", table_xml)
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            self.assertIn(f"w:{edge}", table_xml)

    def test_table_respects_markdown_column_alignment(self):
        document, _ = GfmDocxRenderer("Times New Roman", Pt(12)).render(
            "| left | right | center |\n|:---|---:|:---:|\n| a | b | c |\n"
        )
        body = document.tables[0].rows[1]
        self.assertEqual(body.cells[0].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(body.cells[1].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(body.cells[2].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gfm_docx_renderer -v`
Expected: FAIL — в `tblPr` нет `tblBorders`; все ячейки LEFT.

- [ ] **Step 3: Write the implementation**

В `render()` и в объявлении полей добавить сброс нового состояния рядом с `self._table_rows = None`:

```python
        self._table_alignments = []
```

и в блок аннотаций полей `__init__`:

```python
        self._table_alignments: list[str | None]
```

В `_render_block`, в ветке `table_open`, сбросить выравнивания:

```python
        if token_type == "table_open":
            self._table_rows = []
            self._table_alignments = []
            return
```

Заменить ветку `th_open`/`td_open`, чтобы собирать выравнивание из заголовочной строки. markdown-it выдаёт его в атрибуте `style` вида `text-align:right`:

```python
        if token_type in {"th_open", "td_open"}:
            self._table_cell = []
            if self._table_header:
                style_attr = token.attrGet("style") or ""
                if "right" in style_attr:
                    self._table_alignments.append("right")
                elif "center" in style_attr:
                    self._table_alignments.append("center")
                else:
                    self._table_alignments.append(None)
            return
```

Заменить `_finish_table`:

```python
    def _finish_table(self) -> None:
        if not self._table_rows:
            self._table_rows = None
            return
        columns = max(len(row) for row in self._table_rows)
        table = self.document.add_table(rows=len(self._table_rows), cols=columns)
        table.style = "Table Grid"
        self._apply_table_borders(table)
        alignments = {
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }
        for row_index, values in enumerate(self._table_rows):
            for column_index, value in enumerate(values):
                paragraph = table.cell(row_index, column_index).paragraphs[0]
                run = paragraph.add_run(value)
                if row_index == 0:
                    run.bold = True
                column_alignment = (
                    self._table_alignments[column_index]
                    if column_index < len(self._table_alignments)
                    else None
                )
                paragraph.alignment = alignments.get(
                    column_alignment or "", WD_ALIGN_PARAGRAPH.LEFT
                )
        self._table_rows = None
        self._table_alignments = []

    @staticmethod
    def _apply_table_borders(table: Any) -> None:
        """Write borders as direct formatting so every viewer renders them."""
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element = OxmlElement(f"w:{edge}")
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "4")
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), "000000")
            borders.append(element)
        table._tbl.tblPr.append(borders)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: `Ran 32 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add gfm_docx_renderer.py tests/test_gfm_docx_renderer.py
git commit -m "feat: give tables direct borders and Markdown column alignment"
```

---

### Task 4: Проверка на реальном документе

**Files:**
- Create: `/private/tmp/claude-501/-Users-dubr1k-Syncthing-development-MDtoWord/e6e922e0-736c-417f-9be2-93aa24406758/scratchpad/verify.py` (артефакт проверки, в git не идёт)

- [ ] **Step 1: Full suite**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter
```
Expected: `Ran 32 tests ... OK`.

- [ ] **Step 2: Конвертировать образец и подтвердить все четыре требования**

Конвертировать `/private/tmp/claude-501/-Users-dubr1k-Syncthing-development-MDtoWord/e6e922e0-736c-417f-9be2-93aa24406758/scratchpad/sample.md` и распечатать по документу: цвет каждого стиля `Heading N` и цвет ссылки; выравнивание каждого абзаца с его стилем; наличие `tblBorders` в `tblPr` таблицы и выравнивание ячеек; число `<w:br/>` в блоке кода и длину самой длинной строки кода.

Expected: все `Heading` и ссылка — `000000`; абзацы Normal/List/Quote — `JUSTIFY`, Heading и код — не JUSTIFY; `tblBorders` присутствует; в блоке кода 13 `<w:br/>`.

- [ ] **Step 3: Пересобрать приложение**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && bash scripts/build_macos.sh
```
Expected: `Готово: .../dist/MDtoWORD.app`, codesign `valid on disk`.
