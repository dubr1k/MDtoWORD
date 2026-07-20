# MDtoWord GUI Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Исправить подтверждённые дефекты GUI: окно не влезает в экран (min height 944px), drop на dashed-зону не работает, вкладка «Текст» остаётся пустой кликабельной в режиме Word→MD, «блоб» заголовков QGroupBox, невидимая стрелка QComboBox, двойная вложенность карточек, прыжок layout при показе прогресса.

**Architecture:** Точечные правки двух файлов — `md_to_word_converter.py` (drag&drop на окно, setTabVisible, плоская компоновка вкладки «Файлы», кликабельная дроп-зона, состояния кнопок) и `gui_theme.py` (заголовок QGroupBox внутри карточки, QPalette для Fusion-примитивов, контраст границ). Каждая правка закрывается юнит-тестом в offscreen-режиме, включая пиксельные проверки рендера.

**Tech Stack:** Python 3.12, PyQt6 6.10.2, unittest (offscreen QPA).

## Global Constraints

- Интерпретатор: **только** `/opt/anaconda3/envs/mdtoword/bin/python` (conda env `mdtoword`). Системный `python3` падает из-за двух копий Qt (anaconda base + PyQt6) — не использовать.
- Команда тестов (pytest в env нет, `unittest discover` не работает — нет `tests/__init__.py`):
  ```bash
  cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter
  ```
  Baseline до начала работ: `Ran 14 tests ... OK`.
- Все видимые строки — через `self.translations`, всегда в обеих локалях (`ru` и `en`).
- Тёмная тема — по умолчанию; каждое визуальное изменение проверяется в обеих темах.
- Коммит после каждой задачи. Сообщения — в стиле репозитория (`fix:`, `feat:`, `chore:`), тело можно по-русски.
- `git status` до начала должен быть чистым (кроме этого файла плана).

## Подтверждённые факты (offscreen-проба от 2026-07-20)

| Дефект | Доказательство |
|---|---|
| Окно нельзя сжать ниже 944px по высоте | `minimumSizeHint = 475×944` при `resize(860, 760)` — окно принудительно 944px, не влезает в MacBook-экран |
| Drop на dashed-подсказку не работает | `drop_hint.acceptDrops()=False`, `centralWidget().acceptDrops()=False`; drop принимает только QListWidget |
| Вкладка «Текст» в режиме Word→MD пустая, но кликабельная | после `_toggle_converter_type()`: `tabs.isTabVisible(1)=True`, `text_tab.isHidden()=True`, `setCurrentIndex(1)` проходит |
| «Блоб» заголовка QGroupBox | пиксель в строке заголовка `output_group`: слева `#181d24` (плашка), справа `#0d1117` (фон окна) — заголовок плавает над рамкой |
| Прыжок layout при конвертации | `progress.sizePolicy().retainSizeWhenHidden()=False` |
| Скриншот пользователя — со старой сборки | `dist/MDtoWORD.app` собран 17:31, финальный коммит af72ffd — 18:12 |
| Вертикальный бюджет | settings_group minHint=110, output_group=103 (по ~49px оверхеда QGroupBox), tabs=494 (тройная вложенность: tab margins 9 + group margins 9 + QGroupBox padding 18/15 + margin-top 16) |

---

### Task 1: Drag&drop на всё окно

Сейчас файлы можно бросить только на QListWidget; dashed-зона с надписью «Перетащите файлы сюда» — очевидная цель — отклоняет drop.

**Files:**
- Modify: `md_to_word_converter.py` (класс `DropFileList` ~строка 140, класс `ConverterGUI.__init__` ~строка 182)
- Test: `tests/test_drop_queue.py`

**Interfaces:**
- Produces: модульная функция `_dropped_local_paths(event) -> list[str]`; методы `ConverterGUI.dragEnterEvent/dragMoveEvent/dropEvent`. Позже ни одна задача их не меняет.

- [ ] **Step 1: Write the failing test**

В `tests/test_drop_queue.py`, класс `DropFileListTests`:

```python
    def test_window_accepts_dropped_files(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            md_file = Path(directory) / "input.md"
            md_file.write_text("# hi", encoding="utf-8")

            mime = QMimeData()
            mime.setUrls([QUrl.fromLocalFile(str(md_file))])
            enter = QDragEnterEvent(
                QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            )
            window.dragEnterEvent(enter)
            self.assertTrue(enter.isAccepted())

            drop = QDropEvent(
                QPointF(10, 10), Qt.DropAction.CopyAction, mime,
                Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            )
            window.dropEvent(drop)
            self.assertEqual(window.selected_files, [md_file.resolve()])
```

`md_file.resolve()` обязателен: `discover_sources` канонизирует пути (на macOS `/var/...` → `/private/var/...`).

- [ ] **Step 2: Run test to verify it fails**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: FAIL — `test_window_accepts_dropped_files`, `enter.isAccepted()` False (у QMainWindow нет обработчиков).

- [ ] **Step 3: Write minimal implementation**

В `md_to_word_converter.py` добавить модульную функцию перед классом `DropFileList`:

```python
def _dropped_local_paths(event: Any | None) -> list[str]:
    """Local filesystem paths carried by a drop event, if any."""
    if event is None:
        return []
    mime_data = event.mimeData()
    if mime_data is None:
        return []
    return [url.toLocalFile() for url in mime_data.urls() if url.isLocalFile()]
```

Упростить `DropFileList.dropEvent` через неё:

```python
    def dropEvent(self, event: QDropEvent | None) -> None:
        paths = _dropped_local_paths(event)
        if event is None:
            return
        if paths:
            self.paths_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()
```

В `ConverterGUI.__init__` после `self._create_widgets()` добавить:

```python
        self.setAcceptDrops(True)
```

В класс `ConverterGUI` (рядом с `_set_icon`) добавить:

```python
    def dragEnterEvent(self, event: QDragEnterEvent | None) -> None:
        DropFileList._accept_url_event(event)

    def dragMoveEvent(self, event: QDragMoveEvent | None) -> None:
        DropFileList._accept_url_event(event)

    def dropEvent(self, event: QDropEvent | None) -> None:
        paths = _dropped_local_paths(event)
        if event is None:
            return
        if paths:
            self._add_sources(paths)
            event.acceptProposedAction()
        else:
            event.ignore()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: PASS (все тесты модуля).

- [ ] **Step 5: Run full suite**

Run: полная команда тестов из Global Constraints.
Expected: `Ran 15 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add md_to_word_converter.py tests/test_drop_queue.py
git commit -m "fix: accept drag-and-drop on the whole window"
```

---

### Task 2: Скрывать вкладку «Текст» в режиме Word→MD

`_update_ui` вызывает `self.text_tab.setVisible(is_markdown)` — это прячет **содержимое страницы**, но не саму вкладку: в режиме Word→MD «Текст» остаётся кликабельной и показывает пустую панель.

**Files:**
- Modify: `md_to_word_converter.py` (`_update_ui`, строка `self.text_tab.setVisible(is_markdown)`)
- Test: `tests/test_drop_queue.py`

**Interfaces:**
- Consumes: `ConverterGUI.toggle_button` (существует), `ConverterGUI.tabs`, `ConverterGUI.text_tab`.

- [ ] **Step 1: Write the failing test**

В `tests/test_drop_queue.py`:

```python
    def test_text_tab_hidden_in_word_mode(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            text_index = window.tabs.indexOf(window.text_tab)
            window.tabs.setCurrentIndex(text_index)

            window.toggle_button.click()

            self.assertFalse(window.tabs.isTabVisible(text_index))
            self.assertNotEqual(window.tabs.currentIndex(), text_index)

            window.toggle_button.click()
            self.assertTrue(window.tabs.isTabVisible(text_index))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: FAIL — `isTabVisible` возвращает True после переключения режима.

- [ ] **Step 3: Write minimal implementation**

В `_update_ui` заменить строку:

```python
        self.text_tab.setVisible(is_markdown)
```

на:

```python
        self.tabs.setTabVisible(self.tabs.indexOf(self.text_tab), is_markdown)
```

(Qt сам переключает текущую вкладку, когда активная скрывается.)

- [ ] **Step 4: Run test to verify it passes**

Run: та же команда.
Expected: PASS.

- [ ] **Step 5: Run full suite**

Expected: `Ran 16 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add md_to_word_converter.py tests/test_drop_queue.py
git commit -m "fix: hide the text tab in Word to Markdown mode"
```

---

### Task 3: Заголовки карточек без «блоба», QPalette, стрелка комбобокса

Три дефекта темы: (1) заголовок QGroupBox рисуется над рамкой плашкой цвета surface на фоне окна — виден «блоб» `#181d24` на `#0d1117`; (2) `QComboBox::down-arrow { width: 8px; height: 8px; }` без `image:` убивает нативную стрелку Fusion — комбобокс без индикатора; (3) QPalette приложению не назначается — Fusion-примитивы (стрелки спинбокса, меню) рисуются системными цветами. Плюс границы `#21262D` на `#161B22` почти невидимы, и `QLabel#drop-zone` задан двумя блоками.

**Files:**
- Modify: `gui_theme.py` (палитры, `apply`, QSS)
- Test: `tests/test_gui_theme.py`, `tests/test_drop_queue.py`

**Interfaces:**
- Produces: `ThemeManager._widget_palette(palette: ThemePalette) -> QPalette` (staticmethod); `apply()` дополнительно вызывает `app.setPalette(...)`. Токены `border`/`border_strong` меняют значения — имена прежние.

- [ ] **Step 1: Write the failing tests**

В `tests/test_gui_theme.py` — добавить импорт и тесты:

```python
from PyQt6.QtGui import QPalette
```

```python
    def test_apply_sets_fusion_palette(self):
        manager = ThemeManager(settings=self.settings)
        manager.apply(self.app)
        self.assertEqual(
            self.app.palette().color(QPalette.ColorRole.Window).name().upper(),
            "#0D1117",
        )

    def test_stylesheet_keeps_native_combo_arrow(self):
        stylesheet = ThemeManager.stylesheet("dark")
        self.assertNotIn("down-arrow", stylesheet)
        self.assertNotIn("::drop-down", stylesheet)
```

В `tests/test_drop_queue.py` — добавить импорт `QApplication` уже есть; добавить тест пиксельной проверки:

```python
    def test_card_title_row_has_uniform_background(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            window.show()
            QApplication.processEvents()

            image = window.grab().toImage()
            group = window.output_group
            origin = group.mapTo(window, QPoint(0, 0))
            left = image.pixelColor(origin.x() + 40, origin.y() + 4)
            right = image.pixelColor(
                origin.x() + group.width() - 40, origin.y() + 4
            )

            self.assertEqual(left.name(), right.name())
            self.assertEqual(left.name(), "#161b22")

            window.hide()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gui_theme tests.test_drop_queue -v`
Expected: FAIL — `test_apply_sets_fusion_palette` (палитра системная), `test_stylesheet_keeps_native_combo_arrow` (`down-arrow` присутствует), `test_card_title_row_has_uniform_background` (слева плашка, справа фон окна).

- [ ] **Step 3: Write the implementation**

В `gui_theme.py`:

3a. Импорт вверху:

```python
from PyQt6.QtGui import QColor, QPalette
```

3b. В `_PALETTES["dark"]` заменить границы:

```python
        border="#2A3038",
        border_strong="#3B4552",
```

В `_PALETTES["light"]`:

```python
        border="#D0D7E2",
        border_strong="#AEB9C9",
```

3c. Заменить `apply`:

```python
    def apply(self, app: QApplication) -> None:
        app.setStyle("Fusion")
        app.setPalette(self._widget_palette(_PALETTES[self.theme]))
        app.setStyleSheet(self.stylesheet(self.theme))

    @staticmethod
    def _widget_palette(palette: ThemePalette) -> QPalette:
        qpalette = QPalette()
        color_roles = {
            QPalette.ColorRole.Window: palette.background,
            QPalette.ColorRole.WindowText: palette.text,
            QPalette.ColorRole.Base: palette.input,
            QPalette.ColorRole.AlternateBase: palette.surface,
            QPalette.ColorRole.Text: palette.text,
            QPalette.ColorRole.Button: palette.surface,
            QPalette.ColorRole.ButtonText: palette.text,
            QPalette.ColorRole.BrightText: palette.text,
            QPalette.ColorRole.Highlight: palette.accent,
            QPalette.ColorRole.HighlightedText: "#FFFFFF",
            QPalette.ColorRole.Link: palette.accent,
            QPalette.ColorRole.ToolTipBase: palette.surface_raised,
            QPalette.ColorRole.ToolTipText: palette.text,
            QPalette.ColorRole.PlaceholderText: palette.text_muted,
        }
        for role, value in color_roles.items():
            qpalette.setColor(role, QColor(value))
        disabled = QColor(palette.disabled)
        for role in (
            QPalette.ColorRole.Text,
            QPalette.ColorRole.ButtonText,
            QPalette.ColorRole.WindowText,
        ):
            qpalette.setColor(QPalette.ColorGroup.Disabled, role, disabled)
        return qpalette
```

3d. В QSS заменить оба блока `QGroupBox` / `QGroupBox::title` (заголовок теперь ВНУТРИ карточки, плашка не нужна):

```css
            QGroupBox {{
                background: {palette.surface};
                border: 1px solid {palette.border};
                border-radius: 12px;
                margin-top: 0;
                padding: 30px 14px 12px 14px;
                font-weight: 700;
            }}
            QGroupBox::title {{
                subcontrol-origin: border;
                subcontrol-position: top left;
                left: 14px;
                top: 9px;
                background: transparent;
                padding: 0;
                color: {palette.text};
            }}
```

3e. Удалить целиком блоки `QComboBox::drop-down { ... }` и `QComboBox::down-arrow { ... }` — Fusion сам нарисует стрелку из палитры.

3f. Удалить первый (короткий) блок `QLabel#drop-zone {{ color: ...; font-size: 15px; font-weight: 600; }}` рядом с QLabel-правилами и заменить второй на единый:

```css
            QLabel#drop-zone {{
                background: {palette.surface_raised};
                border: 2px dashed {palette.accent};
                border-radius: 12px;
                color: {palette.accent};
                font-size: 15px;
                font-weight: 600;
                padding: 12px 16px;
            }}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_gui_theme tests.test_drop_queue -v`
Expected: PASS все.

- [ ] **Step 5: Run full suite**

Expected: `Ran 19 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add gui_theme.py tests/test_gui_theme.py tests/test_drop_queue.py
git commit -m "feat: keep card titles inside cards and theme Fusion primitives"
```

---

### Task 4: Компактная компоновка — окно ≤ 760px, плоская вкладка «Файлы», живая дроп-зона

Минимальная высота окна 944px не влезает в экран MacBook (эффективно ~780–900pt). Крупнейшие расходы: QGroupBox вокруг очереди внутри уже оформленной панели вкладки (тройная вложенность, дублирующий заголовок «Файлы Markdown» при вкладке «Файлы»), min-высоты 100+180, отступы 9+9+18/15. Дроп-зона: нет переноса текста, нет реакции на клик.

**Files:**
- Modify: `md_to_word_converter.py` (`DropFileList` — рядом добавить `DropZoneLabel`; `_create_widgets`; `_update_ui`; `translations`)
- Test: `tests/test_drop_queue.py`

**Interfaces:**
- Produces: класс `DropZoneLabel(QLabel)` с сигналом `clicked` (без аргументов); `self.drop_hint` становится `DropZoneLabel`. Атрибут `self.queue_group` и ключи переводов `queue_md`/`queue_word` УДАЛЯЮТСЯ — после этой задачи на них нельзя ссылаться.
- Consumes: `self._select_files` (существует).

- [ ] **Step 1: Write the failing tests**

В `tests/test_drop_queue.py` — добавить импорт:

```python
from PyQt6.QtTest import QTest
```

Изменить существующий тест `test_drop_target_has_room_for_dragging` (порог 100 → 80):

```python
            self.assertGreaterEqual(window.drop_hint.minimumHeight(), 80)
```

Добавить два теста:

```python
    def test_window_fits_small_screens(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            window.show()
            QApplication.processEvents()

            self.assertLessEqual(window.minimumSizeHint().height(), 760)

            window.hide()

    def test_drop_zone_click_requests_files(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            received = []
            window.drop_hint.clicked.connect(lambda: received.append(True))

            QTest.mouseClick(window.drop_hint, Qt.MouseButton.LeftButton)

            self.assertEqual(received, [True])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: FAIL — `test_window_fits_small_screens` (944 > 760), `test_drop_zone_click_requests_files` (AttributeError: нет сигнала `clicked` у QLabel).

- [ ] **Step 3: Write the implementation**

В `md_to_word_converter.py`:

3a. После класса `DropFileList` добавить:

```python
class DropZoneLabel(QLabel):
    """Clickable drop hint that doubles as a file-picker button."""

    clicked = pyqtSignal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("drop-zone")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setWordWrap(True)
        self.setMinimumHeight(80)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mouseReleaseEvent(self, event: Any | None) -> None:
        if event is not None and event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)
```

3b. В `_create_widgets` заменить корневые отступы:

```python
        layout.setContentsMargins(20, 14, 20, 16)
        layout.setSpacing(10)
```

3c. В `_create_widgets` заменить весь блок вкладки «Файлы» — от `files_tab = QWidget()` до `self.tabs.addTab(files_tab, "")` включительно — на плоскую версию без `queue_group`:

```python
        files_tab = QWidget()
        files_tab.setObjectName("tab-page")
        files_layout = QVBoxLayout(files_tab)
        files_layout.setContentsMargins(16, 16, 16, 16)
        files_layout.setSpacing(10)
        self.drop_hint = DropZoneLabel()
        self.drop_hint.clicked.connect(self._select_files)
        files_layout.addWidget(self.drop_hint)
        actions = QHBoxLayout()
        self.add_files_button = QPushButton()
        self.add_files_button.clicked.connect(self._select_files)
        self.add_folder_button = QPushButton()
        self.add_folder_button.clicked.connect(self._select_folder)
        actions.addWidget(self.add_files_button)
        actions.addWidget(self.add_folder_button)
        actions.addStretch()
        files_layout.addLayout(actions)
        self.files_listbox = DropFileList()
        self.files_listbox.paths_dropped.connect(self._add_sources)
        self.files_listbox.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.files_listbox.setMinimumHeight(100)
        files_layout.addWidget(self.files_listbox, 1)
        removal = QHBoxLayout()
        self.remove_button = QPushButton()
        self.remove_button.setObjectName("danger-button")
        self.remove_button.clicked.connect(self._remove_selected_files)
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("danger-button")
        self.clear_button.clicked.connect(self._clear_files)
        removal.addWidget(self.remove_button)
        removal.addWidget(self.clear_button)
        removal.addStretch()
        files_layout.addLayout(removal)
        self.tabs.addTab(files_tab, "")
```

Прежние строки `self.drop_hint.setObjectName("drop-zone")`, `setAlignment(...)`, `setMinimumHeight(100)` не переносить — это делает конструктор `DropZoneLabel`.

3d. В `__init__` заменить `self.resize(860, 760)` на `self.resize(860, 720)`.

3e. В `_update_ui` удалить строку:

```python
        self.queue_group.setTitle(text["queue_md"] if is_markdown else text["queue_word"])
```

3f. В `self.translations` удалить из обеих локалей ключи `queue_md` и `queue_word`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: PASS все, включая `test_window_fits_small_screens`. Если высота всё ещё > 760 — уменьшить `files_listbox.setMinimumHeight` до 90 и padding-top QGroupBox в QSS до 28px, перезапустить.

- [ ] **Step 5: Run full suite**

Expected: `Ran 21 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add md_to_word_converter.py tests/test_drop_queue.py
git commit -m "feat: flatten files tab, clickable drop zone, fit small screens"
```

---

### Task 5: Состояния: прогресс без прыжка, кнопки очереди, счётчик в статусе

Прогресс-бар при показе/скрытии сдвигает layout; «Удалить выбранные»/«Очистить очередь» активны даже при пустой очереди; статус `«5 · Готово к конвертации»` нечитаем.

**Files:**
- Modify: `md_to_word_converter.py` (`_create_widgets`, `_refresh_queue`, `translations`, новый метод `_update_queue_buttons`)
- Test: `tests/test_drop_queue.py`

**Interfaces:**
- Produces: `ConverterGUI._update_queue_buttons() -> None`; ключ переводов `queued` (`"В очереди: {count}"` / `"In queue: {count}"`).

- [ ] **Step 1: Write the failing tests**

В `tests/test_drop_queue.py`:

```python
    def test_progress_reserves_space_while_hidden(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))

            self.assertTrue(window.progress.isHidden())
            self.assertTrue(window.progress.sizePolicy().retainSizeWhenHidden())

    def test_queue_buttons_follow_queue_state(self):
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(
                str(Path(directory) / "theme.ini"),
                QSettings.Format.IniFormat,
            )
            window = ConverterGUI(theme_manager=ThemeManager(settings=settings))
            md_file = Path(directory) / "input.md"
            md_file.write_text("# hi", encoding="utf-8")

            self.assertFalse(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())

            window._add_sources([str(md_file)])
            self.assertTrue(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())
            self.assertIn("В очереди: 1", window.status_label.text())

            window.files_listbox.setCurrentRow(0)
            self.assertTrue(window.remove_button.isEnabled())

            window._clear_files()
            self.assertFalse(window.clear_button.isEnabled())
            self.assertFalse(window.remove_button.isEnabled())
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue -v`
Expected: FAIL — `retainSizeWhenHidden()` False; кнопки всегда enabled; статус без «В очереди».

- [ ] **Step 3: Write the implementation**

3a. В `_create_widgets` после `self.progress = QProgressBar()`:

```python
        policy = self.progress.sizePolicy()
        policy.setRetainSizeWhenHidden(True)
        self.progress.setSizePolicy(policy)
```

3b. Там же, после `self.files_listbox.paths_dropped.connect(self._add_sources)`:

```python
        self.files_listbox.itemSelectionChanged.connect(self._update_queue_buttons)
```

3c. В `translations["ru"]` добавить `"queued": "В очереди: {count}",` и в `translations["en"]` — `"queued": "In queue: {count}",`.

3d. Заменить `_refresh_queue` и добавить `_update_queue_buttons`:

```python
    def _refresh_queue(self) -> None:
        self.files_listbox.clear()
        for source in self.selected_files:
            self.files_listbox.addItem(f"{source.name}\n{source.parent}")
        count = len(self.selected_files)
        if count:
            self.status_label.setText(self._text["queued"].format(count=count))
        else:
            self.status_label.setText(self._text["ready"])
        self.convert_button.setText(
            self._text["convert"] if not count else f"{self._text['convert']} ({count})"
        )
        self._update_queue_buttons()

    def _update_queue_buttons(self) -> None:
        self.clear_button.setEnabled(bool(self.selected_files))
        self.remove_button.setEnabled(bool(self.files_listbox.selectedItems()))
```

- [ ] **Step 4: Run tests to verify they pass**

Run: та же команда.
Expected: PASS все.

- [ ] **Step 5: Run full suite**

Expected: `Ran 23 tests ... OK`.

- [ ] **Step 6: Commit**

```bash
git add md_to_word_converter.py tests/test_drop_queue.py
git commit -m "feat: polish queue button states, status copy, progress placement"
```

---

### Task 6: Финальная верификация и пересборка приложения

Скриншот пользователя сделан со сборки 17:31, которая старше финального кода (18:12) — без пересборки `dist/MDtoWORD.app` пользователь продолжит видеть старый GUI.

**Files:**
- Create: `/tmp/mdtoword-preview-dark.png`, `/tmp/mdtoword-preview-light.png` (артефакты проверки)
- Modify: `dist/MDtoWORD.app` (пересборка; в git не попадает)

- [ ] **Step 1: Full suite + компиляция**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord
QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python -m unittest tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow tests.test_gfm_docx_renderer tests.test_markdown_converter
/opt/anaconda3/envs/mdtoword/bin/python -m py_compile md_to_word_converter.py gui_theme.py conversion_workflow.py gfm_docx_renderer.py
```

Expected: `Ran 23 tests ... OK`, компиляция без вывода.

- [ ] **Step 2: Скриншоты обеих тем для глазной проверки**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && QT_QPA_PLATFORM=offscreen /opt/anaconda3/envs/mdtoword/bin/python - <<'EOF'
import os
from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication
from gui_theme import ThemeManager
from md_to_word_converter import ConverterGUI

app = QApplication([])
settings = QSettings("/tmp/mdtoword-preview.ini", QSettings.Format.IniFormat)
gui = ConverterGUI(theme_manager=ThemeManager(settings=settings))
gui.show(); app.processEvents()
gui.grab().save("/tmp/mdtoword-preview-dark.png")
gui.theme_button.click(); app.processEvents()
gui.grab().save("/tmp/mdtoword-preview-light.png")
print("saved")
EOF
open /tmp/mdtoword-preview-dark.png /tmp/mdtoword-preview-light.png
```

Expected: `saved`; на скриншотах — заголовки внутри карточек без плашек, видимые границы, окно без обрезки. Показать пользователю.

- [ ] **Step 3: Пересборка macOS-бандла**

```bash
cd /Users/dubr1k/Syncthing/development/MDtoWord && bash scripts/build_macos.sh
```

Expected: завершение с `Готово: /Users/dubr1k/Syncthing/development/MDtoWord/dist/MDtoWORD.app`, codesign `valid on disk`.

- [ ] **Step 4: Ручной smoke-тест собранного приложения**

```bash
open dist/MDtoWORD.app
```

Чек-лист (руками, 2 минуты):
1. Бросить .md-файл на dashed-зону (не на список) — файл попадает в очередь.
2. Клик по dashed-зоне — открывается диалог выбора файлов.
3. Переключить «Режим: MD → Word» — вкладка «Текст» исчезает, возврат — появляется.
4. Стрелка у комбобокса шрифта видна; выпадающий список тёмный.
5. Переключить тему на светлую и обратно — без системно-серых участков.
6. Сжать окно до минимума — всё влезает, ничего не обрезано.

- [ ] **Step 5: Commit плана**

```bash
git add docs/superpowers/plans/2026-07-20-mdtoword-gui-fixes.md
git commit -m "docs: plan GUI fixes round two"
```

---

## Отложено сознательно (не входит в план)

- **Конвертация в GUI-потоке** (`QApplication.processEvents()` в `_convert_files`): на больших пакетах интерфейс замирает. Правильное решение — QThread-воркер с сигналами прогресса и блокировкой кнопок. Отдельная задача с собственными тестами; текущий план не трогает конвейер конвертации.
- **Пустое состояние списка очереди** (подсказка внутри пустого QListWidget) — косметика, не мешает работе.
- **Светлая тема: тонкая подстройка** оттенков после ручного осмотра скриншотов из Task 6.
