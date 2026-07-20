# MDtoWord Theme Refinement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Deliver a persisted dark/light GUI theme that styles every visible Qt surface and removes the gray tab-pane artifact.

**Architecture:** Add a small `ThemeManager` that maps a theme identifier to a complete tokenized QSS stylesheet and persists the chosen theme through `QSettings`. `ConverterGUI` creates a footer toggle and delegates startup and runtime application to the manager. Keep conversion logic unchanged.

**Tech Stack:** Python 3.12, PyQt6, `unittest`, macOS app bundle through PyInstaller.

## Global Constraints

- Dark mode is the initial theme.
- The user can switch themes from the application footer.
- Theme preference persists between launches through `QSettings`.
- Style all visible Qt surfaces explicitly, including the `QTabWidget::pane` artifact source.
- Preserve current conversion and queue behavior.
- Leave changes uncommitted for user review.

---

### Task 1: Add testable theme manager

**Files:**
- Create: `gui_theme.py`
- Create: `tests/test_gui_theme.py`

**Interfaces:**
- Produces: `ThemeManager(settings: QSettings | None = None)` with `theme: str`, `toggle() -> str`, `apply(app: QApplication) -> None`, and `stylesheet(theme: str) -> str`.
- Consumes: `QSettings` and `QApplication` from PyQt6.

- [ ] **Step 1: Write failing tests**

```python
class ThemeManagerTests(unittest.TestCase):
    def test_defaults_to_dark_and_styles_tab_pane(self):
        theme = ThemeManager(settings=QSettings("MDtoWordTests", "Default"))
        self.assertEqual(theme.theme, "dark")
        stylesheet = theme.stylesheet("dark")
        self.assertIn("QTabWidget::pane", stylesheet)
        self.assertIn("QComboBox QAbstractItemView", stylesheet)

    def test_toggle_applies_and_persists_theme(self):
        settings = QSettings("MDtoWordTests", "Toggle")
        settings.clear()
        manager = ThemeManager(settings=settings)
        self.assertEqual(manager.toggle(), "light")
        self.assertEqual(ThemeManager(settings=settings).theme, "light")
```

- [ ] **Step 2: Run the focused test and observe import failure**

Run: `QT_QPA_PLATFORM=offscreen .venv-macos-build/bin/python -m unittest tests.test_gui_theme -v`

Expected: FAIL because `gui_theme` does not exist.

- [ ] **Step 3: Implement `ThemeManager`**

```python
class ThemeManager:
    def __init__(self, settings: QSettings | None = None) -> None:
        self._settings = settings or QSettings("dubr1k", "MDtoWord")
        self.theme = str(self._settings.value("theme", "dark"))

    def toggle(self) -> str:
        self.theme = "light" if self.theme == "dark" else "dark"
        self._settings.setValue("theme", self.theme)
        return self.theme

    def apply(self, app: QApplication) -> None:
        app.setStyle("Fusion")
        app.setStyleSheet(self.stylesheet(self.theme))
```

Implement theme-token dictionaries and build QSS containing complete selectors for tab panes, tab bars, fields, queue, popup views, scrollbars, buttons, focus, selection, disabled controls, and progress bars.

- [ ] **Step 4: Run focused tests**

Run: `QT_QPA_PLATFORM=offscreen .venv-macos-build/bin/python -m unittest tests.test_gui_theme -v`

Expected: PASS.

### Task 2: Integrate the footer theme control

**Files:**
- Modify: `md_to_word_converter.py`
- Modify: `tests/test_drop_queue.py`

**Interfaces:**
- Consumes: `ThemeManager` from `gui_theme.py`.
- Produces: `ConverterGUI.theme_manager`, `ConverterGUI.theme_button`, and `ConverterGUI._toggle_theme() -> None`.

- [ ] **Step 1: Write failing GUI interaction test**

```python
def test_theme_button_switches_application_theme(qapp):
    window = ConverterGUI()
    self.assertEqual(window.theme_manager.theme, "dark")
    window.theme_button.click()
    self.assertEqual(window.theme_manager.theme, "light")
    self.assertIn("Light", window.theme_button.toolTip())
```

- [ ] **Step 2: Run the focused test and observe missing UI members**

Run: `QT_QPA_PLATFORM=offscreen .venv-macos-build/bin/python -m unittest tests.test_drop_queue.DropFileListTests.test_theme_button_switches_application_theme -v`

Expected: FAIL because the GUI has no theme manager or button.

- [ ] **Step 3: Integrate runtime control**

```python
self.theme_manager = ThemeManager()
self.theme_manager.apply(QApplication.instance())
self.theme_button = QPushButton()
self.theme_button.clicked.connect(self._toggle_theme)


def _toggle_theme(self) -> None:
    self.theme_manager.toggle()
    self.theme_manager.apply(QApplication.instance())
    self._update_theme_button()
```

Place the button in the footer next to language and mode controls. `_update_theme_button()` sets a readable sun/moon label, accessible name, and tooltip for the current action.

- [ ] **Step 4: Run focused GUI test**

Run: `QT_QPA_PLATFORM=offscreen .venv-macos-build/bin/python -m unittest tests.test_drop_queue -v`

Expected: PASS.

### Task 3: Verify production behavior and package

**Files:**
- Modify: `docs/superpowers/specs/2026-07-20-mdtoword-theme-design.md` only if testing reveals a design correction.

- [ ] **Step 1: Run all tests and compile sources**

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv-macos-build/bin/python -m unittest discover -s tests -v
.venv-macos-build/bin/python -m compileall -q md_to_word_converter.py gui_theme.py conversion_workflow.py gfm_docx_renderer.py
```

Expected: all tests pass and compilation has no output.

- [ ] **Step 2: Build and validate the distributable**

Run:

```bash
./scripts/build_macos.sh
codesign --verify --deep --strict --verbose=2 dist/MDtoWORD.app
```

Expected: `dist/MDtoWORD.app: valid on disk` and a satisfied designated requirement.

- [ ] **Step 3: Smoke-test the bundle**

Launch `dist/MDtoWORD.app/Contents/MacOS/MDtoWORD` and confirm it stays running before stopping it.

- [ ] **Step 4: Leave changes uncommitted**

Do not create a commit or push.
