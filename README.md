<a name="english"></a>

# 📄 MDtoWORD — Markdown to Word converter

<div align="center">

**🇺🇸 English** · [🇷🇺 Русский](#russian)

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-PyQt6-orange?style=for-the-badge)
![LaTeX](https://img.shields.io/badge/LaTeX-native%20OMML-red?style=for-the-badge)

**A desktop app that turns GitHub Flavored Markdown into a clean Word document — formulas included, and still editable once they get there**

[🚀 Quick Start](#-quick-start) • [🧭 Interface](#-the-interface) • [📝 Markdown](#-what-markdown-is-understood) • [🧮 Formulas](#-latex-formulas) • [🔧 Troubleshooting](#-troubleshooting) • [🇷🇺 Русский](#russian)

</div>

---

## 🎯 What it does

MDtoWORD takes your `.md` files — one, a dozen, or a whole folder — and drops finished `.docx` files beside them. The markup is not approximated: headings become Word styles, tables get real borders, footnotes are collected into their own section, and a formula like `$E = mc^2$` becomes a **native, editable Word equation** rather than an image or a line of plain text.

It also works the other way round: a `.docx` can be turned back into simplified Markdown.

---

## 🚀 Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/dubr1k/MDtoWORD.git
cd MDtoWORD

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the program
python -m mdtoword
```

Then drop files or a folder onto the window and press **"Convert"**. By default the results land next to their sources.

### Linux (quick launch)
From the project root: `./scripts/launch_mdtoword.sh` — the script changes to the project directory and runs the app (conda env `mdtoword` or system `python3`).

---

## 💻 Installation

### Requirements

- **Python 3.10+** — the code uses `X | Y` annotation syntax, so older versions will not run it
- **python-docx 1.1.2** — writes the `.docx`
- **PyQt6 6.10.2** — the graphical interface
- **Pillow 12.0.0** — images and icons
- **markdown-it-py 4.0.0** — Markdown parsing
- **mdit-py-plugins 0.5.0** — footnotes, `$…$` math and amsmath environments
- **linkify-it-py 2.0.3** — automatic link detection in running text

```bash
pip install -r requirements.txt
```

### Conda (virtual environment)

```bash
# Create the environment from environment.yml (Python 3.11)
conda env create -f environment.yml
conda activate mdtoword

python -m mdtoword
```

For Fish, make sure `conda init fish` has been run (conda on PATH).

> `environment.yml` carries the same package set as `requirements.txt`, so the environment is ready to use as soon as it is created. `tests/test_packaging.py` asserts the two files agree, so they cannot drift apart unnoticed.

---

## 🧭 The interface

**Drag and drop.** Drop files **or whole folders** anywhere on the window — folders are scanned recursively and only matching files join the queue. Clicking the dashed drop zone at the top of the "Files" tab does the same thing.

**The queue.** **"Add files"** and **"Add folder"** open the usual pickers. **"Remove selected"** drops entries you don't want (you can select several at once), and **"Clear queue"** empties the list. Duplicates are never added twice.

**Save location.** By default each result is written next to its own source file. **"Choose folder"** in the "Save location" card sends everything to one directory instead; if a batch contains files with the same name, later ones get a numeric suffix — `report.docx`, `report (2).docx`. **"Reset"** restores the default behaviour.

**Two modes.** The switch in the footer flips the direction: **"Mode: MD → Word"** and **"Mode: Word → MD"**. Switching re-filters the queue against the new set of extensions.

**The "Text" tab.** In Markdown → Word mode there is a "Text" tab next to "Files": paste markup straight from the clipboard, press Convert, and choose where to save the `.docx`. No file needed.

**Appearance.** A font dropdown (Arial, Times New Roman, Calibri, Georgia, Helvetica, Courier New) and a size field from 6 to 72 pt set the document's base formatting.

**Themes and language.** The footer holds a round theme button (☀ / ☾) that switches between dark and light. The choice is stored via `QSettings` and restored on the next launch. The neighbouring **EN / RU** button switches the interface language.

While a batch runs you get a progress bar and the name of the current file; at the end, a dialog with the number of successes, errors and warnings.

---

## 📝 What Markdown is understood

The markup goes through a GitHub Flavored Markdown parser, so support goes well beyond headings and bold text.

| Element | Markdown | Result in Word |
|---|---|---|
| **Headings** | `# … ######` | Heading 1–6 styles with a size hierarchy |
| **Paragraphs** | plain text | Normal style, justified |
| **Italic / bold** | `*text*`, `**text**` | Italic / Bold |
| **Strikethrough** | `~~text~~` | Strikethrough |
| **Inline code** | `` `code` `` | Courier New, 10 pt |
| **Links** | `[text](url)` | A real hyperlink, black and underlined |
| **Images** | `![alt](path)` | Embedded picture; local paths resolve against the source file, `http(s)` URLs are downloaded |
| **Blockquotes** | `> text` | Quote style, justified |
| **Thematic breaks** | `---` | Horizontal rule |
| **Code blocks** | ````` ```python ````` | Courier New 10 pt, language as an italic caption above |
| **Lists** | `-`, `1.` | Bulleted / Numbered, nesting indented 18 pt per level |
| **Task lists** | `- [ ]`, `- [x]` | ☐ and ☒ at the start of the line |
| **Tables** | `\| … \|` | Bordered table, bold header, column alignment from the markup |
| **Footnotes** | `[^1]` and `[^1]: …` | A `[1]` reference in place and a "Footnotes" section at the end |
| **Formulas** | `$…$`, `$$…$$` | Native Word equations — see [below](#-latex-formulas) |

On top of that: bare URLs in the text are linkified automatically, and a single newline inside a paragraph is preserved as a line break. Raw HTML inside Markdown is not processed and does not reach the document.

---

## 🧮 LaTeX formulas

The headline feature of this release. A formula can be written four ways:

```markdown
Inline: $E = mc^2$

As a display block:
$$\int_0^1 x^2\,dx = \frac{1}{3}$$

With an equation number:
$$a^2 + b^2 = c^2$$ (1)

In an amsmath environment:
\begin{align}
  f(x) &= ax + b \\
  g(x) &= cx + d
\end{align}
```

All four become **real Word equations (OMML)** — they open in the equation editor, they can be edited, and they scale with the text. Not a picture, not a text imitation.

The supported amsmath environments are `equation`, `multline`, `gather`, `align`, `alignat`, `flalign` and `eqnarray`. A multi-line environment stays **one** Word equation: `gather` and `multline` with their lines stacked, `align` and its relatives stacked *and* aligned on the `&` column, so the `=` signs line up the way LaTeX draws them. The alignment is written with OMML's own `<m:aln/>` marker — the one Word itself uses.

### What's supported

| Construct | Examples |
|---|---|
| Fractions | `\frac`, `\dfrac`, `\tfrac` |
| Roots | `\sqrt{x}`, `\sqrt[3]{x}` |
| Sub- and superscripts | `x^2`, `a_i`, `x_i^2` |
| Greek letters | `\alpha`, `\beta`, … `\omega`, `\Gamma`, `\Delta`, … `\Omega` |
| Operators and relations | `\times`, `\div`, `\pm`, `\cdot`, `\leq`, `\geq`, `\neq`, `\approx`, `\equiv`, `\in`, `\subset`, `\cup`, `\cap`, `\forall`, `\exists`, `\to`, `\Rightarrow`, `\infty`, `\partial`, `\nabla` and more |
| Function names | `\sin`, `\cos`, `\tan`, `\log`, `\ln`, `\exp`, `\det`, `\max`, `\min`, `\sup`, `\inf` — set upright |
| Text inside a formula | `\text{}`, `\mathrm{}`, `\operatorname{}` — Cyrillic included |
| Styles | `\mathbf`, `\mathit`, `\boldsymbol`, `\bm` |
| N-ary operators | `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\oint`, `\bigcup`, `\bigcap`, `\bigoplus`, `\bigotimes`, `\bigvee`, `\bigwedge` — with `_` and `^` limits |
| Limits | `\lim`, `\limsup`, `\liminf` with a lower limit: `\lim_{x \to 0}` |
| Stretchy delimiters | `\left( … \right)`, `\left\{ … \right\}`, `\left\| … \right\|`, `\langle`, `\lfloor`, `\lceil`, plus `\left.` / `\right.` |
| Accents | `\hat`, `\widehat`, `\tilde`, `\widetilde`, `\bar`, `\vec`, `\dot`, `\ddot`, `\acute`, `\grave`, `\check`, `\breve` |
| Overline and underline | `\overline`, `\underline` |
| Binomial coefficient | `\binom{n}{k}`, and the infix `{n \choose k}` |
| Infix fractions | `{a \over b}`, `{n \atop k}` — each splits the group it sits in |
| Matrices | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases` |
| Formula tables | `\begin{array}{lcr} … \end{array}` — the `l`, `c`, `r` column alignment carries into Word |
| Stacked limits | `\substack{i < j \\ i \in S}` |
| Line break | `\\` in any formula, not just inside a matrix or amsmath: the lines stack |
| Line alignment | `&` between the lines of a multi-line formula: `a &= b \\ c &= d` puts the `=` signs under one another |
| Spacing and escapes | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_` |

### Limits of support

The four cases below are refused **deliberately**. This is not a to-do list: each one has a reason why no correct behaviour exists, so the converter refuses honestly instead of producing something that looks close but is wrong.

| Case | Why it is refused | What to do |
|---|---|---|
| `\begin{array}{c\|c}` — a vertical rule | A Word matrix has no rule between columns, and dropping it silently is not an option: an augmented matrix would become an ordinary one. Same for `p{5cm}`, `@{…}`, `\hline` | Split into two matrices, or do without the rule |
| `\matrix{…}`, `\cases{…}` — the plain-TeX spelling | A **different** construct with different grouping rules, not a synonym for the environment | Write `\begin{matrix} … \end{matrix}` |
| `a \over b \over c` — two infix commands in a row | Ambiguous: there is no telling what divides what. TeX itself rejects it too | Brace the halves: `{{a \over b} \over c}` |
| `$a & b$` — a lone `&` outside a multi-line formula | There is nothing to align against, and `Tom & Jerry` inside `$…$` is far more likely a missing escape than a formula | Write `\&` |

That last case is a guard against a typo rather than a limitation: between the lines of a multi-line formula `&` works and aligns (`a &= b \\ c &= d`), as the table above shows.

**Nothing is lost silently.** When a construct isn't supported, the formula goes into the document verbatim, character for character, in a monospace font — and the result dialog carries a warning naming exactly what failed:

```
Formula kept as text: "\begin{array}{c|c} a & b \end{array}"
(Column specification is not supported in \begin{array}: '|'
 (only 'l', 'c' and 'r' columns are))
```

### 💲 A literal dollar sign in prose

This is neither a bug nor a limitation but a convention — the same one Jupyter, Pandoc and MyST use. Since `$` opens a formula, a literal dollar in running text is written `\$`.

So that you don't have to escape everything, the converter works out the common cases by itself:

| You wrote | What you get | Warning |
|---|---|---|
| `It costs $5 and $10` | The text as written — a digit straight after a dollar does not open a formula | none |
| `A price of \$100` | `A price of $100` | none |
| `$E = mc^2$` | A real Word equation | none |
| `$\text{path}$` | A real Word equation — a word inside `\text{}` is legitimate | none |
| `Set $PATH and $HOME` | The text as written, nothing lost | yes — reads as prose, not as a formula |
| `Переменные $HOME и $PATH` | The text as written | yes — Cyrillic outside `\text{}` |

The last two rows are the only ones that reach the result dialog, and the text is kept character for character either way:

```
Inline math "$PATH and $" contains no mathematical symbols and may be
ordinary prose rather than a formula; write a literal "$" as "\$".
```

---

## 🎨 How the Word document is formatted

**Colour.** Every run is black, headings included. The default python-docx template colours headings through the document theme, so the theme colours are cleared explicitly.

**Headings.** Sizes are scaled from the chosen body size. At 12 pt that gives 18 / 16 / 14 / 13 / 12 / 12 pt for levels 1–6; every level is bold, and level 6 is additionally italic so it stays distinct from level 5. Headings use the chosen font, which required stripping the *theme fonts* from the styles — in OOXML a theme attribute overrides an explicitly set font name.

**Alignment.** Paragraphs, list items, quotes and footnotes are justified. Headings, code blocks and table cells are not. Display formulas are centred.

**Tables.** Borders are written as direct formatting rather than left to the style, because some viewers ignore style-level borders and render the table without a grid.

---

## 🏗️ Project structure

```
MDtoWORD/
├── 📦 mdtoword/                  # Application package (run: python -m mdtoword)
│   ├── __init__.py
│   ├── __main__.py               # Entry point
│   ├── app.py                    # PyQt6 GUI
│   ├── converters.py             # Qt-free conversion core, used by the GUI and the MCP server
│   ├── gfm_renderer.py           # Renders GFM markup into a Word document
│   ├── latex_omml.py             # Parses LaTeX and builds OMML equations
│   ├── mcp_server.py             # MCP server: three conversion tools over stdio
│   ├── workflow.py               # Source discovery and output path allocation
│   └── theme.py                  # Dark and light themes, persisted choice
├── 📁 tests/                     # Test suite (unittest)
│   ├── test_drop_queue.py
│   ├── test_gui_theme.py
│   ├── test_conversion_workflow.py
│   ├── test_gfm_docx_renderer.py
│   ├── test_converters.py
│   ├── test_latex_omml.py
│   ├── test_mcp_server.py
│   └── test_packaging.py
├── 📁 scripts/
│   ├── build_macos.sh            # Builds MDtoWORD.app (Apple Silicon)
│   ├── build_windows.ps1         # Builds the Windows bundle and archive
│   └── launch_mdtoword.sh        # Linux (bash) launcher
├── 📁 packaging/
│   ├── MDtoWORD.desktop          # Desktop entry (Linux)
│   └── windows_version_info.txt  # Version metadata for the Windows build
├── 📁 assets/                    # Application icons (png, icns, ico)
├── 📁 docs/
│   ├── DESCRIPTION.md            # Repository description (RU/EN)
│   ├── ИНСТРУКЦИЯ.txt            # Quick guide (RU)
│   ├── INSTRUCTION_EN.txt        # Quick guide (EN)
│   ├── 📁 design/                # Development plans and specifications
│   └── 📁 releases/              # Release notes (1.0, 1.1, 1.1.1)
├── 📄 MDtoWORD.spec              # PyInstaller configuration (macOS)
├── 📋 requirements.txt           # Application dependencies
├── 📋 requirements-build.txt     # Build dependencies (PyInstaller)
├── 📋 requirements-mcp.txt       # MCP server dependencies
├── 📋 environment.yml            # Conda environment (Python 3.11)
└── 📖 README.md                  # Documentation (this file)
```

---

## ⚙️ Technical details

### Default formatting
- **Font**: Times New Roman, 12 pt
- **Colour**: black (RGB 0, 0, 0)
- **Code**: Courier New, 10 pt
- **Tables**: borders as direct formatting, bold header row

### Supported formats
- **Markdown → Word**: input `.md`, `.markdown` → output `.docx`
- **Word → Markdown**: input `.docx` → output `.md`

---

## 🤖 MCP server

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
mixed together; directories are scanned recursively. Where the two converting
tools write, they overwrite an existing output file without warning.

---

## 🛠️ Development

Run the tests from the project root:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest \
    tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow \
    tests.test_gfm_docx_renderer tests.test_converters tests.test_latex_omml \
    tests.test_mcp_server tests.test_packaging
```

The suite currently holds **195 tests**. `QT_QPA_PLATFORM=offscreen` lets the interface tests run without a display.

Standalone bundles:

- `./scripts/build_macos.sh` — builds `dist/MDtoWORD.app` for Apple Silicon: creates a dedicated virtualenv, installs the dependencies, runs PyInstaller against `MDtoWORD.spec` and ad-hoc signs the result;
- `scripts/build_windows.ps1` — builds the Windows bundle, packs it into `dist/MDtoWORD-Windows-x64.zip` and computes the SHA-256.

---

## 🔧 Troubleshooting

**The program doesn't start**
```bash
python --version            # 3.10 or newer required
pip install -r requirements.txt
```

**Encoding error**
Make sure your `.md` files are saved as UTF-8.

**A table lost its formatting**
Check the syntax: a table must have a `|---|---|` separator row. Column alignment is read from that same row — `:---`, `:---:`, `---:`.

**A formula didn't convert**
Check the warning in the result dialog: it names the exact construct, e.g. `LaTeX environment is not supported: \begin{array}`. The formula itself is preserved verbatim in the document — rewrite it using a supported construct from [the table above](#whats-supported) and convert again.

**A dollar sign in the text came out oddly**
Write a literal dollar as `\$`. If the converter sees `$…$` wrapped around ordinary words it leaves the text alone and warns you — but escaping it up front is better.

**A formula inside a table cell stayed as text**
That is deliberate: formulas in table cells are not converted and are kept with their dollars. A warning says so. Move the formula out of the table if it needs to be an equation.

**A display formula inside a list lost its indent**
Display formulas always become their own centred paragraph, so list numbering and indentation do not carry over to them.

**An image didn't make it into the document**
Local paths resolve relative to the `.md` file, and `http(s)` URLs are downloaded with a 10-second timeout. If the file is missing or the network is unavailable, `[alt text]` appears in its place and the dialog carries a warning with the address.

**A large batch makes the window sluggish**
Conversion runs on the interface thread, so long queues leave the window slow to respond. The progress bar still updates — let it finish.

**Word → Markdown doesn't return everything**
The reverse direction is deliberately simplified: it reads headings from styles plus bold and italic runs, and tables are appended at the end of the file rather than in their original position.

---

<a name="russian"></a>

# 📄 MDtoWORD — конвертер Markdown в Word

<div align="center">

[🇺🇸 English](#english) · **🇷🇺 Русский**

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-PyQt6-orange?style=for-the-badge)
![LaTeX](https://img.shields.io/badge/LaTeX-native%20OMML-red?style=for-the-badge)

**Настольное приложение, которое превращает GitHub Flavored Markdown в аккуратный документ Word — вместе с формулами, которые в Word можно редактировать**

[🚀 Быстрый старт](#-быстрый-старт) • [🧭 Интерфейс](#-интерфейс) • [📝 Markdown](#-что-понимается-в-markdown) • [🧮 Формулы](#-формулы-latex) • [🔧 Решение проблем](#-решение-проблем) • [🇺🇸 English](#english)

</div>

---

## 🎯 Что умеет

MDtoWORD берёт ваши `.md`-файлы — один, десяток или целую папку — и складывает рядом готовые `.docx`. Разметка не «приблизительно похожа», а переносится по-настоящему: заголовки становятся стилями Word, таблицы получают настоящие границы, сноски собираются в отдельный раздел, а формула вроде `$E = mc^2$` превращается в **родное редактируемое уравнение Word**, а не в картинку и не в голый текст.

Работает и в обратную сторону: из `.docx` можно вытащить упрощённый Markdown.

---

## 🚀 Быстрый старт

```bash
# 1. Клонируйте репозиторий
git clone https://github.com/dubr1k/MDtoWORD.git
cd MDtoWORD

# 2. Установите зависимости
pip install -r requirements.txt

# 3. Запустите программу
python -m mdtoword
```

Дальше перетащите файлы или папку в окно и нажмите **«Конвертировать»**. Результат по умолчанию появится рядом с исходниками.

### Linux (быстрый запуск)
Из корня проекта: `./scripts/launch_mdtoword.sh` — скрипт сам перейдёт в каталог проекта и запустит приложение (conda-окружение `mdtoword` или системный `python3`).

---

## 💻 Установка

### Требования

- **Python 3.10+** — код использует синтаксис аннотаций `X | Y`, поэтому более старые версии не подойдут
- **python-docx 1.1.2** — запись `.docx`
- **PyQt6 6.10.2** — графический интерфейс
- **Pillow 12.0.0** — работа с изображениями и иконками
- **markdown-it-py 4.0.0** — разбор Markdown
- **mdit-py-plugins 0.5.0** — сноски, `$…$` и окружения amsmath
- **linkify-it-py 2.0.3** — автоматическое распознавание ссылок в тексте

```bash
pip install -r requirements.txt
```

### Conda (виртуальное окружение)

```bash
# Создать окружение из environment.yml (Python 3.11)
conda env create -f environment.yml
conda activate mdtoword

python -m mdtoword
```

Для Fish убедитесь, что выполнен `conda init fish` (conda в PATH).

> `environment.yml` содержит тот же набор пакетов, что и `requirements.txt` — окружение готово к работе сразу после создания. Совпадение двух файлов проверяется тестом `tests/test_packaging.py`, поэтому они не разойдутся незаметно.

---

## 🧭 Интерфейс

**Перетаскивание.** Бросьте файлы **или целые папки** в любое место окна — папки просматриваются рекурсивно, и в очередь попадают только подходящие файлы. То же самое делает клик по пунктирной зоне вверху вкладки «Файлы».

**Очередь.** Кнопки **«Добавить файлы»** и **«Добавить папку»** открывают обычные диалоги выбора. Ненужное убирается кнопкой **«Удалить выбранные»** (выделять можно сразу несколько строк), а **«Очистить очередь»** сбрасывает список целиком. Дубликаты повторно не добавляются.

**Место сохранения.** По умолчанию каждый результат кладётся рядом со своим исходником. Кнопка **«Выбрать папку»** в карточке «Место сохранения» переключает вывод в одну общую директорию; если в пачке встретятся файлы с одинаковыми именами, к повторам добавится номер — `отчёт.docx`, `отчёт (2).docx`. Кнопка **«Сбросить»** возвращает поведение по умолчанию.

**Два режима.** Переключатель в нижней строке меняет направление: **«Режим: MD → Word»** и **«Режим: Word → MD»**. Очередь при переключении фильтруется по новому набору расширений.

**Вкладка «Текст».** В режиме Markdown → Word рядом с вкладкой «Файлы» есть вкладка «Текст»: вставьте туда разметку прямо из буфера обмена, нажмите «Конвертировать» и укажите, куда сохранить `.docx`. Файл при этом не нужен.

**Оформление.** Выпадающий список шрифтов (Arial, Times New Roman, Calibri, Georgia, Helvetica, Courier New) и поле размера от 6 до 72 pt задают базовое оформление документа.

**Темы и язык.** В нижней строке — круглая кнопка темы (☀ / ☾), переключающая тёмное и светлое оформление. Выбор запоминается через `QSettings` и восстанавливается при следующем запуске. Соседняя кнопка **EN / RU** переключает язык интерфейса.

По ходу конвертации показывается прогресс-бар и имя текущего файла, а в конце — диалог с числом успешных файлов, ошибками и предупреждениями.

---

## 📝 Что понимается в Markdown

Разметка разбирается парсером GitHub Flavored Markdown, поэтому поддержка не ограничивается заголовками и жирным текстом.

| Элемент | Markdown | Результат в Word |
|---|---|---|
| **Заголовки** | `# … ######` | Стили Heading 1–6 с иерархией размеров |
| **Абзацы** | обычный текст | Стиль Normal, выравнивание по ширине |
| **Курсив / жирный** | `*текст*`, `**текст**` | Italic / Bold |
| **Зачёркнутый** | `~~текст~~` | Strikethrough |
| **Код в строке** | `` `код` `` | Courier New, 10 pt |
| **Ссылки** | `[текст](url)` | Настоящая гиперссылка, чёрная с подчёркиванием |
| **Изображения** | `![alt](путь)` | Картинка в документе; локальные пути считаются от файла-исходника, ссылки `http(s)` скачиваются |
| **Цитаты** | `> текст` | Стиль Quote, по ширине |
| **Разделители** | `---` | Горизонтальная линия |
| **Блоки кода** | ````` ```python ````` | Courier New 10 pt, язык — курсивной подписью сверху |
| **Списки** | `-`, `1.` | Bullet / Numbered, вложенность с отступом 18 pt на уровень |
| **Чек-листы** | `- [ ]`, `- [x]` | Символы ☐ и ☒ в начале строки |
| **Таблицы** | `\| … \|` | Таблица с границами, жирная шапка, выравнивание колонок из разметки |
| **Сноски** | `[^1]` и `[^1]: …` | Ссылка `[1]` в тексте и раздел «Footnotes» в конце |
| **Формулы** | `$…$`, `$$…$$` | Родные уравнения Word — см. [раздел ниже](#-формулы-latex) |

Кроме того: голые URL в тексте распознаются как ссылки автоматически, а одиночный перенос строки внутри абзаца сохраняется как перенос. Сырой HTML внутри Markdown не обрабатывается и в документ не переносится.

---

## 🧮 Формулы LaTeX

Главная возможность этой версии. Формулу можно записать четырьмя способами:

```markdown
Внутри строки: $E = mc^2$

Отдельным блоком:
$$\int_0^1 x^2\,dx = \frac{1}{3}$$

С номером уравнения:
$$a^2 + b^2 = c^2$$ (1)

В окружении amsmath:
\begin{align}
  f(x) &= ax + b \\
  g(x) &= cx + d
\end{align}
```

Все четыре превращаются в **настоящие уравнения Word (OMML)** — они открываются в редакторе формул, их можно править, они масштабируются вместе с текстом. Это не картинка и не текстовая имитация.

Из окружений amsmath поддерживаются `equation`, `multline`, `gather`, `align`, `alignat`, `flalign` и `eqnarray`. Многострочное окружение остаётся **одним** уравнением Word: `gather` и `multline` — со сложенными в столбик строками, `align` и родственные — ещё и с выравниванием по `&`, то есть знаки `=` встают друг под другом, как в LaTeX. Выравнивание записывается штатным маркером OMML `<m:aln/>` — тем же, которым пользуется сам Word.

### Что поддерживается

| Конструкция | Примеры |
|---|---|
| Дроби | `\frac`, `\dfrac`, `\tfrac` |
| Корни | `\sqrt{x}`, `\sqrt[3]{x}` |
| Индексы и степени | `x^2`, `a_i`, `x_i^2` |
| Греческие буквы | `\alpha`, `\beta`, … `\omega`, `\Gamma`, `\Delta`, … `\Omega` |
| Операторы и отношения | `\times`, `\div`, `\pm`, `\cdot`, `\leq`, `\geq`, `\neq`, `\approx`, `\equiv`, `\in`, `\subset`, `\cup`, `\cap`, `\forall`, `\exists`, `\to`, `\Rightarrow`, `\infty`, `\partial`, `\nabla` и другие |
| Имена функций | `\sin`, `\cos`, `\tan`, `\log`, `\ln`, `\exp`, `\det`, `\max`, `\min`, `\sup`, `\inf` — прямым шрифтом |
| Текст в формуле | `\text{}`, `\mathrm{}`, `\operatorname{}` — в том числе с кириллицей внутри |
| Начертания | `\mathbf`, `\mathit`, `\boldsymbol`, `\bm` |
| Крупные операторы | `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\oint`, `\bigcup`, `\bigcap`, `\bigoplus`, `\bigotimes`, `\bigvee`, `\bigwedge` — с пределами через `_` и `^` |
| Пределы | `\lim`, `\limsup`, `\liminf` с нижним пределом: `\lim_{x \to 0}` |
| Растягивающиеся скобки | `\left( … \right)`, `\left\{ … \right\}`, `\left\| … \right\|`, `\langle`, `\lfloor`, `\lceil`, а также `\left.` / `\right.` |
| Диакритика | `\hat`, `\widehat`, `\tilde`, `\widetilde`, `\bar`, `\vec`, `\dot`, `\ddot`, `\acute`, `\grave`, `\check`, `\breve` |
| Черта сверху и снизу | `\overline`, `\underline` |
| Биномиальный коэффициент | `\binom{n}{k}`, а также инфиксное `{n \choose k}` |
| Инфиксные дроби | `{a \over b}`, `{n \atop k}` — делят группу, в которой стоят |
| Матрицы | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases` |
| Таблицы формул | `\begin{array}{lcr} … \end{array}` — выравнивание колонок `l`, `c`, `r` переносится в Word |
| Стопка под знаком суммы | `\substack{i < j \\ i \in S}` |
| Перенос строки | `\\` в любой формуле, не только в матрице или amsmath: строки складываются в столбик |
| Выравнивание строк | `&` между строками многострочной формулы: `a &= b \\ c &= d` — знаки `=` встают друг под другом |
| Пробелы и экранирование | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_` |

### Границы поддержки

Четыре случая ниже отвергаются **намеренно**. Это не список «доделать позже»: у каждого есть причина, по которой правильного поведения просто не существует, — поэтому конвертер честно отказывается вместо того, чтобы выдать похожий, но неверный результат.

| Случай | Почему отвергается | Что делать |
|---|---|---|
| `\begin{array}{c\|c}` — вертикальная линейка | В матрице Word линейки между колонками не бывает, а молча выбросить её нельзя: расширенная матрица превратилась бы в обычную. Туда же `p{5cm}`, `@{…}`, `\hline` | Разбить на две матрицы или обойтись без линейки |
| `\matrix{…}`, `\cases{…}` — plain-TeX-запись | Это **другая** конструкция с другими правилами группировки, а не синоним окружения | Писать `\begin{matrix} … \end{matrix}` |
| `a \over b \over c` — два инфикса подряд | Неоднозначно: непонятно, что делить на что. Сам TeX такое тоже отвергает | Расставить скобки: `{{a \over b} \over c}` |
| `$a & b$` — одиночный `&` вне многострочной формулы | Выравнивать не с чем. А `Tom & Jerry` внутри `$…$` куда вероятнее забытое экранирование, чем формула | Писать `\&` |

Последний случай — не ограничение, а защита от опечатки: между строками многострочной формулы `&` работает и выравнивает (`a &= b \\ c &= d`), см. таблицу выше.

**Ничего не теряется молча.** Если конструкция не поддерживается, формула попадает в документ буквально, символ в символ, моноширинным шрифтом — а в итоговом диалоге появляется предупреждение с точным указанием, что именно не удалось:

```
Formula kept as text: "\begin{array}{c|c} a & b \end{array}"
(Column specification is not supported in \begin{array}: '|'
 (only 'l', 'c' and 'r' columns are))
```

### 💲 Знак доллара в обычном тексте

Это не ошибка и не ограничение, а договорённость — та же, что в Jupyter, Pandoc и MyST. Раз `$` открывает формулу, буквальный доллар в прозе пишется как `\$`.

Чтобы не заставлять экранировать вообще всё, конвертер разбирает частые случаи сам:

| Вы написали | Что получится | Предупреждение |
|---|---|---|
| `Стоит $5 и $10` | Текст как есть — цифра сразу после доллара формулу не открывает | нет |
| `Цена \$100` | `Цена $100` | нет |
| `$E = mc^2$` | Настоящее уравнение Word | нет |
| `$\text{путь}$` | Настоящее уравнение Word — кириллица внутри `\text{}` законна | нет |
| `Set $PATH and $HOME` | Текст как есть, ничего не потеряно | есть — похоже на прозу, а не на формулу |
| `Переменные $HOME и $PATH` | Текст как есть | есть — кириллица вне `\text{}` |

Последние две строки — единственные, где что-то попадает в итоговый диалог, и текст при этом сохраняется дословно:

```
Inline math "$PATH and $" contains no mathematical symbols and may be
ordinary prose rather than a formula; write a literal "$" as "\$".
```

---

## 🎨 Как оформляется документ Word

**Цвет.** Весь текст чёрный, включая заголовки. Шаблон python-docx по умолчанию красит заголовки через тему документа, поэтому цвета темы вычищаются принудительно.

**Заголовки.** Размеры считаются от выбранного размера текста. При 12 pt получается 18 / 16 / 14 / 13 / 12 / 12 pt для уровней 1–6; все уровни жирные, шестой дополнительно курсивный, чтобы отличаться от пятого. Заголовки используют выбранный шрифт — для этого из стилей вычищаются *шрифты темы*, потому что в OOXML атрибут темы перекрывает явно заданное имя шрифта.

**Выравнивание.** Абзацы, элементы списков, цитаты и сноски выровнены по ширине. Заголовки, блоки кода и ячейки таблиц — нет. Блочные формулы центрируются.

**Таблицы.** Границы записываются прямым форматированием, а не только стилем: часть просмотрщиков игнорирует границы уровня стиля, и таблица приезжает без сетки.

---

## 🏗️ Структура проекта

```
MDtoWORD/
├── 📦 mdtoword/                  # Пакет приложения (запуск: python -m mdtoword)
│   ├── __init__.py
│   ├── __main__.py               # Точка входа
│   ├── app.py                    # GUI на PyQt6
│   ├── converters.py             # Ядро конвертации без Qt, общее для GUI и MCP-сервера
│   ├── gfm_renderer.py           # Рендер GFM-разметки в документ Word
│   ├── latex_omml.py             # Разбор LaTeX и сборка уравнений OMML
│   ├── mcp_server.py             # MCP-сервер: три инструмента конвертации по stdio
│   ├── workflow.py               # Поиск исходников и раскладка результатов
│   └── theme.py                  # Тёмная и светлая темы, сохранение выбора
├── 📁 tests/                     # Тесты (unittest)
│   ├── test_drop_queue.py
│   ├── test_gui_theme.py
│   ├── test_conversion_workflow.py
│   ├── test_gfm_docx_renderer.py
│   ├── test_converters.py
│   ├── test_latex_omml.py
│   ├── test_mcp_server.py
│   └── test_packaging.py
├── 📁 scripts/
│   ├── build_macos.sh            # Сборка MDtoWORD.app (Apple Silicon)
│   ├── build_windows.ps1         # Сборка бандла и архива для Windows
│   └── launch_mdtoword.sh        # Запуск из Linux (bash)
├── 📁 packaging/
│   ├── MDtoWORD.desktop          # Ярлык рабочего стола (Linux)
│   └── windows_version_info.txt  # Метаданные версии для Windows-сборки
├── 📁 assets/                    # Иконки приложения (png, icns, ico)
├── 📁 docs/
│   ├── DESCRIPTION.md            # Описание репозитория (RU/EN)
│   ├── ИНСТРУКЦИЯ.txt            # Краткая инструкция (RU)
│   ├── INSTRUCTION_EN.txt        # Краткая инструкция (EN)
│   ├── 📁 design/                # Планы и спецификации разработки
│   └── 📁 releases/              # Заметки к выпускам (1.0, 1.1, 1.1.1)
├── 📄 MDtoWORD.spec              # Конфигурация PyInstaller (macOS)
├── 📋 requirements.txt           # Зависимости приложения
├── 📋 requirements-build.txt     # Зависимости сборки (PyInstaller)
├── 📋 requirements-mcp.txt       # Зависимости MCP-сервера
├── 📋 environment.yml            # Conda-окружение (Python 3.11)
└── 📖 README.md                  # Документация (этот файл)
```

---

## ⚙️ Технические детали

### Оформление по умолчанию
- **Шрифт**: Times New Roman, 12 pt
- **Цвет**: чёрный (RGB 0, 0, 0)
- **Код**: Courier New, 10 pt
- **Таблицы**: границы прямым форматированием, жирная шапка

### Поддерживаемые форматы
- **Markdown → Word**: вход `.md`, `.markdown` → выход `.docx`
- **Word → Markdown**: вход `.docx` → выход `.md`

---

## 🤖 MCP-сервер

MDtoWORD включает MCP-сервер, чтобы агенты могли выполнять те же конвертации, что и графический интерфейс.

Установите зависимости сервера:

```bash
python -m pip install -r requirements-mcp.txt
```

Подключите его в любом MCP-клиенте (пути должны быть абсолютными):

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

Для Claude Code:

```bash
claude mcp add mdtoword --scope user \
  -- /path/to/MDtoWord/.venv/bin/python -m mdtoword.mcp_server
```

### Инструменты

| Инструмент | Что делает |
| --- | --- |
| `markdown_to_word` | Конвертирует файлы и папки `.md` / `.markdown` в `.docx`: GFM, сноски, изображения и формулы LaTeX → уравнения OMML. |
| `word_to_markdown` | Конвертирует файлы и папки `.docx` в Markdown. С потерями: сохраняются заголовки, жирный, курсив и таблицы; всё остальное упрощается. |
| `preview_markdown` | Рендерит Markdown в памяти и сообщает только о том, что не переживёт конвертацию. Ничего не записывает на диск. |

Все три инструмента принимают пути, а не содержимое файлов, и работают с файлами
и папками вперемешку; папки просматриваются рекурсивно. Там, где два
конвертирующих инструмента пишут, существующий выходной файл перезаписывается
без предупреждения.

---

## 🛠️ Разработка

Тесты запускаются из корня проекта:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest \
    tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow \
    tests.test_gfm_docx_renderer tests.test_converters tests.test_latex_omml \
    tests.test_mcp_server tests.test_packaging
```

Сейчас в наборе **195 тестов**. Переменная `QT_QPA_PLATFORM=offscreen` нужна, чтобы тесты интерфейса работали без экрана.

Автономные сборки:

- `./scripts/build_macos.sh` — собирает `dist/MDtoWORD.app` для Apple Silicon: создаёт отдельное окружение, ставит зависимости, запускает PyInstaller по `MDtoWORD.spec` и подписывает результат ad-hoc-подписью;
- `scripts/build_windows.ps1` — собирает бандл для Windows, упаковывает его в `dist/MDtoWORD-Windows-x64.zip` и считает SHA-256.

---

## 🔧 Решение проблем

**Программа не запускается**
```bash
python --version            # нужен 3.10 или новее
pip install -r requirements.txt
```

**Ошибка кодировки**
Убедитесь, что `.md`-файлы сохранены в UTF-8.

**Таблица потеряла форматирование**
Проверьте синтаксис: у таблицы обязательно должна быть строка-разделитель `|---|---|`. Выравнивание колонок берётся из неё же — `:---`, `:---:`, `---:`.

**Формула не сконвертировалась**
Посмотрите предупреждение в итоговом диалоге: там названа конкретная конструкция, например `LaTeX environment is not supported: \begin{array}`. Сама формула при этом сохранена в документе буквально — перепишите её через поддерживаемую конструкцию из [таблицы выше](#что-поддерживается) и запустите конвертацию заново.

**Доллар в тексте превратился во что-то странное**
Пишите буквальный знак как `\$`. Если конвертер увидел `$…$` вокруг обычного текста, он оставит его как есть и предупредит об этом — но лучше экранировать сразу.

**Формула внутри ячейки таблицы осталась текстом**
Так и задумано: в ячейках формулы не конвертируются, а сохраняются вместе с долларами. Об этом сообщает предупреждение. Вынесите формулу из таблицы, если она должна быть уравнением.

**Блочная формула внутри списка потеряла отступ**
Блочные формулы всегда становятся отдельным центрированным абзацем, поэтому нумерация и отступ списка к ним не применяются.

**Изображение не попало в документ**
Локальные пути считаются относительно `.md`-файла, ссылки `http(s)` скачиваются с таймаутом 10 секунд. Если файл не найден или сеть недоступна, на его месте окажется `[alt-текст]`, а в диалоге появится предупреждение с адресом.

**Большая пачка файлов «подвешивает» окно**
Конвертация идёт в потоке интерфейса, поэтому на длинных очередях окно откликается вяло. Прогресс при этом обновляется — дождитесь конца.

**Word → Markdown отдаёт не всё**
Обратное направление намеренно упрощённое: разбираются заголовки по стилям, жирный и курсив, а таблицы дописываются в конец файла, а не на своё место в тексте.
