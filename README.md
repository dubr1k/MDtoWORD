# RU
# 📄 MDtoWORD — конвертер Markdown в Word

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-PyQt6-orange?style=for-the-badge)
![LaTeX](https://img.shields.io/badge/LaTeX-native%20OMML-red?style=for-the-badge)

**Настольное приложение, которое превращает GitHub Flavored Markdown в аккуратный документ Word — вместе с формулами, которые в Word можно редактировать**

[🚀 Быстрый старт](#-быстрый-старт) • [🧭 Интерфейс](#-интерфейс) • [📝 Markdown](#-что-понимается-в-markdown) • [🧮 Формулы](#-формулы-latex) • [🔧 Решение проблем](#-решение-проблем) • [🇺🇸 EN](#EN)

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
python md_to_word_converter.py
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

python md_to_word_converter.py
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

Из окружений amsmath поддерживаются `equation`, `multline`, `gather`, `align`, `alignat` и `flalign`. Многострочное окружение сначала пробуется целиком, а если так не выходит — разбивается по `\\` на отдельные уравнения, по одному на строку.

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
| Биномиальный коэффициент | `\binom{n}{k}` |
| Матрицы | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases` |
| Пробелы и экранирование | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_` |

### Чего пока нет

Окружения `array` и `substack`, команды `\choose`, `\over`, `\atop`, а также перенос строки `\\` вне матрицы или окружения amsmath.

**Ничего не теряется молча.** Если конструкция не поддерживается, формула попадает в документ буквально, символ в символ, моноширинным шрифтом — а в итоговом диалоге появляется предупреждение с точным указанием, что именно не удалось:

```
Formula kept as text: "\begin{array}{c} x \end{array}"
(LaTeX environment is not supported: \begin{array})
```

### ⚠️ Знак доллара в обычном тексте

Раз `$` открывает формулу, буквальный доллар в прозе нужно писать как `\$` — та же договорённость, что в Jupyter, Pandoc и MyST.

Суммы вида `$5` и `$10` распознаются как деньги и остаются нетронутыми: цифра сразу после доллара формулу не открывает. А вот пара долларов вокруг слов — `Set $PATH and $HOME` — формально выглядит как формула. Такой текст сохраняется как есть, но конвертер об этом предупреждает:

```
Inline math "$PATH and $" contains no mathematical symbols and may be
ordinary prose rather than a formula; write a literal "$" as "\$".
```

Кириллица внутри `$…$` тоже считается признаком обычного текста — если только она не завёрнута в `\text{}`, как и положено настоящей формуле.

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
├── 📄 md_to_word_converter.py    # GUI на PyQt6 и оба конвертера
├── 📄 gfm_docx_renderer.py       # Рендер GFM-разметки в документ Word
├── 📄 latex_omml.py              # Разбор LaTeX и сборка уравнений OMML
├── 📄 conversion_workflow.py     # Поиск исходников и раскладка результатов
├── 📄 gui_theme.py               # Тёмная и светлая темы, сохранение выбора
├── 📁 tests/                     # Тесты (unittest)
│   ├── test_drop_queue.py
│   ├── test_gui_theme.py
│   ├── test_conversion_workflow.py
│   ├── test_gfm_docx_renderer.py
│   ├── test_markdown_converter.py
│   └── test_latex_omml.py
├── 📁 scripts/
│   ├── build_macos.sh            # Сборка MDtoWORD.app (Apple Silicon)
│   ├── build_windows.ps1         # Сборка бандла и архива для Windows
│   └── launch_mdtoword.sh        # Запуск из Linux (bash)
├── 📁 packaging/
│   └── windows_version_info.txt  # Метаданные версии для Windows-сборки
├── 📁 assets/                    # Иконки приложения (png, icns, ico)
├── 📁 docs/superpowers/          # Планы и спецификации разработки
├── 📄 MDtoWORD.spec              # Конфигурация PyInstaller (macOS)
├── 📋 requirements.txt           # Зависимости приложения
├── 📋 requirements-build.txt     # Зависимости сборки (PyInstaller)
├── 📋 environment.yml            # Conda-окружение (Python 3.11)
├── 📖 README.md                  # Документация (этот файл)
├── 📄 DESCRIPTION.md             # Описание репозитория (RU/EN)
├── 📄 ИНСТРУКЦИЯ.txt             # Краткая инструкция (RU)
└── 📄 INSTRUCTION_EN.txt         # Краткая инструкция (EN)
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

## 🛠️ Разработка

Тесты запускаются из корня проекта:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest \
    tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow \
    tests.test_gfm_docx_renderer tests.test_markdown_converter tests.test_latex_omml
```

Сейчас в наборе **111 тестов**. Переменная `QT_QPA_PLATFORM=offscreen` нужна, чтобы тесты интерфейса работали без экрана.

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

---

# EN

# 📄 MDtoWORD — Markdown to Word converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-PyQt6-orange?style=for-the-badge)
![LaTeX](https://img.shields.io/badge/LaTeX-native%20OMML-red?style=for-the-badge)

**A desktop app that turns GitHub Flavored Markdown into a clean Word document — formulas included, and still editable once they get there**

[🚀 Quick Start](#-quick-start) • [🧭 Interface](#-the-interface) • [📝 Markdown](#-what-markdown-is-understood) • [🧮 Formulas](#-latex-formulas) • [🔧 Troubleshooting](#-troubleshooting) • [🇷🇺 RU](#RU)

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
python md_to_word_converter.py
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

python md_to_word_converter.py
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

The supported amsmath environments are `equation`, `multline`, `gather`, `align`, `alignat` and `flalign`. A multi-line environment is first tried as a whole; if that fails it is split on `\\` into one equation per line.

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
| Binomial coefficient | `\binom{n}{k}` |
| Matrices | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases` |
| Spacing and escapes | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_` |

### What isn't there yet

The `array` and `substack` environments, the `\choose`, `\over` and `\atop` commands, and a `\\` line break outside a matrix or an amsmath environment.

**Nothing is lost silently.** When a construct isn't supported, the formula goes into the document verbatim, character for character, in a monospace font — and the result dialog carries a warning naming exactly what failed:

```
Formula kept as text: "\begin{array}{c} x \end{array}"
(LaTeX environment is not supported: \begin{array})
```

### ⚠️ A literal dollar sign in prose

Since `$` opens a formula, a literal dollar in running text should be written `\$` — the same convention Jupyter, Pandoc and MyST use.

Amounts like `$5` and `$10` are recognised as money and stay intact: a digit straight after a dollar does not open a formula. A pair of dollars around words, though — `Set $PATH and $HOME` — technically looks like a formula. That text is kept as written, but the converter warns you about it:

```
Inline math "$PATH and $" contains no mathematical symbols and may be
ordinary prose rather than a formula; write a literal "$" as "\$".
```

Cyrillic inside `$…$` counts as another sign of ordinary prose — unless it is wrapped in `\text{}`, which is how a genuine formula writes a word.

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
├── 📄 md_to_word_converter.py    # PyQt6 GUI and both converters
├── 📄 gfm_docx_renderer.py       # Renders GFM markup into a Word document
├── 📄 latex_omml.py              # Parses LaTeX and builds OMML equations
├── 📄 conversion_workflow.py     # Source discovery and output path allocation
├── 📄 gui_theme.py               # Dark and light themes, persisted choice
├── 📁 tests/                     # Test suite (unittest)
│   ├── test_drop_queue.py
│   ├── test_gui_theme.py
│   ├── test_conversion_workflow.py
│   ├── test_gfm_docx_renderer.py
│   ├── test_markdown_converter.py
│   └── test_latex_omml.py
├── 📁 scripts/
│   ├── build_macos.sh            # Builds MDtoWORD.app (Apple Silicon)
│   ├── build_windows.ps1         # Builds the Windows bundle and archive
│   └── launch_mdtoword.sh        # Linux (bash) launcher
├── 📁 packaging/
│   └── windows_version_info.txt  # Version metadata for the Windows build
├── 📁 assets/                    # Application icons (png, icns, ico)
├── 📁 docs/superpowers/          # Development plans and specifications
├── 📄 MDtoWORD.spec              # PyInstaller configuration (macOS)
├── 📋 requirements.txt           # Application dependencies
├── 📋 requirements-build.txt     # Build dependencies (PyInstaller)
├── 📋 environment.yml            # Conda environment (Python 3.11)
├── 📖 README.md                  # Documentation (this file)
├── 📄 DESCRIPTION.md             # Repository description (RU/EN)
├── 📄 ИНСТРУКЦИЯ.txt             # Quick guide (RU)
└── 📄 INSTRUCTION_EN.txt         # Quick guide (EN)
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

## 🛠️ Development

Run the tests from the project root:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest \
    tests.test_drop_queue tests.test_gui_theme tests.test_conversion_workflow \
    tests.test_gfm_docx_renderer tests.test_markdown_converter tests.test_latex_omml
```

The suite currently holds **111 tests**. `QT_QPA_PLATFORM=offscreen` lets the interface tests run without a display.

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
