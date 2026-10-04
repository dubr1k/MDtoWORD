# MDtoWORD 1.2

## Русский

Выпуск про верность Markdown. Аудит стресс-документом нашёл дефекты, видные в любом реальном файле: сквозную нумерацию списков, «съехавшие» сноски, картинки шире страницы. Всё это исправлено. Кроме того, появились пресет ГОСТ 7.32, настоящие сноски Word, шаблоны документа и полноценное обратное направление Word → Markdown. MCP-сервер научился проверять формулы и готовые документы.

### Исправлено

| Было | Стало |
|---|---|
| Все нумерованные списки документа шли одним счётчиком: второй список начинался с 7, `5.` игнорировалось, вложенный список продолжал номер родителя | У каждого списка своя нумерация Word, `start` учитывается, вложенность — настоящие уровни списка |
| Второй абзац пункта списка становился новым пунктом, код внутри пункта терял отступ | Абзацы-продолжения и блоки внутри пункта стоят на уровне его текста, без номера |
| Сноски: «11. [1]» в конце документа, текст сноски отдельным абзацем | Настоящие сноски Word внизу страницы (`footnotes="section"` — прежний раздел, но без общего счётчика) |
| Картинка 3000 px выходила на 1058 мм при ширине страницы 216 мм | Картинки уменьшаются до ширины колонки с сохранением пропорций |
| YAML front matter превращался в горизонтальную линию и заголовок «title: …» | Титульный блок и свойства документа |
| Ссылки `[…](#раздел)` были внешними и не работали | Закладки на заголовках и внутренние ссылки |
| Одиночный перенос строки становился разрывом — текст с hard-wrap рассыпался | Одиночный перенос — пробел (CommonMark); `line_breaks="preserve"` возвращает прежнее |
| В ячейках таблиц терялись жирный, курсив, код, ссылки; формулы оставались текстом | Ячейки рендерятся как обычный текст, формулы в них — уравнения |
| Страница US Letter, язык документа en-US | A4, язык определяется по тексту |
| SVG и WebP не вставлялись | SVG — нативно, WebP и прочее — через PNG |
| Кириллица в имени файла картинки: «Image not found» | Путь декодируется |
| `<br>`, `<sub>`, `<details>` и прочий HTML печатались как текст | Белый список тегов, остальное — с предупреждением |
| Перенаправление при загрузке картинки могло увести на внутренний адрес | Загрузка только с публичных адресов, каждое перенаправление проверяется |

### Устойчивость

Отдельное ревью с фаззингом (≈500 случайных и враждебных документов, проверка схемы через officecli) нашло и закрыло:

- управляющие символы (ESC из вставленных логов терминала, form feed) роняли конвертацию — теперь удаляются с одним предупреждением;
- BOM в начале файла превращал первый заголовок или front matter в текст;
- значения front matter длиннее 255 символов и неверный тег `lang` обрывали конвертацию — теперь укорачиваются или игнорируются с предупреждением;
- сноска внутри сноски делала файл нечитаемым в LibreOffice — теперь её текст дописывается к внешней сноске;
- таблица на 1000 строк строилась минуту из-за квадратичного обхода ячеек — теперь за доли секунды;
- формулы и картинки внутри текста ссылки уезжали за ссылку; бейджи `[![…](…)](url)` стали кликабельными;
- `$$…$$` посреди предложения оставлял лишние знаки доллара;
- HTML-блоки: текст между `<` и `>`, ссылка на несколько абзацев, `<iframe/>` больше не теряют текст;
- загрузка картинки укладывается в общий таймаут на всех этапах (подключение, TLS, заголовки, тело), перебирает все проверенные адреса;
- SVG из Illustrator (с DOCTYPE) очищается перед вставкой;
- шаблон `.dotx` поддерживается, `.dotm` отклоняется с понятным сообщением.

### Новое в Markdown

- Подписи: одиночная картинка — рисунок «Рисунок N — …», абзац `Таблица: …` рядом с таблицей — подпись над ней. Нумерация — полями `SEQ`.
- Плашки GitHub `> [!NOTE]`, `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]`, `[!CAUTION]`.
- `H~2~O`, `x^2^`, `==выделение==`, списки определений.
- `[TOC]` и опция оглавления; оглавление сразу заполнено ссылками на заголовки.
- Нумерованные формулы: `\tag{n}` и `$$ … $$ (n)` — формула по центру, номер справа.
- Шапка таблицы повторяется на каждой странице, ширина колонок следует за содержимым.
- Предупреждения называют строку исходника и несут стабильный код.

### Формулы LaTeX

Из 147 конструкций, которые чаще всего пишут LLM, раньше конвертировались 83, теперь — 146. Осталась только `\sideset`: в OMML нельзя повесить индексы с обеих сторон большого оператора.

- Окружения `aligned`, `alignedat`, `split`, `gathered` внутри `$$…$$`, `smallmatrix`, `dcases`, `rcases`, `subarray`.
- Алфавиты `\mathbb`, `\mathcal`, `\mathscr`, `\mathfrak`, `\mathsf`, `\mathtt`, `\mathbfit`; `\textbf`, `\textit`, `\emph`.
- `\overset`, `\underset`, `\stackrel`, `\overbrace`/`\underbrace` с подписью, `\xrightarrow[…]{…}` и родственные стрелки.
- `\boxed`, `\cancel`, `\phantom`, `\hspace`, цвет `\color`/`\textcolor`.
- `\tag`, `\label`, `\nonumber`; `\limits`, `\displaystyle`, `\big`…`\Bigg`, `\middle`.
- `\cfrac`, `\dbinom`, `\pmod`, `\bmod`, `\Pr`, `\operatorname*`, `\not`, `\|`, `\lVert`, `\ket`/`\bra`/`\braket` и больше 150 новых символов.

Изменения поведения, о которых стоит знать:

- `\mathrm{…}` теперь прямой математический шрифт, а не буквальный текст: раньше `\mathrm{m^2}` печаталось как «m^2».
- `\epsilon` — ϵ, `\phi` — ϕ, как в LaTeX, MathJax и Word; раньше `\epsilon` совпадал с `\varepsilon`, а `\phi` и `\varphi` были перепутаны.
- Колонки `cases` выровнены влево, как в LaTeX.
- У `\max`, `\min`, `\sup`, `\inf` нижний индекс уходит под имя, как у `\lim`.
- `\color` действует до конца группы, строки или ячейки `&`, как в LaTeX, KaTeX и MathJax 3.

### Word → Markdown переписан

Раньше обратное направление брало только заголовки, жирный, курсив и таблицы, причём таблицы уезжали в конец файла, а текст гиперссылок терялся. Теперь документ обходится по порядку, и переносятся:

- заголовки (включая нумерованные), всё строчное форматирование с корректной расстановкой маркеров;
- ссылки, в том числе на закладки заголовков — как `#якоря`;
- списки с настоящей нумерацией Word, вложенностью и флажками;
- код, цитаты и врезки, таблицы на своих местах;
- изображения — в папку `<имя>_media/`;
- формулы — обратно в LaTeX;
- сноски и концевые сноски;
- метаданные — во front matter.

Спецсимволы Markdown экранируются, так что текст читается обратно без изменений. Чего в Markdown нет — объединённые ячейки, надписи, диаграммы, OLE-объекты, — попадает в предупреждения с количеством. Проверено на 537 реальных документах: ни одного падения, около 18 мс на файл.

### Оформление

- **Пресет ГОСТ 7.32**: A4, поля 30/15/20/20 мм, 14 pt, интервал 1,5, абзацный отступ 1,25 см, подписи по ГОСТ, номер страницы.
- **Шаблон** (`template=` в MCP): стили, поля и колонтитулы из своего `.docx`.
- Блоки кода — стиль «Source Code» с фоном и рамкой; код в строке — на сером фоне и в размер окружающего текста.

### Интерфейс

- Выбор стиля «Обычный / ГОСТ 7.32» и флажок «Оглавление».
- Предупреждения в итоговом окне с номером строки: `отчёт.md:42: …`.

### MCP-сервер

- Новые инструменты `check_latex` и `inspect_docx`, ресурс `mdtoword://guide/markdown`, prompt `prepare_markdown_for_word`.
- Опции `preset`, `page_size`, `language`, `line_breaks`, `template`, `toc`, `footnotes`.
- **Несовместимое изменение:** `warnings` — объекты `{message, code, line}` вместо строк.
- Аннотации инструментов, работа в отдельном потоке, уведомления о прогрессе.
- Установка пакетом: `pip install ".[mcp]"` и команда `mdtoword-mcp`.

## English

A release about Markdown fidelity. A stress-document audit found defects that show up in any real file — document-wide list numbering, footnotes in the wrong place, pictures wider than the page — and all of them are fixed. On top of that come a GOST 7.32 preset, native Word footnotes, document templates and a full Word → Markdown direction. The MCP server can now check formulas and finished documents.

### Fixed

- Every list now has its own Word numbering; `5.` starts at 5; nesting uses real list levels.
- A list item's second paragraph (or code block) stays inside the item instead of becoming a new item.
- Footnotes are native Word footnotes (`footnotes="section"` keeps a numbered section at the end, without sharing the list counter).
- Pictures are scaled to the text width, aspect ratio kept.
- YAML front matter becomes a title block and the document properties instead of a rule and an H2.
- `[…](#heading)` links work: headings carry bookmarks.
- A single newline inside a paragraph is a space (CommonMark); `line_breaks="preserve"` restores the old behaviour.
- Table cells keep bold, italic, code, links, and their formulas become equations.
- A4 instead of US Letter; the document language is detected from the text.
- SVG embeds natively, WebP and friends via PNG; Cyrillic file names resolve.
- A small HTML whitelist is honoured; other tags are stripped with a warning.
- Remote images are fetched from public addresses only, with every redirect re-checked.

### Robustness

A separate review with fuzzing (≈500 random and hostile documents, schema-checked with officecli) found and closed: control characters crashing a conversion (now removed with a warning); a BOM hiding the first heading or the front matter; front-matter values over 255 characters and invalid `lang` tags aborting a file; a footnote referenced inside a footnote making the file unreadable in LibreOffice; quadratic table rendering (a 1000-row table took a minute); math and images escaping link text; stray dollars around inline `$$…$$`; lost text in HTML blocks; remote downloads exceeding their total timeout while connecting or reading headers, and trying only the first resolved address; Illustrator SVGs with a DOCTYPE; `.dotx` templates (now supported).

### LaTeX

Of the 147 constructs an LLM most often writes, 83 used to convert; now 146 do. The one left is `\sideset`, which OMML cannot express. New: `aligned`/`alignedat`/`split`/`gathered` inside `$$…$$`, `smallmatrix`, `dcases`, `rcases`; math alphabets (`\mathbb`, `\mathcal`, `\mathscr`, `\mathfrak`, `\mathsf`, `\mathtt`); `\overset`, `\underset`, `\overbrace`/`\underbrace`, `\xrightarrow`; `\boxed`, `\cancel`, `\phantom`, colour; `\tag`, `\label`, `\limits`, `\big`…`\Bigg`, `\middle`; `\cfrac`, `\pmod`, `\operatorname*`, `\not`; 150+ more symbols.

Behaviour changes: `\mathrm` is upright math rather than literal text; `\epsilon` is ϵ and `\phi` is ϕ as in LaTeX; `cases` columns are left-aligned; `\max`/`\min`/`\sup`/`\inf` put their subscript underneath; `\color` lasts to the end of the group.

### Word → Markdown rewritten

The document is walked in order (tables stay in place) and carries over headings (numbered ones too), inline formatting with correctly placed markers, links including heading anchors, real Word list numbering, code, quotes and callouts, tables, pictures (to `<name>_media/`), equations back to LaTeX, footnotes and endnotes, and metadata as front matter. Markdown-significant characters are escaped so the text reads back unchanged; what Markdown cannot hold is reported with a count. Tested on 537 real documents without a crash, about 18 ms per file.

### New

- Figure and table captions with `SEQ` numbering; GitHub alerts; `H~2~O`, `x^2^`, `==mark==`; definition lists; `[TOC]`; numbered equations via `\tag{n}`; repeating table headers.
- GOST 7.32 preset and reference templates.
- GUI: style selector and table-of-contents checkbox; warnings show the source line.
- MCP: `check_latex`, `inspect_docx`, a Markdown guide resource and a prompt; document options; structured warnings `{message, code, line}` (**breaking** for clients that read strings); tool annotations; progress notifications; `pip install ".[mcp]"` with `mdtoword-mcp`.
