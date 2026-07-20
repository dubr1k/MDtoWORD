# MDtoWORD 1.1

## Русский

Второй релиз MDtoWORD. Главное — формулы LaTeX теперь становятся настоящими уравнениями Word, а не текстом с долларами. Кроме этого переработан интерфейс и исправлено оформление получающихся документов.

### Формулы LaTeX

Формулы, записанные как `$...$` в строке, `$$...$$` отдельным блоком или в окружениях `equation`, `align`, `gather`, `multline`, `alignat`, `flalign`, попадают в документ **нативными уравнениями Word** — их можно открыть в редакторе формул и править.

Поддерживаются дроби, корни со степенью, верхние и нижние индексы, греческие буквы и операторы, `\text{}` с кириллицей внутри формул, суммы и интегралы с пределами, `\lim`, растягивающиеся скобки `\left`/`\right`, акценты `\hat` `\vec` `\bar` `\overline`, биномиальные коэффициенты и матрицы — `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases`.

Всё, что конвертер не понял, сохраняется в документе дословно и сопровождается предупреждением с названием конкретной конструкции. Ничего не теряется молча.

До этого релиза формулы не только оставались текстом, но и портились: `$a*b*c$` превращалось в `$abc$` — звёздочки съедались как курсив, и в документ уходило математически другое выражение. Это исправлено.

**Важно:** поскольку `$` ограничивает формулу, литеральный знак доллара в тексте пишется как `\$`. Суммы вида `$5` обрабатываются правильно и остаются на месте, но пара вроде `Set $PATH and $HOME` была бы прочитана как формула — теперь конвертер предупреждает об этом вместо молчаливой порчи. Так же устроены Jupyter, Pandoc и MyST.

### Оформление документа Word

- Весь текст чёрный, включая заголовки. Раньше шаблон подсовывал им синий цвет из темы.
- Заголовки получили размерную иерархию, выведенную из выбранного вами кегля: при 12pt это 18/16/14/13/12/12pt для уровней с первого по шестой. Все жирные, шестой курсивом.
- Заголовки используют выбранный в интерфейсе шрифт. Раньше шаблон навязывал им шрифт темы независимо от настройки.
- Абзацы, списки, цитаты и сноски выровнены по ширине. Заголовки, блоки кода и ячейки таблиц — нет.
- Таблицы получили настоящие границы, записанные прямым форматированием: границы из стиля отрисовывают не все просмотрщики.
- Таблицы уважают выравнивание колонок из Markdown — `|:---|---:|:---:|` даёт влево, вправо и по центру.
- Заголовок раздела сносок следует языку интерфейса.

### Интерфейс

- Перетаскивание файлов **и папок** на любое место окна; папки просматриваются рекурсивно.
- Очередь файлов: добавление, удаление выбранных, очистка. Кнопки неактивны, когда действовать не над чем.
- Тёмная и светлая темы с переключателем; выбор запоминается между запусками.
- Окно уменьшено с 944 до 772 пикселей по минимальной высоте — теперь помещается на экран ноутбука.
- Пунктирная зона приёма кликабельна и открывает выбор файлов.
- Вкладка «Текст» скрывается в режиме Word → Markdown, где она не нужна.
- Прогресс-бар резервирует своё место и больше не дёргает вёрстку.
- Во время конвертации интерфейс блокируется, а очередь обрабатывается по снимку — добавление файла в процессе больше не роняет конвертацию.

### Загрузки

- **macOS 12+ / Apple Silicon:** `MDtoWORD-macOS-arm64.zip`
- **Windows 10/11 / x64:** `MDtoWORD-Windows-x64.zip`

Распакуйте архив и запустите приложение. В macOS при первом запуске может потребоваться нажать правой кнопкой на приложении и выбрать «Открыть». Windows SmartScreen также может показать предупреждение, поскольку сборки не подписаны коммерческими сертификатами.

### Известные ограничения

- Конвертация выполняется в потоке интерфейса, поэтому очень большие пакеты подтормаживают.
- Формулы внутри ячеек таблиц сохраняются текстом с предупреждением.
- Выключная формула внутри списка теряет отступ списка.

---

## English

The second release of MDtoWORD. The headline change: LaTeX formulas now become real Word equations instead of text with dollar signs. The interface has also been reworked and the formatting of the produced documents fixed.

### LaTeX formulas

Formulas written as inline `$...$`, display `$$...$$`, or inside the `equation`, `align`, `gather`, `multline`, `alignat` and `flalign` environments are written into the document as **native Word equations** — you can open them in the equation editor and edit them.

Supported: fractions, roots with an optional degree, superscripts and subscripts, Greek letters and operators, `\text{}` including Cyrillic inside formulas, sums and integrals with limits, `\lim`, stretchy `\left`/`\right` delimiters, the `\hat` `\vec` `\bar` `\overline` accents, binomial coefficients, and the `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix` and `cases` environments.

Anything the converter does not understand is kept in the document character for character and reported as a warning naming the exact construct. Nothing is lost silently.

Before this release formulas were not merely left as text — they were corrupted: `$a*b*c$` became `$abc$`, the asterisks eaten as emphasis, so a mathematically different expression reached the document. That is fixed.

**Note:** because `$` delimits a formula, a literal dollar sign in prose should be written `\$`. Amounts such as `$5` are handled correctly and stay intact, but a pair like `Set $PATH and $HOME` would otherwise be read as a formula — the converter now warns about that shape instead of silently mangling it. Jupyter, Pandoc and MyST work the same way.

### Word document formatting

- All text is black, headings included. The template used to give them a blue theme colour.
- Headings gained a size hierarchy derived from your chosen body size: at 12pt that is 18/16/14/13/12/12pt for levels one through six. All bold, level six italic.
- Headings use the font chosen in the interface. The template used to impose its theme font regardless of the setting.
- Paragraphs, lists, quotes and footnotes are justified. Headings, code blocks and table cells are not.
- Tables have real borders written as direct formatting, because style-level borders are not drawn by every viewer.
- Tables honour the column alignment from the Markdown spec — `|:---|---:|:---:|` gives left, right and centre.
- The footnotes section heading follows the interface language.

### Interface

- Drag files **and folders** onto anywhere in the window; folders are scanned recursively.
- A file queue you can add to, remove selected entries from, or clear. The buttons are disabled when there is nothing to act on.
- Dark and light themes with a toggle; the choice is remembered between launches.
- The window's minimum height dropped from 944 to 772 pixels, so it now fits a laptop screen.
- The dashed drop zone is clickable and opens the file picker.
- The Text tab is hidden in Word → Markdown mode, where it does not apply.
- The progress bar reserves its row and no longer makes the layout jump.
- The interface is locked during a conversion and the queue is processed from a snapshot, so adding a file mid-run no longer breaks the batch.

### Downloads

- **macOS 12+ / Apple Silicon:** `MDtoWORD-macOS-arm64.zip`
- **Windows 10/11 / x64:** `MDtoWORD-Windows-x64.zip`

Unpack the archive and run the application. On macOS the first launch may require right-clicking the app and choosing "Open". Windows SmartScreen may also show a warning, since the builds are not signed with commercial certificates.

### Known limitations

- Conversion runs on the interface thread, so very large batches feel sluggish.
- Formulas inside table cells are kept as text with a warning.
- A display formula inside a list loses the list indent.
