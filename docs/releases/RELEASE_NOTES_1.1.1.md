# MDtoWORD 1.1.1

## Русский

Точечный выпуск поверх 1.1. Закрывает список «Чего пока нет» из README: конструкции LaTeX, которые раньше попадали в документ текстом с предупреждением, теперь становятся настоящими уравнениями Word. Плюс исправление в генерируемом OOXML, из-за которого `\mathbf` формально нарушал схему Word.

### Новые конструкции LaTeX

| Конструкция | Что делает |
|---|---|
| `\begin{array}{lcr} … \end{array}` | Таблица формул; выравнивание колонок `l`, `c`, `r` переносится в Word |
| `\substack{i < j \\ i \in S}` | Стопка строк под знаком суммы |
| `{n \choose k}` | Биномиальный коэффициент инфиксной записью — то же, что `\binom{n}{k}` |
| `{a \over b}` | Дробь инфиксной записью |
| `{n \atop k}` | Стопка без черты |
| `\\` | Перенос строки в **любой** формуле, а не только в матрице или окружении amsmath: строки складываются в столбик |
| `&` | Выравнивание строк многострочной формулы: `a &= b \\ c &= d` ставит знаки `=` друг под другом |

### Окружения amsmath стали одним уравнением

Раньше многострочное окружение разрезалось по `\\` на отдельные абзацы, по одному уравнению на строку, и выравнивание по `&` терялось.

Теперь `gather` и `multline` остаются **одним** уравнением Word со сложенными в столбик строками, а `align`, `alignat`, `flalign` и `eqnarray` — ещё и с настоящим выравниванием: знаки `=` встают друг под другом, как в LaTeX. Выравнивание записывается штатным маркером OMML `<m:aln/>` — тем же, которым пользуется сам Word.

### Исправление: `\mathbf` нарушал схему OOXML

В OOXML элементы `<m:nor>` и `<m:sty>` внутри `<m:rPr>` — взаимоисключающий выбор, а не последовательность. `\mathbf{x}` выдавал оба сразу, потому что конструкция одновременно «прямая» и «жирная». Такой документ формально не проходит проверку по схеме ISO/IEC 29500-4.

Значение `<m:sty m:val="b"/>` само по себе означает «жирный прямой», поэтому `<m:nor>` там был лишним и убран. `<m:nor>` остался за нежирным буквальным текстом `\text{…}`.

Ошибка присутствовала с 1.1. На практике Word её обычно переваривал, но строгие валидаторы и сторонние читатели `.docx` — не обязаны.

### Метаданные версии Windows

`packaging/windows_version_info.txt` оставался на `1.0.0` и в выпуске 1.1 — свойства `MDtoWORD.exe` показывали неправильную версию. Приведён к `1.1.1`.

### Чего по-прежнему нет

Только то, чему в OMML нет соответствия:

- вертикальные линейки в `array` — `\begin{array}{c|c}`, а также `p{5cm}`, `@{…}` и `\hline`;
- одиночный `&` в однострочной формуле — выравнивать не с чем, а `Tom & Jerry` внутри `$…$` вероятнее забытое экранирование, поэтому такой `&` отвергается: пишите `\&`;
- plain-TeX-запись окружений — `\matrix{…}`, `\cases{…}`;
- два инфикса в одной группе — `a \over b \over c` неоднозначно, и сам TeX такое отвергает.

Как и раньше, ничего не теряется молча: неподдержанная конструкция попадает в документ дословно, моноширинным шрифтом, с предупреждением в итоговом диалоге.

### Загрузки

- **macOS 12+ / Apple Silicon:** `MDtoWORD-macOS-arm64.zip`
- **Windows 10/11 / x64:** `MDtoWORD-Windows-x64.zip`

Распакуйте архив и запустите приложение. В macOS при первом запуске может потребоваться нажать правой кнопкой на приложении и выбрать «Открыть». Windows SmartScreen также может показать предупреждение, поскольку сборки не подписаны коммерческими сертификатами.

---

## English

A point release on top of 1.1. It closes the README's "what isn't there yet" list: LaTeX constructs that previously landed in the document as text with a warning now become real Word equations. Plus a fix to the generated OOXML, where `\mathbf` was formally violating Word's schema.

### New LaTeX constructs

| Construct | What it does |
|---|---|
| `\begin{array}{lcr} … \end{array}` | A formula table; the `l`, `c`, `r` column alignment carries into Word |
| `\substack{i < j \\ i \in S}` | Stacked lines under a summation sign |
| `{n \choose k}` | Binomial coefficient written infix — the same as `\binom{n}{k}` |
| `{a \over b}` | Fraction written infix |
| `{n \atop k}` | A stack with no bar |
| `\\` | A line break in **any** formula, not just inside a matrix or an amsmath environment: the lines stack |
| `&` | Alignment between the lines of a multi-line formula: `a &= b \\ c &= d` puts the `=` signs under one another |

### amsmath environments became one equation

Previously a multi-line environment was cut on `\\` into separate paragraphs, one equation per line, and the `&` alignment was discarded.

Now `gather` and `multline` stay **one** Word equation with their lines stacked, and `align`, `alignat`, `flalign` and `eqnarray` are stacked *and* genuinely aligned: the `=` signs line up the way LaTeX draws them. The alignment is written with OMML's own `<m:aln/>` marker — the one Word itself uses.

### Fix: `\mathbf` violated the OOXML schema

In OOXML the `<m:nor>` and `<m:sty>` elements inside `<m:rPr>` are a mutually exclusive choice, not a sequence. `\mathbf{x}` emitted both at once, because the construct is upright and bold simultaneously. Such a document does not formally validate against the ISO/IEC 29500-4 schema.

The value `<m:sty m:val="b"/>` already means "bold upright" on its own, so `<m:nor>` was redundant there and has been removed. `<m:nor>` remains for the unbolded literal text of `\text{…}`.

The bug has been present since 1.1. In practice Word usually digested it, but strict validators and third-party `.docx` readers are under no obligation to.

### Windows version metadata

`packaging/windows_version_info.txt` was still on `1.0.0` in the 1.1 release too — the properties of `MDtoWORD.exe` showed the wrong version. It now reads `1.1.1`.

### What still isn't there

Only what OMML has no equivalent for:

- vertical rules in `array` — `\begin{array}{c|c}`, and likewise `p{5cm}`, `@{…}` and `\hline`;
- a lone `&` in a single-line formula — there is nothing to align against, and `Tom & Jerry` inside `$…$` is more likely a missing escape, so that `&` is refused: write `\&`;
- the plain-TeX spelling of environments — `\matrix{…}`, `\cases{…}`;
- two infix commands in one group — `a \over b \over c` is ambiguous, and TeX itself rejects it.

As before, nothing is lost silently: an unsupported construct lands in the document character for character, in a monospace font, with a warning in the result dialog.

### Downloads

- **macOS 12+ / Apple Silicon:** `MDtoWORD-macOS-arm64.zip`
- **Windows 10/11 / x64:** `MDtoWORD-Windows-x64.zip`

Unpack the archive and run the application. On macOS the first launch may require right-clicking the app and choosing "Open". Windows SmartScreen may also show a warning, since the builds are not signed with commercial certificates.
