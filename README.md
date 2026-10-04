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

MDtoWORD takes your `.md` files — one, a dozen, or a whole folder — and drops finished `.docx` files beside them. The markup is not approximated: headings become Word styles with bookmarks you can link to, lists get real Word numbering, footnotes become native Word footnotes at the bottom of the page, tables keep their formatting, and a formula like `$E = mc^2$` becomes a **native, editable Word equation** rather than an image or a line of plain text. A GOST 7.32 preset lays the document out the way Russian reports and theses expect.

It also works the other way round: a `.docx` turns back into Markdown with headings, lists, tables, links, images, footnotes and equations in place — see [Word → Markdown](#-word--markdown).

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
- **mcp 1.28.1** — only for the [MCP server](#-mcp-server), installed separately

```bash
pip install -r requirements.txt
```

### Conda (virtual environment)

```bash
# Create the environment from environment.yml (Python 3.12)
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

**Appearance.** A font dropdown (Arial, Times New Roman, Calibri, Georgia, Helvetica, Courier New) and a size field from 6 to 72 pt set the document's base formatting. **Style** switches between *Standard* and *GOST 7.32* — A4, margins 30/15/20/20 mm, 1.5 line spacing, a 1.25 cm first-line indent and 14 pt (the size follows the style unless you changed it yourself). **Table of contents** puts a TOC at the top of the document.

**Themes and language.** The footer holds a round theme button (☀ / ☾) that switches between dark and light. The choice is stored via `QSettings` and restored on the next launch. The neighbouring **EN / RU** button switches the interface language.

While a batch runs you get a progress bar and the name of the current file; at the end, a dialog with the number of successes, errors and warnings. A warning names the source line it comes from: `report.md:42: Image not found: chart.png`.

---

## 📝 What Markdown is understood

The markup goes through a CommonMark + GitHub Flavored Markdown parser with a few widely used extensions.

| Element | Markdown | Result in Word |
|---|---|---|
| **Headings** | `# … ######` | Heading 1–6 styles with a size hierarchy, each carrying a bookmark |
| **Paragraphs** | plain text | Normal style, justified. A single newline inside a paragraph is a space, as in CommonMark |
| **Line breaks** | two trailing spaces or a trailing `\` | A line break |
| **Italic / bold / strikethrough** | `*text*`, `**text**`, `~~text~~` | Italic / Bold / Strikethrough |
| **Sub-, superscript, highlight** | `H~2~O`, `x^2^`, `==text==` | Subscript, superscript, yellow highlight |
| **Inline code** | `` `code` `` | Courier New on a light grey background, sized to the surrounding text |
| **Links** | `[text](url)`, `<url>`, a bare `https://…` | A real hyperlink, black and underlined |
| **Links to headings** | `[see](#heading-slug)` | An internal link to the heading's bookmark (GitHub slug rules, Cyrillic included) |
| **Images** | `![alt](path "title")` | Embedded picture scaled to the text width. PNG, JPEG, GIF, BMP, TIFF, WebP and SVG. Local paths resolve against the source file; `http(s)` URLs are downloaded (GUI behaviour — the [MCP server](#-mcp-server) does not fetch them by default) |
| **Figures** | an image alone in its paragraph | Centred picture with a numbered caption from the title or the alt text: "Figure 1: …", or "Рисунок 1 — …" in a Russian document |
| **Blockquotes** | `> text`, `> > nested` | Indented per level with a rule on the left; lists and code inside keep the quote's indent |
| **GitHub alerts** | `> [!NOTE]`, `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]`, `[!CAUTION]` | A shaded callout with a bold title |
| **Thematic breaks** | `---` | Horizontal rule |
| **Code blocks** | ````` ```python ````` or indented by four spaces | "Source Code" style: monospace, grey background, thin frame, whitespace kept |
| **Lists** | `-`, `1.`, `5.` | Real Word numbering: every list restarts, `5.` starts at 5, nesting up to nine levels, items with several paragraphs or code |
| **Task lists** | `- [ ]`, `- [x]` | ☐ and ☒ at the start of the item |
| **Definition lists** | `Term` then `: definition` | Bold term, indented definition |
| **Tables** | `\| … \|` | Bordered table: bold header row repeated on every page, column alignment from the markup, inline formatting, links and formulas inside cells |
| **Table captions** | `Table: …`, `Таблица: …` or `: …` right before or after a table | A numbered caption above the table |
| **Footnotes** | `[^1]` and `[^1]: …` | Native Word footnotes at the bottom of the page |
| **Front matter** | a `---` YAML block at the very top | A title block (title, subtitle, author, date) and the document properties (title, author, subject, keywords, language) |
| **Table of contents** | `[TOC]` on its own line | A Word TOC field, pre-filled with linked headings; Word adds page numbers when it updates fields |
| **Formulas** | `$…$`, `$$…$$`, amsmath | Native Word equations — see [below](#-latex-formulas) |
| **HTML** | `<br>`, `<sub>`, `<sup>`, `<kbd>`, `<mark>`, `<u>`, `<del>`, `<b>`, `<i>`, `<code>`, `<a href>`, `<img>`, `<details>` | The matching formatting. Comments are dropped; any other tag is removed with a warning, its text kept |

**Nothing is lost silently.** Whatever cannot be represented faithfully — a missing image, an unsupported formula, an unknown HTML tag, a link to a heading that does not exist, a footnote nothing refers to — ends up in the result dialog together with the line of the Markdown source it comes from.

---

## 🧮 LaTeX formulas

The headline feature of this release. A formula can be written four ways:

```markdown
Inline: $E = mc^2$

As a display block:
$$\int_0^1 x^2\,dx = \frac{1}{3}$$

With an equation number:
$$a^2 + b^2 = c^2$$ (1)
$$E = mc^2 \tag{2}$$

In an amsmath environment:
\begin{align}
  f(x) &= ax + b \\
  g(x) &= cx + d
\end{align}
```

All four become **real Word equations (OMML)** — they open in the equation editor, they can be edited, and they scale with the text. Not a picture, not a text imitation.

The supported amsmath environments are `equation`, `multline`, `gather`, `align`, `alignat`, `flalign` and `eqnarray`, and inside `$$…$$` also `aligned`, `alignedat`, `split` and `gathered` — the form most LLM-written Markdown uses. A multi-line environment stays **one** Word equation: `gather` and `multline` with their lines stacked, `align` and its relatives stacked *and* aligned on the `&` column, so the `=` signs line up the way LaTeX draws them. The alignment is written with OMML's own `<m:aln/>` marker — the one Word itself uses.

### What's supported

| Construct | Examples |
|---|---|
| Fractions | `\frac`, `\dfrac`, `\tfrac`, `\cfrac` |
| Roots | `\sqrt{x}`, `\sqrt[3]{x}` |
| Sub- and superscripts | `x^2`, `a_i`, `x_i^2`, primes `f'`, `f''` |
| Greek letters | `\alpha` … `\omega`, `\Gamma` … `\Omega`, variants `\varepsilon`, `\varphi`, `\vartheta`, `\varkappa`, … |
| Math alphabets | `\mathbb{R}`, `\mathcal{L}`, `\mathscr{F}`, `\mathfrak{g}`, `\mathsf`, `\mathtt`, `\mathbf`, `\mathit`, `\boldsymbol`, `\bm`, `\mathbfit`, `\mathrm` |
| Operators and relations | `\times`, `\pm`, `\cdot`, `\leq`, `\neq`, `\approx`, `\equiv`, `\sim`, `\cong`, `\propto`, `\in`, `\subseteq`, `\mid`, `\perp`, `\parallel`, `\forall`, `\exists`, `\implies`, `\iff`, `\mapsto`, `\hookrightarrow`, `\Longrightarrow`, `\top`, `\dagger`, `\therefore`, `\square`, `\checkmark`, … — over 300 symbols in all, amssymb included |
| Negation | `\not=`, `\not\in`, `\not\subset`, … — the ready-made negated character where Unicode has one |
| Function names | `\sin`, `\cos`, `\log`, `\ln`, `\exp`, `\det`, `\Pr`, `\gcd`, `\max`, `\min`, `\sup`, `\inf`, `\operatorname{…}`, `\operatorname*{argmax}_x`, `\mathop` — set upright; `\max`/`\min`/`\sup`/`\inf` take their subscript underneath, like `\lim` |
| Modulo | `\bmod`, `\pmod{n}`, `\mod`, `\pod` |
| Text inside a formula | `\text{}`, `\textbf{}`, `\textit{}`, `\emph{}`, `\mbox{}` — Cyrillic included, `$…$` inside `\text{}` too |
| N-ary operators | `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\iiiint`, `\oint`, `\oiint`, `\bigcup`, `\bigcap`, `\bigoplus`, `\bigotimes`, `\bigodot`, `\biguplus`, `\bigsqcup`, `\bigvee`, `\bigwedge` — with `_` and `^` limits, `\limits` / `\nolimits` |
| Limits | `\lim`, `\limsup`, `\liminf` with a lower limit: `\lim_{x \to 0}` |
| Delimiters | `\left( … \right)`, `\left\{ … \right\}`, `\left\| … \right\|`, `\lVert … \rVert`, `\langle`, `\lfloor`, `\lceil`, `\left.` / `\right.`, `\middle|`; fixed sizes `\big`, `\Big`, `\bigg`, `\Bigg` (`l`/`r`/`m`); `\ket`, `\bra`, `\braket` |
| Accents | `\hat`, `\widehat`, `\tilde`, `\widetilde`, `\bar`, `\vec`, `\dot`, `\ddot`, `\acute`, `\grave`, `\check`, `\breve`, `\overrightarrow`, `\overleftarrow` |
| Over and under | `\overline`, `\underline`, `\overset`, `\underset`, `\stackrel`, `\overbrace{…}^{n}`, `\underbrace{…}_{n}`, `\xrightarrow[below]{above}`, `\xleftarrow` and their relatives |
| Boxes and decorations | `\boxed`, `\fbox`, `\cancel`, `\bcancel`, `\xcancel`, `\phantom`, `\hphantom`, `\vphantom`, `\smash` |
| Colour | `\color{red}`, `\textcolor{blue}{…}`, `\color[HTML]{1F77B4}` — about 22 named colours plus HTML, rgb, RGB and gray models |
| Binomial coefficient | `\binom`, `\dbinom`, `\tbinom`, and the infix `{n \choose k}` |
| Infix fractions | `{a \over b}`, `{n \atop k}`, `\brace`, `\brack` — each splits the group it sits in |
| Matrices | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `smallmatrix`, `pmatrix*[r]`, `cases`, `dcases`, `rcases` |
| Formula tables | `\begin{array}{lcr} … \end{array}` — the `l`, `c`, `r` column alignment carries into Word |
| Stacked limits | `\substack{i < j \\ i \in S}`, `\begin{subarray}{l}` |
| Line break | `\\` in any formula, not just inside a matrix or amsmath: the lines stack |
| Line alignment | `&` between the lines of a multi-line formula: `a &= b \\ c &= d` puts the `=` signs under one another |
| Equation numbers | `\tag{n}` (`\tag*{n}` without parentheses), `$$ … $$ (n)`; `\label`, `\nonumber`, `\notag` are accepted |
| Spacing and escapes | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\hspace{…}`, `\kern`, `\mkern`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_`; `\displaystyle` and the other style switches are accepted |

### Limits of support

Of the 147 constructs an LLM most often writes, 146 convert; the one left is `\sideset` — OMML cannot attach scripts to both sides of a big operator, and the warning says so. The four cases below are refused **deliberately**. This is not a to-do list: each one has a reason why no correct behaviour exists, so the converter refuses honestly instead of producing something that looks close but is wrong.

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

**Page.** A4 by default, with 25.4 mm margins; the MCP server can switch to US Letter.

**Language.** The document language is detected from the text — Russian when at least 30 % of the letters are Cyrillic, English otherwise — or taken from `lang` in the front matter, so Word checks spelling and hyphenates in the right language.

**Colour.** Every run is black, headings included. The default python-docx template colours headings through the document theme, so the theme colours are cleared explicitly.

**Headings.** Sizes are scaled from the chosen body size. At 12 pt that gives 18 / 16 / 14 / 13 / 12 / 12 pt for levels 1–6; every level is bold, and level 6 is additionally italic so it stays distinct from level 5. Headings use the chosen font, which required stripping the *theme fonts* from the styles — in OOXML a theme attribute overrides an explicitly set font name.

**Alignment.** Paragraphs, list items, quotes and footnotes are justified. Headings, code blocks and table cells are not. Display formulas and figures are centred. A line that ends in a manual break is not stretched across a justified paragraph.

**Numbered equations.** `$$ … $$ (1)` or `\tag{1}` centres the formula and puts the number flush right, the way LaTeX and GOST lay it out.

**Tables.** Borders are written as direct formatting rather than left to the style, because some viewers ignore style-level borders and render the table without a grid. Column widths follow the length of their content, and the header row repeats on every page.

**GOST 7.32 preset.** A4; margins 30 mm left, 15 mm right, 20 mm top and bottom; 14 pt by default (a size you set yourself wins); 1.5 line spacing; a 1.25 cm first-line indent; headings bold at the body size; "Рисунок N — …" under figures and "Таблица N — …" above tables; the page number centred in the footer; "СОДЕРЖАНИЕ" as the title of the table of contents.

**Templates.** Through the MCP server (`template=`) a `.docx` can supply the styles, margins, headers and footers — like pandoc's `--reference-doc`. Its body text is not copied, and runs carry no direct font or size, so the template's styles decide.

---

## ↩️ Word → Markdown

The reverse direction walks the document in order, so tables stay where they were, and writes Markdown that the forward converter (or any CommonMark reader) reads back as the same text.

| In Word | In Markdown |
|---|---|
| Heading styles, Title, numbered headings | `#` … `######`, numbers kept |
| Bold, italic, strikethrough, monospace runs, super/subscript, underline, highlight | `**`, `*`, `~~`, `` ` ``, `<sup>`, `<sub>`, `<u>`, `==` — adjacent runs merged, markers never split by spaces |
| Hyperlinks, links to bookmarks | `[text](url)`, `<url>`, `[text](#heading-slug)` |
| Bulleted and numbered lists | Real list numbering read from Word: counters, restarts, nesting, task boxes `[ ]` / `[x]` |
| Code paragraphs | Fenced blocks with tabs and indentation kept |
| Quotes and shaded callouts | `>` and `> [!NOTE]` … |
| Tables | GFM tables in place, with formatting, `<br>` for multi-line cells and escaped `\|` |
| Pictures | Saved to `<name>_media/` next to the `.md` and linked (SVG preferred over its PNG fallback) |
| Equations | LaTeX: `$…$`, `$$…$$`, `\begin{aligned}` for aligned lines |
| Footnotes and endnotes | `[^1]` with the definitions at the end |
| Title, author, subject, keywords | YAML front matter |

Characters that mean something in Markdown — `*`, `_`, `$`, `#` at a line start, `|` in a table — are escaped, so the text reads back unchanged. What Markdown has no way to express — merged table cells, nested tables, text boxes, charts and shapes, embedded OLE objects, the table of contents — is reported once per kind with a count; page layout, headers and footers are dropped.

---

## 🏗️ Project structure

```
MDtoWORD/
├── 📦 mdtoword/                  # Application package (run: python -m mdtoword)
│   ├── __main__.py               # Entry point
│   ├── app.py                    # PyQt6 main window
│   ├── gui/                      # Drag-and-drop widgets, interface texts (RU/EN)
│   ├── converters.py             # Qt-free conversion core, used by the GUI and the MCP server
│   ├── errors.py                 # ConversionError and ConversionWarning (code + source line)
│   ├── options.py                # DocumentOptions: preset, page size, language, template, …
│   ├── md_extensions.py          # ^sup^, ==mark== and the front matter reader
│   ├── safe_fetch.py             # SSRF-hardened download of remote images
│   ├── gfm_renderer/             # Markdown → Word: one renderer class assembled from
│   │                             #   per-concern parts (blocks, inline, tables, images,
│   │                             #   equations, footnotes, HTML, document setup, TOC, …)
│   ├── latex_omml/               # LaTeX → OMML: tokenizer, parser, command handlers
│   │                             #   grouped by kind, environments, symbol tables
│   ├── ooxml/                    # Word primitives: numbering, footnotes, bookmarks,
│   │                             #   fields, page setup, images, SVG, shading
│   ├── docx_to_markdown/         # Word → Markdown: document walk, inline, lists,
│   │                             #   tables, quotes, media, notes, captions, escaping
│   ├── omml_latex/               # OMML → LaTeX for the Word → Markdown direction
│   ├── mcp_server/               # MCP server: tools, models, batch runner, guide
│   ├── agent_guide.md            # Markdown guide served to agents as an MCP resource
│   ├── workflow.py               # Source discovery and output path allocation
│   └── theme.py                  # Dark and light themes, persisted choice
├── 📁 tests/                     # Test suite (unittest)
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
│   └── 📁 releases/              # Release notes
├── 📄 MDtoWORD.spec              # PyInstaller configuration (macOS)
├── 📄 pyproject.toml             # Package metadata and the mdtoword-mcp entry point
├── 📋 requirements.txt           # Application dependencies
├── 📋 requirements-core.txt      # Shared conversion core (GUI + MCP server)
├── 📋 requirements-build.txt     # Build dependencies (PyInstaller)
├── 📋 requirements-mcp.txt       # MCP server dependencies
├── 📋 environment.yml            # Conda environment (Python 3.12)
└── 📖 README.md                  # Documentation (this file)
```

---

## ⚙️ Technical details

### Default formatting
- **Page**: A4, 25.4 mm margins
- **Font**: Times New Roman, 12 pt (14 pt with the GOST preset)
- **Colour**: black (RGB 0, 0, 0)
- **Code**: Courier New, 0.85 × the body size in blocks, 0.9 × the surrounding text inline
- **Tables**: borders as direct formatting, bold header row repeated on every page

### Supported formats
- **Markdown → Word**: input `.md`, `.markdown` → output `.docx`
- **Word → Markdown**: input `.docx` → output `.md`

---

## 🤖 MCP server

MDtoWORD ships an MCP server so agents can run the same conversions the GUI does — and check their Markdown and the resulting document along the way.

Install it as a package with the `mcp` extra (no PyQt6, so the server stays runnable headless):

```bash
python -m pip install "/path/to/MDtoWord[mcp]"     # or, from a clone: pip install ".[mcp]"
```

This puts an `mdtoword-mcp` command on the `PATH` of that environment. Register it with any MCP client:

```json
{
  "mcpServers": {
    "mdtoword": {
      "command": "/path/to/venv/bin/mdtoword-mcp"
    }
  }
}
```

For Claude Code:

```bash
claude mcp add mdtoword --scope user -- /path/to/venv/bin/mdtoword-mcp
```

Running straight from a checkout still works — `python -m pip install -r requirements-mcp.txt`, then `"command": "/path/to/venv/bin/python", "args": ["-m", "mdtoword.mcp_server"], "cwd": "/path/to/MDtoWord"` (or `PYTHONPATH=/path/to/MDtoWord` instead of `cwd`).

### Tools

| Tool | What it does |
| --- | --- |
| `markdown_to_word` | Converts `.md` / `.markdown` files and directories to `.docx`: everything in [What Markdown is understood](#-what-markdown-is-understood), LaTeX → native equations, the GOST preset, templates, a table of contents. |
| `word_to_markdown` | Converts `.docx` files and directories to Markdown — headings, lists, tables, links, images, footnotes and equations. |
| `preview_markdown` | Renders Markdown in memory and reports only what would not survive the conversion, with source lines. Writes nothing. |
| `check_latex` | Checks formulas one by one: will each become a native Word equation? Writes nothing. |
| `inspect_docx` | Describes a `.docx`: page setup, outline, counts of tables, images, equations, footnotes, lists, links, properties and the start of the text. Writes nothing. |

Plus a resource, `mdtoword://guide/markdown` — a compact guide to writing Markdown that converts cleanly — and a prompt, `prepare_markdown_for_word`, that walks an agent through `check_latex` → `preview_markdown` → `markdown_to_word` → `inspect_docx`.

The tools are annotated for clients that auto-approve by policy: `preview_markdown`, `check_latex` and `inspect_docx` are read-only; the two converters are marked destructive because they overwrite existing outputs. Conversions run file by file in a worker thread, with a progress notification after each file.

All converting tools take paths, never file contents, and accept files and directories mixed together; directories are scanned recursively. Where they write, they overwrite an existing output file without warning.

`markdown_to_word` and `preview_markdown` do not fetch images referenced by an `http(s)` URL by default — such an image becomes its alt text plus a warning. Pass `fetch_remote_images=true` to enable fetching, and only for Markdown from a source you trust. Even then the download refuses any address that is not publicly routable (loopback, private ranges, link-local such as `169.254.169.254`, CGNAT), connects to the address it checked rather than resolving the name again, re-checks every redirect, and stops reading past 20 MB. The GUI keeps fetching remote images as before, through the same safeguards.

`markdown_to_word` and `preview_markdown` also only read local-filesystem images from within the paths passed in `inputs`: a directory input allows images anywhere under it, a file input allows images only next to it, not in sibling directories. An image outside these roots becomes its alt text plus a warning, same as a missing file. Pass `image_root` to widen the allowed root when your images live elsewhere. The GUI has no such restriction — a human already chose the file there.

### Parameters

`markdown_to_word`:

| Parameter | Default | Meaning |
| --- | --- | --- |
| `inputs: list[str]` | required | Files and/or directories, mixed together; directories are scanned recursively for `.md` / `.markdown` files. |
| `output_dir: str \| None` | `None` | Where outputs go. `None` writes each output next to its source file. |
| `font_name: str` | `"Times New Roman"` | Body font of the produced document. |
| `font_size: float \| None` | `None` | Body size in points; headings scale from it. `None` means 12, or 14 with `preset="gost"`. |
| `preset` | `"default"` | `"gost"` lays the document out by GOST 7.32 (see [formatting](#-how-the-word-document-is-formatted)). |
| `page_size` | `None` | `"A4"` or `"Letter"`; `None` is A4 (or the template's size). |
| `language: str` | `"auto"` | Document language for spelling and hyphenation, e.g. `"ru-RU"`; `auto` detects it from the text or the front matter. |
| `line_breaks` | `"soft"` | `"soft"`: a single newline inside a paragraph is a space (CommonMark). `"preserve"`: it is a line break. |
| `template: str \| None` | `None` | A reference `.docx` whose styles, margins, headers and footers are reused. Must exist. |
| `toc: bool` | `False` | Put a table of contents at the start (`[TOC]` in the text works regardless). |
| `footnotes` | `"native"` | `"native"`: Word footnotes at the bottom of the page. `"section"`: a numbered section at the end, titled `footnotes_heading`. |
| `footnotes_heading: str` | `"Footnotes"` | Title of that section (`"section"` mode only; Russian documents get «Сноски» automatically). |
| `fetch_remote_images: bool` | `False` | Allow fetching images referenced by an `http(s)` URL. |
| `image_root: str \| None` | `None` | Widen the directory local images may be read from. Defaults to a root derived from `inputs` (see above). |

`preview_markdown` takes the same parameters **except `output_dir`** — it never writes a file.

`word_to_markdown` takes `inputs` (required), `output_dir` and `extract_media` (default `true`: pictures are saved next to the output, into `<name>_media/`, and linked from the Markdown).

`check_latex` takes `formulas: list[str]` — raw LaTeX, with or without `$…$`, `$$…$$`, `\(…\)` or `\[…\]` around it; amsmath environments are accepted too.

`inspect_docx` takes `path: str`.

### What the tools return

`markdown_to_word` and `word_to_markdown` return:

```
sources_found: int
converted: [{ source, output, warnings: [{ message, code, line }] }]
failed:    [{ source, error }]
```

`preview_markdown` returns the same shape with `previews: [{ source, warnings: [...] }]` in place of `converted`, and no `output` field. `check_latex` returns `{ results: [{ formula, ok, error }], all_ok }`.

A few things worth knowing before reading these:

- **`sources_found` is the field to check first.** It is the count of supported files the `inputs` resolved to. `0` means the paths matched nothing — that is a signal to check the paths, not to conclude there was nothing to do.
- A failing file does not stop the batch: it lands in `failed` while the rest convert normally. Read both lists.
- `warnings` are non-fatal: the output file was still written. Each one has a stable `code` (`formula_unsupported`, `image_not_found`, `image_outside_root`, `image_remote_disabled`, `image_fetch_failed`, `image_too_large`, `html_dropped`, `link_anchor_missing`, `footnote_unreferenced`, `math_prose`, …) and, when known, the 1-based `line` of the Markdown source — fix the source at that line and convert again.

> **Changed in 1.2:** `warnings` used to be plain strings. Clients that read them as strings should read `message` now.

### Examples

Paths should be absolute. A relative path resolves against the **server's** working directory, not the agent's project directory.

1. Convert a whole folder: `markdown_to_word(inputs=["/abs/path/docs"])`.
2. Check a document before converting: `preview_markdown(inputs=["/abs/path/README.md"])` — nothing is written, you get warnings with line numbers.
3. A report by GOST with a table of contents, into a build directory: `markdown_to_word(inputs=["/abs/path/report.md"], output_dir="/abs/path/build", preset="gost", toc=true)`.
4. Use the company template: `markdown_to_word(inputs=["/abs/path/memo.md"], template="/abs/path/company.docx")`.
5. Make sure the formulas will be equations: `check_latex(formulas=["\\frac{a}{b}", "\\mathbb{R}^n"])`.
6. Verify the result: `inspect_docx(path="/abs/path/build/report.docx")` — outline, page size, how many tables, images and equations made it.
7. Pull a Word file back into Markdown with its pictures: `word_to_markdown(inputs=["/abs/path/report.docx"])`.

### Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| Server fails to start: `ModuleNotFoundError: No module named 'mdtoword'` | You run from a checkout without `cwd`/`PYTHONPATH`. Install the package (`pip install ".[mcp]"`) and use `mdtoword-mcp`, or set `cwd` to the repository. |
| `sources_found: 0` and empty lists | The paths matched no supported file: a wrong path, or a directory with no `.md`/`.markdown` files (or `.docx`, for `word_to_markdown`). |
| A warning says an image was not fetched | Remote fetching is off by default. Pass `fetch_remote_images=true`, and only for Markdown you trust. |
| A warning says an image could not be fetched … not publicly routable | The URL (or a redirect) points at a local or private address; such fetches are always refused. |
| A warning says an image is outside the allowed root | The image lives outside the paths passed in `inputs`. Pass `image_root` to widen it. |
| `template must be an existing .docx file` | The template path is wrong or relative to the server's directory. Pass an absolute path. |
| Output written somewhere unexpected | A relative `output_dir` resolved against the server's working directory. Pass an absolute path. |
| `tests/test_mcp_server.py` skips entirely | The `mcp` SDK is not installed in the interpreter running the tests. Install `requirements-mcp.txt`. |

---

## 🛠️ Development

Run the tests from the project root:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests -p "test_*.py"
```

If the run ends in a segfault or a crash trace instead of a pass count, the
interpreter's Qt build is at fault — rerun with the project's own virtualenv
rather than a system or anaconda Python.

`QT_QPA_PLATFORM=offscreen` lets the interface tests run without a display. The suite renders documents and checks them on the XML level; for layout, `soffice --headless --convert-to pdf` on a result shows numbering, footnotes and equations the way a word processor draws them.

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
Check the warning in the result dialog: it names the exact construct and the line, e.g. `report.md:12: Formula kept as text: … (Unsupported LaTeX command: \qedsymbol)`. The formula itself is preserved verbatim in the document — rewrite it using a supported construct from [the table above](#whats-supported) and convert again.

**A dollar sign in the text came out oddly**
Write a literal dollar as `\$`. If the converter sees `$…$` wrapped around ordinary words it leaves the text alone and warns you — but escaping it up front is better.

**Lines I broke in the source were joined**
That is CommonMark: a single newline inside a paragraph is a space. End a line with two spaces or `\` for a line break, or use `line_breaks="preserve"` through the MCP server.

**An image didn't make it into the document**
Local paths resolve relative to the `.md` file (Cyrillic and spaces in file names are fine), and `http(s)` URLs are downloaded with a 10-second timeout from public addresses only. If the file is missing, the format is unknown or the network is unavailable, `[alt text]` appears in its place and the dialog carries a warning with the address.

**A large batch makes the window sluggish**
Conversion runs on the interface thread, so long queues leave the window slow to respond. The progress bar still updates — let it finish.

**Word asks to update fields when opening the document**
The document has a table of contents; letting Word update fields adds the page numbers to it.

**Word → Markdown didn't keep something**
Text boxes, comments and merged table cells have no Markdown counterpart; the result dialog lists what was simplified.

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

MDtoWORD берёт ваши `.md`-файлы — один, десяток или целую папку — и складывает рядом готовые `.docx`. Разметка не «приблизительно похожа», а переносится по-настоящему: заголовки становятся стилями Word с закладками, на которые можно ссылаться, списки получают настоящую нумерацию Word, сноски — настоящие сноски внизу страницы, таблицы сохраняют форматирование, а формула вроде `$E = mc^2$` превращается в **родное редактируемое уравнение Word**, а не в картинку и не в голый текст. Пресет ГОСТ 7.32 оформляет документ так, как ждут отчёты и диссертации.

Работает и в обратную сторону: из `.docx` получается Markdown с заголовками, списками, таблицами, ссылками, изображениями, сносками и формулами на своих местах — см. [Word → Markdown](#-word--markdown-1).

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
- **mcp 1.28.1** — только для [MCP-сервера](#-mcp-сервер), ставится отдельно

```bash
pip install -r requirements.txt
```

### Conda (виртуальное окружение)

```bash
# Создать окружение из environment.yml (Python 3.12)
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

**Оформление.** Выпадающий список шрифтов (Arial, Times New Roman, Calibri, Georgia, Helvetica, Courier New) и поле размера от 6 до 72 pt задают базовое оформление документа. **Стиль** переключает *Обычный* и *ГОСТ 7.32* — A4, поля 30/15/20/20 мм, полуторный интервал, абзацный отступ 1,25 см и 14 pt (размер следует за стилем, если вы не меняли его сами). Флажок **«Оглавление»** вставляет оглавление в начало документа.

**Темы и язык.** В нижней строке — круглая кнопка темы (☀ / ☾), переключающая тёмное и светлое оформление. Выбор запоминается через `QSettings` и восстанавливается при следующем запуске. Соседняя кнопка **EN / RU** переключает язык интерфейса.

По ходу конвертации показывается прогресс-бар и имя текущего файла, а в конце — диалог с числом успешных файлов, ошибками и предупреждениями. Предупреждение называет строку исходника: `отчёт.md:42: Image not found: chart.png`.

---

## 📝 Что понимается в Markdown

Разметка разбирается парсером CommonMark + GitHub Flavored Markdown с несколькими распространёнными расширениями.

| Элемент | Markdown | Результат в Word |
|---|---|---|
| **Заголовки** | `# … ######` | Стили Heading 1–6 с иерархией размеров, у каждого — закладка |
| **Абзацы** | обычный текст | Стиль Normal, выравнивание по ширине. Одиночный перенос строки внутри абзаца — пробел, как в CommonMark |
| **Разрыв строки** | два пробела в конце строки или `\` | Разрыв строки |
| **Курсив / жирный / зачёркнутый** | `*текст*`, `**текст**`, `~~текст~~` | Italic / Bold / Strikethrough |
| **Индексы, выделение** | `H~2~O`, `x^2^`, `==текст==` | Нижний и верхний индекс, жёлтое выделение |
| **Код в строке** | `` `код` `` | Courier New на светло-сером фоне, размер под окружающий текст |
| **Ссылки** | `[текст](url)`, `<url>`, голый `https://…` | Настоящая гиперссылка, чёрная с подчёркиванием |
| **Ссылки на заголовки** | `[см.](#якорь-заголовка)` | Внутренняя ссылка на закладку заголовка (правила якорей GitHub, кириллица тоже) |
| **Изображения** | `![alt](путь "заголовок")` | Картинка, уменьшенная до ширины текста. PNG, JPEG, GIF, BMP, TIFF, WebP и SVG. Локальные пути считаются от файла-исходника, ссылки `http(s)` скачиваются (поведение GUI — [MCP-сервер](#-mcp-сервер) по умолчанию их не загружает) |
| **Рисунки** | картинка одна в своём абзаце | Картинка по центру с нумерованной подписью из заголовка или alt-текста: «Рисунок 1 — …», в английском документе «Figure 1: …» |
| **Цитаты** | `> текст`, `> > вложенная` | Отступ по уровню и линия слева; списки и код внутри сохраняют отступ цитаты |
| **Плашки GitHub** | `> [!NOTE]`, `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]`, `[!CAUTION]` | Затенённая врезка с жирным заголовком («Примечание», «Совет», «Важно», «Внимание», «Осторожно») |
| **Разделители** | `---` | Горизонтальная линия |
| **Блоки кода** | ````` ```python ````` или отступ в четыре пробела | Стиль «Source Code»: моноширинный, серый фон, тонкая рамка, пробелы сохраняются |
| **Списки** | `-`, `1.`, `5.` | Настоящая нумерация Word: каждый список начинается заново, `5.` начинается с 5, вложенность до девяти уровней, пункты из нескольких абзацев и с кодом |
| **Чек-листы** | `- [ ]`, `- [x]` | Символы ☐ и ☒ в начале пункта |
| **Списки определений** | `Термин`, затем `: определение` | Жирный термин, определение с отступом |
| **Таблицы** | `\| … \|` | Таблица с границами: жирная шапка повторяется на каждой странице, выравнивание колонок из разметки, форматирование, ссылки и формулы в ячейках |
| **Подписи таблиц** | `Таблица: …`, `Table: …` или `: …` вплотную до или после таблицы | Нумерованная подпись над таблицей |
| **Сноски** | `[^1]` и `[^1]: …` | Настоящие сноски Word внизу страницы |
| **Front matter** | блок YAML между `---` в самом начале | Титульный блок (title, subtitle, author, date) и свойства документа (название, автор, тема, ключевые слова, язык) |
| **Оглавление** | `[TOC]` отдельной строкой | Поле оглавления Word, сразу заполненное ссылками на заголовки; номера страниц Word добавит при обновлении полей |
| **Формулы** | `$…$`, `$$…$$`, amsmath | Родные уравнения Word — см. [раздел ниже](#-формулы-latex) |
| **HTML** | `<br>`, `<sub>`, `<sup>`, `<kbd>`, `<mark>`, `<u>`, `<del>`, `<b>`, `<i>`, `<code>`, `<a href>`, `<img>`, `<details>` | Соответствующее форматирование. Комментарии отбрасываются; любой другой тег удаляется с предупреждением, его текст остаётся |

**Ничего не теряется молча.** Всё, что нельзя передать точно, — отсутствующее изображение, неподдерживаемая формула, незнакомый HTML-тег, ссылка на несуществующий заголовок, сноска, на которую никто не ссылается, — попадает в итоговый диалог вместе с номером строки исходного Markdown.

---

## 🧮 Формулы LaTeX

Главная возможность этой версии. Формулу можно записать четырьмя способами:

```markdown
Внутри строки: $E = mc^2$

Отдельным блоком:
$$\int_0^1 x^2\,dx = \frac{1}{3}$$

С номером уравнения:
$$a^2 + b^2 = c^2$$ (1)
$$E = mc^2 \tag{2}$$

В окружении amsmath:
\begin{align}
  f(x) &= ax + b \\
  g(x) &= cx + d
\end{align}
```

Все четыре превращаются в **настоящие уравнения Word (OMML)** — они открываются в редакторе формул, их можно править, они масштабируются вместе с текстом. Это не картинка и не текстовая имитация.

Из окружений amsmath поддерживаются `equation`, `multline`, `gather`, `align`, `alignat`, `flalign` и `eqnarray`, а внутри `$$…$$` ещё `aligned`, `alignedat`, `split` и `gathered` — именно так пишут формулы LLM. Многострочное окружение остаётся **одним** уравнением Word: `gather` и `multline` — со сложенными в столбик строками, `align` и родственные — ещё и с выравниванием по `&`, то есть знаки `=` встают друг под другом, как в LaTeX. Выравнивание записывается штатным маркером OMML `<m:aln/>` — тем же, которым пользуется сам Word.

### Что поддерживается

| Конструкция | Примеры |
|---|---|
| Дроби | `\frac`, `\dfrac`, `\tfrac`, `\cfrac` |
| Корни | `\sqrt{x}`, `\sqrt[3]{x}` |
| Индексы и степени | `x^2`, `a_i`, `x_i^2`, штрихи `f'`, `f''` |
| Греческие буквы | `\alpha` … `\omega`, `\Gamma` … `\Omega`, варианты `\varepsilon`, `\varphi`, `\vartheta`, `\varkappa`, … |
| Математические алфавиты | `\mathbb{R}`, `\mathcal{L}`, `\mathscr{F}`, `\mathfrak{g}`, `\mathsf`, `\mathtt`, `\mathbf`, `\mathit`, `\boldsymbol`, `\bm`, `\mathbfit`, `\mathrm` |
| Операции и отношения | `\times`, `\pm`, `\cdot`, `\leq`, `\neq`, `\approx`, `\equiv`, `\sim`, `\cong`, `\propto`, `\in`, `\subseteq`, `\mid`, `\perp`, `\parallel`, `\forall`, `\exists`, `\implies`, `\iff`, `\mapsto`, `\hookrightarrow`, `\Longrightarrow`, `\top`, `\dagger`, `\therefore`, `\square`, `\checkmark`, … — всего больше 300 символов, включая amssymb |
| Отрицание | `\not=`, `\not\in`, `\not\subset`, … — готовый перечёркнутый символ, если он есть в Unicode |
| Имена функций | `\sin`, `\cos`, `\log`, `\ln`, `\exp`, `\det`, `\Pr`, `\gcd`, `\max`, `\min`, `\sup`, `\inf`, `\operatorname{…}`, `\operatorname*{argmax}_x`, `\mathop` — прямым шрифтом; у `\max`/`\min`/`\sup`/`\inf` нижний индекс уходит под имя, как у `\lim` |
| Сравнение по модулю | `\bmod`, `\pmod{n}`, `\mod`, `\pod` |
| Текст внутри формулы | `\text{}`, `\textbf{}`, `\textit{}`, `\emph{}`, `\mbox{}` — включая кириллицу, и `$…$` внутри `\text{}` |
| N-арные операторы | `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\iiiint`, `\oint`, `\oiint`, `\bigcup`, `\bigcap`, `\bigoplus`, `\bigotimes`, `\bigodot`, `\biguplus`, `\bigsqcup`, `\bigvee`, `\bigwedge` — с пределами через `_` и `^`, `\limits` / `\nolimits` |
| Пределы | `\lim`, `\limsup`, `\liminf` с нижним пределом: `\lim_{x \to 0}` |
| Скобки | `\left( … \right)`, `\left\{ … \right\}`, `\left\| … \right\|`, `\lVert … \rVert`, `\langle`, `\lfloor`, `\lceil`, `\left.` / `\right.`, `\middle|`; фиксированные размеры `\big`, `\Big`, `\bigg`, `\Bigg` (`l`/`r`/`m`); `\ket`, `\bra`, `\braket` |
| Акценты | `\hat`, `\widehat`, `\tilde`, `\widetilde`, `\bar`, `\vec`, `\dot`, `\ddot`, `\acute`, `\grave`, `\check`, `\breve`, `\overrightarrow`, `\overleftarrow` |
| Над и под | `\overline`, `\underline`, `\overset`, `\underset`, `\stackrel`, `\overbrace{…}^{n}`, `\underbrace{…}_{n}`, `\xrightarrow[снизу]{сверху}`, `\xleftarrow` и родственные |
| Рамки и зачёркивания | `\boxed`, `\fbox`, `\cancel`, `\bcancel`, `\xcancel`, `\phantom`, `\hphantom`, `\vphantom`, `\smash` |
| Цвет | `\color{red}`, `\textcolor{blue}{…}`, `\color[HTML]{1F77B4}` — около 22 именованных цветов и модели HTML, rgb, RGB, gray |
| Биномиальный коэффициент | `\binom`, `\dbinom`, `\tbinom` и инфиксная запись `{n \choose k}` |
| Инфиксные дроби | `{a \over b}`, `{n \atop k}`, `\brace`, `\brack` — каждая делит группу, в которой стоит |
| Матрицы | `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `smallmatrix`, `pmatrix*[r]`, `cases`, `dcases`, `rcases` |
| Таблицы формул | `\begin{array}{lcr} … \end{array}` — выравнивание колонок `l`, `c`, `r` переносится в Word |
| Многоэтажные пределы | `\substack{i < j \\ i \in S}`, `\begin{subarray}{l}` |
| Перенос строки | `\\` в любой формуле, а не только в матрице или окружении amsmath: строки складываются в столбик |
| Выравнивание строк | `&` между строками многострочной формулы: `a &= b \\ c &= d` ставит знаки `=` друг под другом |
| Номера формул | `\tag{n}` (`\tag*{n}` — без скобок), `$$ … $$ (n)`; `\label`, `\nonumber`, `\notag` принимаются |
| Пробелы и экранирование | `\,`, `\;`, `\:`, `\!`, `\quad`, `\qquad`, `\hspace{…}`, `\kern`, `\mkern`, `\{`, `\}`, `\%`, `\$`, `\&`, `\#`, `\_`; `\displaystyle` и другие переключатели стиля принимаются |

### Границы поддержки

Из 147 конструкций, которые чаще всего пишут LLM, конвертируются 146; осталась только `\sideset` — OMML не умеет вешать индексы с обеих сторон большого оператора, и предупреждение так и говорит. Четыре случая ниже отвергаются **намеренно**. Это не список «доделать позже»: у каждого есть причина, по которой правильного поведения просто не существует, — поэтому конвертер честно отказывается вместо того, чтобы выдать похожий, но неверный результат.

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

**Страница.** По умолчанию A4 с полями 25,4 мм; через MCP-сервер можно выбрать US Letter.

**Язык.** Язык документа определяется по тексту — русский, если кириллицы хотя бы 30 % букв, иначе английский — или берётся из `lang` в front matter, поэтому Word проверяет орфографию и расставляет переносы для правильного языка.

**Цвет.** Весь текст чёрный, включая заголовки. Шаблон python-docx по умолчанию красит заголовки через тему документа, поэтому цвета темы вычищаются принудительно.

**Заголовки.** Размеры считаются от выбранного размера текста. При 12 pt получается 18 / 16 / 14 / 13 / 12 / 12 pt для уровней 1–6; все уровни жирные, шестой дополнительно курсивный, чтобы отличаться от пятого. Заголовки используют выбранный шрифт — для этого из стилей вычищаются *шрифты темы*, потому что в OOXML атрибут темы перекрывает явно заданное имя шрифта.

**Выравнивание.** Абзацы, элементы списков, цитаты и сноски выровнены по ширине. Заголовки, блоки кода и ячейки таблиц — нет. Блочные формулы и рисунки центрируются. Строка, оканчивающаяся ручным разрывом, не растягивается по ширине абзаца.

**Нумерованные формулы.** `$$ … $$ (1)` или `\tag{1}` ставит формулу по центру, а номер — к правому краю, как в LaTeX и по ГОСТ.

**Таблицы.** Границы записываются прямым форматированием, а не только стилем: часть просмотрщиков игнорирует границы уровня стиля, и таблица приезжает без сетки. Ширина колонок следует за длиной содержимого, шапка повторяется на каждой странице.

**Пресет ГОСТ 7.32.** A4; поля: левое 30 мм, правое 15 мм, верхнее и нижнее 20 мм; 14 pt по умолчанию (размер, заданный вами, важнее); полуторный интервал; абзацный отступ 1,25 см; заголовки жирные, размером с основной текст; «Рисунок N — …» под рисунками и «Таблица N — …» над таблицами; номер страницы по центру внизу; «СОДЕРЖАНИЕ» — заголовок оглавления.

**Шаблоны.** Через MCP-сервер (`template=`) `.docx` может задать стили, поля и колонтитулы — как `--reference-doc` у pandoc. Его текст не копируется, а прогоны не несут прямого шрифта и размера, так что решают стили шаблона.

---

## ↩️ Word → Markdown

Обратное направление обходит документ по порядку, поэтому таблицы остаются на своих местах, и пишет Markdown, который прямой конвертер (и любой читатель CommonMark) прочитает обратно в тот же текст.

| В Word | В Markdown |
|---|---|
| Стили заголовков, Title, нумерованные заголовки | `#` … `######`, номера сохраняются |
| Жирный, курсив, зачёркнутый, моноширинный текст, индексы, подчёркивание, выделение | `**`, `*`, `~~`, `` ` ``, `<sup>`, `<sub>`, `<u>`, `==` — соседние прогоны склеиваются, маркеры не разрываются пробелами |
| Гиперссылки, ссылки на закладки | `[текст](url)`, `<url>`, `[текст](#якорь-заголовка)` |
| Маркированные и нумерованные списки | Настоящая нумерация из Word: счётчики, перезапуски, вложенность, флажки `[ ]` / `[x]` |
| Абзацы с кодом | Блоки в ``` с сохранением табуляции и отступов |
| Цитаты и затенённые врезки | `>` и `> [!NOTE]` … |
| Таблицы | Таблицы GFM на своём месте, с форматированием, `<br>` в многострочных ячейках и экранированным `\|` |
| Изображения | Сохраняются в `<имя>_media/` рядом с `.md` и подключаются ссылкой (SVG — вместо его PNG-заглушки) |
| Формулы | LaTeX: `$…$`, `$$…$$`, `\begin{aligned}` для выровненных строк |
| Сноски и концевые сноски | `[^1]` с определениями в конце |
| Название, автор, тема, ключевые слова | YAML front matter |

Символы, значимые для Markdown, — `*`, `_`, `$`, `#` в начале строки, `|` в таблице — экранируются, так что текст читается обратно без изменений. То, что в Markdown выразить нельзя, — объединённые ячейки, вложенные таблицы, надписи, диаграммы и фигуры, внедрённые OLE-объекты, оглавление — попадает в предупреждения по одному на вид, с количеством; разметка страниц и колонтитулы отбрасываются.

---

## 🏗️ Структура проекта

```
MDtoWORD/
├── 📦 mdtoword/                  # Пакет приложения (запуск: python -m mdtoword)
│   ├── __main__.py               # Точка входа
│   ├── app.py                    # Главное окно на PyQt6
│   ├── gui/                      # Виджеты перетаскивания, тексты интерфейса (RU/EN)
│   ├── converters.py             # Ядро конвертации без Qt, общее для GUI и MCP-сервера
│   ├── errors.py                 # ConversionError и ConversionWarning (код + строка исходника)
│   ├── options.py                # DocumentOptions: пресет, формат страницы, язык, шаблон, …
│   ├── md_extensions.py          # ^sup^, ==mark== и разбор front matter
│   ├── safe_fetch.py             # Защищённая от SSRF загрузка удалённых изображений
│   ├── gfm_renderer/             # Markdown → Word: класс рендерера, собранный из частей
│   │                             #   по задачам (блоки, строчный текст, таблицы, изображения,
│   │                             #   формулы, сноски, HTML, настройка документа, оглавление, …)
│   ├── latex_omml/               # LaTeX → OMML: токенизатор, парсер, обработчики команд
│   │                             #   по видам, окружения, таблицы символов
│   ├── ooxml/                    # Примитивы Word: нумерация, сноски, закладки,
│   │                             #   поля, параметры страницы, изображения, SVG, заливки
│   ├── docx_to_markdown/         # Word → Markdown: обход документа, строчный текст, списки,
│   │                             #   таблицы, цитаты, медиа, сноски, подписи, экранирование
│   ├── omml_latex/               # OMML → LaTeX для направления Word → Markdown
│   ├── mcp_server/               # MCP-сервер: инструменты, модели, пакетный запуск, гайд
│   ├── agent_guide.md            # Руководство по Markdown для агентов (ресурс MCP)
│   ├── workflow.py               # Поиск исходников и раскладка результатов
│   └── theme.py                  # Тёмная и светлая темы, сохранение выбора
├── 📁 tests/                     # Тесты (unittest)
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
│   └── 📁 releases/              # Заметки к выпускам
├── 📄 MDtoWORD.spec              # Конфигурация PyInstaller (macOS)
├── 📄 pyproject.toml             # Метаданные пакета и команда mdtoword-mcp
├── 📋 requirements.txt           # Зависимости приложения
├── 📋 requirements-core.txt      # Общее ядро конвертации (GUI + MCP-сервер)
├── 📋 requirements-build.txt     # Зависимости сборки (PyInstaller)
├── 📋 requirements-mcp.txt       # Зависимости MCP-сервера
├── 📋 environment.yml            # Conda-окружение (Python 3.12)
└── 📖 README.md                  # Документация (этот файл)
```

---

## ⚙️ Технические детали

### Оформление по умолчанию
- **Страница**: A4, поля 25,4 мм
- **Шрифт**: Times New Roman, 12 pt (14 pt с пресетом ГОСТ)
- **Цвет**: чёрный (RGB 0, 0, 0)
- **Код**: Courier New, в блоках 0,85 от размера текста, в строке 0,9 от окружающего
- **Таблицы**: границы прямым форматированием, жирная шапка повторяется на каждой странице

### Поддерживаемые форматы
- **Markdown → Word**: вход `.md`, `.markdown` → выход `.docx`
- **Word → Markdown**: вход `.docx` → выход `.md`

---

## 🤖 MCP-сервер

MDtoWORD включает MCP-сервер, чтобы агенты могли выполнять те же конвертации, что и графический интерфейс, а заодно проверять свой Markdown и получившийся документ.

Установите его как пакет с дополнением `mcp` (без PyQt6 — сервер остаётся запускаемым без дисплея):

```bash
python -m pip install "/path/to/MDtoWord[mcp]"     # или из клона: pip install ".[mcp]"
```

В окружении появится команда `mdtoword-mcp`. Подключите её в любом MCP-клиенте:

```json
{
  "mcpServers": {
    "mdtoword": {
      "command": "/path/to/venv/bin/mdtoword-mcp"
    }
  }
}
```

Для Claude Code:

```bash
claude mcp add mdtoword --scope user -- /path/to/venv/bin/mdtoword-mcp
```

Запуск прямо из клона по-прежнему работает: `python -m pip install -r requirements-mcp.txt`, затем `"command": "/path/to/venv/bin/python", "args": ["-m", "mdtoword.mcp_server"], "cwd": "/path/to/MDtoWord"` (или `PYTHONPATH=/path/to/MDtoWord` вместо `cwd`).

### Инструменты

| Инструмент | Что делает |
| --- | --- |
| `markdown_to_word` | Конвертирует файлы и папки `.md` / `.markdown` в `.docx`: всё из раздела [Что понимается в Markdown](#-что-понимается-в-markdown), LaTeX → родные уравнения, пресет ГОСТ, шаблоны, оглавление. |
| `word_to_markdown` | Конвертирует файлы и папки `.docx` в Markdown — заголовки, списки, таблицы, ссылки, изображения, сноски и формулы. |
| `preview_markdown` | Рендерит Markdown в памяти и сообщает только о том, что не переживёт конвертацию, с номерами строк. Ничего не записывает. |
| `check_latex` | Проверяет формулы по одной: станет ли каждая родным уравнением Word. Ничего не записывает. |
| `inspect_docx` | Описывает `.docx`: параметры страницы, структуру заголовков, число таблиц, изображений, формул, сносок, пунктов списков и ссылок, свойства документа и начало текста. Ничего не записывает. |

Плюс ресурс `mdtoword://guide/markdown` — краткое руководство по Markdown, который конвертируется без потерь, — и prompt `prepare_markdown_for_word`, который ведёт агента по цепочке `check_latex` → `preview_markdown` → `markdown_to_word` → `inspect_docx`.

Инструменты размечены аннотациями для клиентов с автоодобрением по политике: `preview_markdown`, `check_latex` и `inspect_docx` — только чтение; оба конвертера помечены как разрушающие, потому что перезаписывают существующие результаты. Конвертация идёт пофайлово в рабочем потоке, после каждого файла клиент получает уведомление о прогрессе.

Конвертирующие инструменты принимают пути, а не содержимое файлов, и работают с файлами и папками вперемешку; папки просматриваются рекурсивно. Там, где они пишут, существующий выходной файл перезаписывается без предупреждения.

По умолчанию `markdown_to_word` и `preview_markdown` не загружают изображения по `http(s)`-ссылке — такое изображение превращается в альтернативный текст с предупреждением. Чтобы включить загрузку, передайте `fetch_remote_images=true`, и только для Markdown из источника, которому доверяете. Но и тогда загрузка отказывается идти на адреса, недоступные из интернета (loopback, частные сети, link-local вроде `169.254.169.254`, CGNAT), подключается именно к проверенному адресу, а не резолвит имя повторно, проверяет каждое перенаправление и перестаёт читать после 20 МБ. Графический интерфейс, как и раньше, загружает удалённые изображения — через те же защиты.

`markdown_to_word` и `preview_markdown` также читают локальные изображения только из путей, переданных в `inputs`: если это папка — разрешено любое изображение внутри неё, если отдельный файл — только рядом с ним, но не в соседних папках. Изображение за пределами этой границы превращается в альтернативный текст с предупреждением — точно как отсутствующий файл. Если изображения лежат вне `inputs`, передайте `image_root`, чтобы расширить разрешённый корень. На графический интерфейс это ограничение не распространяется — там файл уже выбрал человек.

### Параметры

`markdown_to_word`:

| Параметр | По умолчанию | Значение |
| --- | --- | --- |
| `inputs: list[str]` | обязателен | Файлы и/или папки вперемешку; папки просматриваются рекурсивно на файлы `.md` / `.markdown`. |
| `output_dir: str \| None` | `None` | Куда писать результат. `None` — рядом с каждым исходным файлом. |
| `font_name: str` | `"Times New Roman"` | Шрифт основного текста. |
| `font_size: float \| None` | `None` | Размер основного текста в пунктах; заголовки масштабируются от него. `None` — 12, а с `preset="gost"` — 14. |
| `preset` | `"default"` | `"gost"` — оформление по ГОСТ 7.32 (см. [оформление](#-как-оформляется-документ-word)). |
| `page_size` | `None` | `"A4"` или `"Letter"`; `None` — A4 (или размер шаблона). |
| `language: str` | `"auto"` | Язык документа для орфографии и переносов, например `"ru-RU"`; `auto` определяет его по тексту или front matter. |
| `line_breaks` | `"soft"` | `"soft"`: одиночный перенос строки внутри абзаца — пробел (CommonMark). `"preserve"`: разрыв строки. |
| `template: str \| None` | `None` | Образцовый `.docx`, из которого берутся стили, поля и колонтитулы. Должен существовать. |
| `toc: bool` | `False` | Вставить оглавление в начало (маркер `[TOC]` в тексте работает и без этого). |
| `footnotes` | `"native"` | `"native"`: сноски Word внизу страницы. `"section"`: нумерованный раздел в конце с заголовком `footnotes_heading`. |
| `footnotes_heading: str` | `"Footnotes"` | Заголовок этого раздела (только режим `"section"`; в русском документе автоматически «Сноски»). |
| `fetch_remote_images: bool` | `False` | Разрешить загрузку изображений по `http(s)`-ссылке. |
| `image_root: str \| None` | `None` | Расширить директорию, из которой разрешено читать локальные изображения. По умолчанию выводится из `inputs` (см. выше). |

`preview_markdown` принимает те же параметры, **кроме `output_dir`** — он ничего не пишет на диск.

`word_to_markdown` принимает `inputs` (обязателен), `output_dir` и `extract_media` (по умолчанию `true`: картинки сохраняются рядом с результатом в папку `<имя>_media/` и подключаются из Markdown).

`check_latex` принимает `formulas: list[str]` — LaTeX как есть, с обрамлением `$…$`, `$$…$$`, `\(…\)`, `\[…\]` или без него; окружения amsmath тоже принимаются.

`inspect_docx` принимает `path: str`.

### Что возвращают инструменты

`markdown_to_word` и `word_to_markdown` возвращают:

```
sources_found: int
converted: [{ source, output, warnings: [{ message, code, line }] }]
failed:    [{ source, error }]
```

`preview_markdown` возвращает ту же форму, но вместо `converted` — `previews: [{ source, warnings: [...] }]`, без поля `output`. `check_latex` возвращает `{ results: [{ formula, ok, error }], all_ok }`.

Что важно понимать при чтении этих полей:

- **`sources_found` стоит проверять в первую очередь.** Это количество подходящих файлов, в которые развернулись `inputs`. `0` значит, что пути не совпали ни с чем — это повод перепроверить пути, а не считать работу выполненной.
- Отказ одного файла не останавливает пакет: он попадает в `failed`, а остальные конвертируются как обычно. Читайте оба списка.
- `warnings` не фатальны: файл всё равно записан. У каждого предупреждения есть стабильный `code` (`formula_unsupported`, `image_not_found`, `image_outside_root`, `image_remote_disabled`, `image_fetch_failed`, `image_too_large`, `html_dropped`, `link_anchor_missing`, `footnote_unreferenced`, `math_prose`, …) и, если известна, строка `line` исходного Markdown (с единицы) — исправьте исходник в этой строке и запустите конвертацию заново.

> **Изменено в 1.2:** раньше `warnings` были простыми строками. Клиентам, читавшим их как строки, теперь нужно поле `message`.

### Примеры

Пути должны быть абсолютными. Относительный путь разрешается относительно рабочей директории **сервера**, а не проекта агента.

1. Конвертировать всю папку: `markdown_to_word(inputs=["/abs/path/docs"])`.
2. Проверить документ перед конвертацией: `preview_markdown(inputs=["/abs/path/README.md"])` — ничего не записывается, вы получаете предупреждения с номерами строк.
3. Отчёт по ГОСТ с оглавлением в папку сборки: `markdown_to_word(inputs=["/abs/path/report.md"], output_dir="/abs/path/build", preset="gost", toc=true)`.
4. Оформить по шаблону организации: `markdown_to_word(inputs=["/abs/path/memo.md"], template="/abs/path/company.docx")`.
5. Убедиться, что формулы станут уравнениями: `check_latex(formulas=["\\frac{a}{b}", "\\mathbb{R}^n"])`.
6. Проверить результат: `inspect_docx(path="/abs/path/build/report.docx")` — структура, формат страницы, сколько таблиц, изображений и формул попало в документ.
7. Вернуть Word-файл в Markdown вместе с картинками: `word_to_markdown(inputs=["/abs/path/report.docx"])`.

### Решение проблем

| Симптом | Причина и решение |
| --- | --- |
| Сервер не запускается: `ModuleNotFoundError: No module named 'mdtoword'` | Запуск из клона без `cwd`/`PYTHONPATH`. Установите пакет (`pip install ".[mcp]"`) и используйте `mdtoword-mcp` либо укажите `cwd` на репозиторий. |
| `sources_found: 0` и пустые списки | Пути не совпали ни с одним подходящим файлом: опечатка в пути, или папка без файлов `.md`/`.markdown` (либо `.docx` для обратного направления). |
| Предупреждение о том, что изображение не загружено | Загрузка удалённых изображений по умолчанию выключена. Передайте `fetch_remote_images=true`, и только для Markdown, которому доверяете. |
| Предупреждение «could not be fetched … not publicly routable» | Ссылка (или перенаправление) ведёт на локальный или частный адрес — такие загрузки запрещены всегда. |
| Предупреждение о том, что изображение вне разрешённого корня | Изображение лежит вне путей, переданных в `inputs`. Передайте `image_root`, чтобы расширить корень. |
| `template must be an existing .docx file` | Путь к шаблону неверен или относителен к директории сервера. Передайте абсолютный путь. |
| Результат записан не туда, куда ожидали | Относительный `output_dir` разрешился относительно рабочей директории сервера. Передайте абсолютный путь. |
| `tests/test_mcp_server.py` целиком пропускается | В интерпретаторе, которым запускаются тесты, не установлен SDK `mcp`. Установите `requirements-mcp.txt`. |

---

## 🛠️ Разработка

Тесты запускаются из корня проекта:

```bash
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests -p "test_*.py"
```

Если прогон завершается segfault'ом или трейсом падения вместо количества
пройденных тестов — дело в сборке Qt у интерпретатора: перезапустите через
собственный virtualenv проекта, а не через системный или anaconda Python.

Переменная `QT_QPA_PLATFORM=offscreen` нужна, чтобы тесты интерфейса работали без экрана. Тесты рендерят документы и проверяют их на уровне XML; вёрстку удобно смотреть через `soffice --headless --convert-to pdf` — так видно нумерацию, сноски и формулы, как их рисует текстовый процессор.

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
Посмотрите предупреждение в итоговом диалоге: там названа конкретная конструкция и строка, например `отчёт.md:12: Formula kept as text: … (Unsupported LaTeX command: \qedsymbol)`. Сама формула при этом сохранена в документе буквально — перепишите её через поддерживаемую конструкцию из [таблицы выше](#что-поддерживается) и запустите конвертацию заново.

**Доллар в тексте превратился во что-то странное**
Пишите буквальный знак как `\$`. Если конвертер увидел `$…$` вокруг обычного текста, он оставит его как есть и предупредит об этом — но лучше экранировать сразу.

**Строки, которые я переносил в исходнике, склеились**
Так работает CommonMark: одиночный перенос строки внутри абзаца — это пробел. Для разрыва строки закончите её двумя пробелами или `\`, либо используйте `line_breaks="preserve"` через MCP-сервер.

**Изображение не попало в документ**
Локальные пути считаются относительно `.md`-файла (кириллица и пробелы в именах файлов допустимы), ссылки `http(s)` скачиваются с таймаутом 10 секунд и только с публичных адресов. Если файл не найден, формат незнаком или сеть недоступна, на его месте окажется `[alt-текст]`, а в диалоге появится предупреждение с адресом.

**Большая пачка файлов «подвешивает» окно**
Конвертация идёт в потоке интерфейса, поэтому на длинных очередях окно откликается вяло. Прогресс при этом обновляется — дождитесь конца.

**Word при открытии предлагает обновить поля**
В документе есть оглавление; если разрешить обновление, Word допишет в него номера страниц.

**Word → Markdown что-то не сохранил**
У надписей (текстовых полей), примечаний и объединённых ячеек таблиц нет аналога в Markdown; итоговый диалог перечисляет, что было упрощено.
