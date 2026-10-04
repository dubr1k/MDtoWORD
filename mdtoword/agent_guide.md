# Markdown that converts cleanly to Word

A guide for agents using the mdtoword MCP server. Input is CommonMark +
GitHub Flavored Markdown (GFM) with the extensions below; everything listed
here becomes native Word structure (real headings, lists, tables, footnotes,
equations), not imitation formatting.

## Workflow

1. Write the Markdown using only the constructs below.
2. `check_latex` on every formula — fix any with `ok: false`.
3. `preview_markdown` — fix each warning at its `line`; repeat until clean.
4. `markdown_to_word` — pick options (`preset`, `toc`, `language`, …).
5. `inspect_docx` — confirm outline, counts and page setup.

## Document metadata (YAML front matter)

```yaml
---
title: Annual report
subtitle: Draft for review
author: Jane Doe
date: 2026-10-04
subject: Quarterly results     # `description` is accepted as well
keywords: [finance, report]    # or `tags`; a list or a comma-separated string
lang: en-US
---
```

`title`, `subtitle`, `author` and `date` form the title block; title,
author, subject, keywords and language also go into the document
properties. The front matter must be the very first thing in the file.

## Text

- Paragraphs are separated by a blank line. A single newline inside a
  paragraph is a space (`line_breaks="soft"`, CommonMark). For a hard line
  break end the line with `\` or two spaces, or call the tool with
  `line_breaks="preserve"`.
- `*italic*`, `**bold**`, `***both***`, `~~strikethrough~~`, `` `code` ``.
- `H~2~O` subscript, `x^2^` superscript, `==highlight==`.
- A literal dollar sign is `\$`; otherwise `$…$` starts a formula.
- Inline HTML is limited to `<br>`, `<sub>`, `<sup>`, `<kbd>`, `<mark>`,
  `<u>`, `<ins>`, `<del>`, `<s>`, `<b>`, `<strong>`, `<i>`, `<em>`,
  `<code>`, `<a href>`, `<img>` and `<details>`/`<summary>`. HTML comments
  are dropped silently. Any other tag is stripped (its text is kept) with a
  warning — use Markdown instead.

## Headings and links

- `#` … `######` become Word Heading 1–6 and appear in the table of
  contents. Do not skip levels.
- Every heading gets a bookmark with a GitHub-style slug: lowercase, spaces
  to `-`, punctuation dropped (`## Results & Discussion` → `#results--discussion`).
  Link to it with `[see results](#results--discussion)`.
- Links: `[text](https://…)`, autolinks `<https://…>`, and bare URLs
  (`https://example.com`) all become clickable hyperlinks.
- `[TOC]` alone on its own line inserts a table of contents at that spot
  (or pass `toc=true` to put one at the start). Word fills it in when the
  document's fields are updated.

## Lists

- `-`/`*`/`+` bullets and `1.` numbers; an ordered list keeps its start
  number (`3.` starts at 3) and each separate list restarts numbering.
- Nest by indenting under the parent item's text (up to 9 levels).
- An item may hold several paragraphs, code or a table: indent them to the
  item's text column and separate with blank lines.
- Task lists: `- [ ] todo`, `- [x] done`.
- Definition lists: a `Term` line followed by `: Definition` on the next line.

## Blocks

- `>` quotes, nested with `> >`. GitHub alerts become shaded callouts with a
  bold title (Note/Примечание, …):
  `> [!NOTE]` (also `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]`, `[!CAUTION]`) as
  the first line of the quote, the text on the following `>` lines.
- Fenced code blocks (three backticks or tildes, optional language) keep
  their whitespace and use a monospace font.

## Tables

GFM pipe tables (see the example below). Column alignment (`:--`, `--:`, `:-:`) is kept; cells may contain inline
formatting, links and inline math. A paragraph `Table: caption` (or
`Таблица: caption`, or pandoc's `: caption`) directly before or after a
table becomes its numbered caption above the table — "Таблица N — …" in a
Russian document, "Table N: …" otherwise. Keep one row per line; no merged
cells.

## Images

`![Alt text](images/figure.png "Optional title")`

- Use local paths relative to the .md file. Images must sit inside the
  folder passed to the tool (or pass `image_root`). Remote `http(s)` images
  are skipped unless `fetch_remote_images=true`.
- Images are scaled down to the text width.
- An image alone in its paragraph with alt text or a title becomes a centred
  figure with a numbered caption below it (the title wins over the alt
  text) — "Рисунок N — …" in a Russian document, "Figure N: …" otherwise.
  An image inside a sentence stays inline and gets no caption.
- PNG, JPEG, GIF, BMP, TIFF, WebP and SVG are accepted.

## Footnotes

`Text with a note.[^1]` and, anywhere in the file, `[^1]: The note.` They
become native Word footnotes (`footnotes="section"` collects them at the end
instead). Labels may be words: `[^source]`.

## Math

- Inline `$E = mc^2$`, display `$$ … $$` on its own lines.
- amsmath environments as separate blocks: `equation`, `align`, `gather`,
  `multline`, `alignat`, `flalign` (starred forms too); `aligned`, `split`,
  `gathered` inside `$$ … $$`. Use `\\` for new lines and `&` to align.
- `\tag{n}` puts the number `(n)` at the right of the equation.
- Everything becomes an editable Word equation. Run `check_latex` on
  unusual commands; unsupported formulas are kept as text with a warning.
- No blank lines inside a formula.

## Example

```markdown
---
title: Heat transfer note
author: Lab 3
lang: en-US
---

[TOC]

# Model

The flux is given by Fourier's law[^f]:

$$
q = -k \nabla T \tag{1}
$$

> [!TIP]
> For a slab, $q = k \frac{\Delta T}{L}$.

![Temperature profile](img/profile.png)

Table: Materials

| Material | $k$, W/(m·K) |
|----------|-------------:|
| Copper   |          401 |

[^f]: J. Fourier, 1822.
```

## Options worth knowing

- `preset="gost"`: GOST 7.32-2017 — A4, margins 30/15/20/20 mm
  (left/right/top/bottom), 1.5 spacing, 1.25 cm indent, 14 pt, centred page
  numbers, "Рисунок N — …" / "Таблица N — …" captions.
- `template="/abs/path/reference.docx"`: reuse its styles, margins, headers.
- `language="ru-RU"` (default `auto`), `page_size="Letter"`.
