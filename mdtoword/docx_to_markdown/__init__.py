"""High-fidelity conversion of Word documents (.docx) to Markdown.

The converter reads the WordprocessingML tree directly rather than through
python-docx's convenience proxies, because those hide most of what matters
here: ``Paragraph.runs`` skips runs inside hyperlinks, nothing exposes
equations, footnotes or list numbering, and ``iter_inner_content`` skips
content controls. The design, top down:

* **Blocks in order.** The body is walked child by child (content controls
  opened), so tables stay where they were. Each paragraph is classified --
  heading, list item, code, quote, display equation, thematic break -- from
  its style chain, numbering and content, then written by a small state
  machine that groups consecutive code lines, quote paragraphs and list
  items into one Markdown block each.
* **Inline content as a tree.** A paragraph's runs become a flat list of
  formatted text segments plus atomic items (links, images, math, footnote
  references). Adjacent segments with identical formatting are merged, the
  list is folded into a tree of emphasis spans (longest span outermost), and
  the tree is printed with whitespace kept *outside* the markers. A span
  whose ``*``/``~~``/``==`` delimiters would not parse back -- CommonMark's
  flanking rules reject e.g. ``x**(a)**`` -- falls back to its HTML tag.
* **Escaping.** Plain text is escaped so it re-parses to the same text:
  emphasis and link punctuation everywhere, ``$`` (the forward converter's
  math delimiter), and block syntax (``#``, ``>``, ``-``, ``1.``) wherever a
  line starts -- including after a hard line break.
* **Lists** follow Word's own numbering: ``w:numPr`` (direct or from the
  style), the abstract numbering's level formats and the running counters
  per list, with ``w:startOverride`` honoured. Nesting is decided by the
  item's text indent, which also works for style-based lists such as
  ``List Bullet 2`` that live on level 0 of a separate list.
* **Equations** go through :mod:`mdtoword.omml_latex`; images are extracted
  to ``<output stem>_media/``; footnotes and endnotes are read from their own
  parts (relationships resolved against *those* parts) and appended as
  ``[^N]:`` definitions.

What cannot be represented is reported, once per kind and with a count, as
:class:`~mdtoword.errors.ConversionWarning` objects with stable codes such as
``table_merged_cells`` or ``textbox_skipped``.

Package layout, from the entry point down:

* :mod:`.converter` -- :class:`WordToMarkdownConverter`, the public API;
* :mod:`.document` -- one conversion's shared state, the story walk and the
  final assembly;
* :mod:`.blocks` -- the block writer (paragraph dispatch, code grouping),
  with :mod:`.lists`, :mod:`.quotes`, :mod:`.tables`, :mod:`.captions`,
  :mod:`.equations` and :mod:`.toc` for the block kinds it delegates;
* :mod:`.classify` -- paragraph classification and run formatting, with
  :mod:`.code` for code recognition;
* :mod:`.collect` -- runs to inline items, with :mod:`.fields`,
  :mod:`.media` (pictures) and :mod:`.notes` (footnotes, endnotes);
* :mod:`.inline`, :mod:`.render`, :mod:`.escaping` -- the inline model and
  its Markdown printing;
* :mod:`.styles`, :mod:`.numbering` -- style chains and list numbering;
* :mod:`.anchors`, :mod:`.metadata`, :mod:`.diagnostics` -- heading slugs,
  front matter and warnings;
* :mod:`.wordml` -- XML names and element helpers shared by all of them.
"""

from __future__ import annotations

from .converter import WordToMarkdownConverter

__all__ = ["WordToMarkdownConverter"]
