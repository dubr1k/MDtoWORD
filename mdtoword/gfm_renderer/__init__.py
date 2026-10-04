"""Render a GitHub Flavored Markdown token stream into a Word document.

The renderer walks markdown-it's flat token stream once, keeping a small
amount of container state -- open lists, blockquotes, definition
descriptions, the footnote being filled -- and decides for every paragraph
where it goes (body or a native footnote) and how it is indented. Word has
no notion of nesting, so "inside a list inside a quote" is expressed purely
as numbering levels, left indents and borders on flat paragraphs.

Everything the renderer cannot represent faithfully ends up as a
:class:`~mdtoword.errors.ConversionWarning` carrying a stable ``code`` and
the Markdown source line, never as a silent loss.

Layout
------
``GfmDocxRenderer`` (``renderer.py``) is a thin class -- constructor, parser
setup and ``render`` -- assembled from one mixin per concern:

- ``state.py``       ``RendererState``: the shared per-render state
  (``_reset``), ``_warn`` and the GOST/Russian/template switches. Every
  mixin derives from it.
- ``source.py``      input normalisation and the pre-passes (language,
  unreferenced footnotes, heading bookmarks).
- ``document.py``    document creation, template, page setup, styles,
  core properties.
- ``toc.py``         front-matter title block and table of contents.
- ``blocks.py``      the block-token dispatch table; code blocks, rules.
- ``special.py``     figures, captions, TOC markers, GitHub alerts.
- ``containers.py``  paragraph placement and list/quote/definition
  decoration.
- ``inline.py``      runs, formatting, links, breaks.
- ``raw_html.py``    inline HTML tags and HTML blocks.
- ``images.py``      image fetching/reading, sandboxing, embedding.
- ``equations.py``   inline/display math, amsmath, equation numbers.
- ``tables.py``      table buffering and building.
- ``footnotes.py``   native and section-mode footnotes.
- ``constants.py``, ``helpers.py``  stateless constants, functions and
  small records shared by the mixins.

How the mixins share state
--------------------------
The renderer is stateful and single-pass, so the concerns are not separate
objects passing data around: they are mixins of one class, and every
method still runs on the one renderer instance. A mixin reads and writes
the shared attributes (``self._paragraph``, ``self._lists``, …) and calls
methods of other mixins through ``self`` exactly as the single class did.
``state.py`` documents every attribute and which module writes it. The
mixin modules import only ``constants``, ``helpers`` and ``state`` from
this package -- never each other -- so there are no import cycles; only
``renderer.py`` imports the mixins.
"""

from .constants import _MAX_REMOTE_IMAGE_BYTES
from .helpers import _is_remote_target
from .renderer import GfmDocxRenderer

# The two private names are re-exported for the tests and for callers of the
# former single module. fetch_image and Path are deliberately *not*
# re-exported: patch them in mdtoword.gfm_renderer.images, where the image
# code looks them up -- a patch here would silently have no effect.
__all__ = ["GfmDocxRenderer", "_MAX_REMOTE_IMAGE_BYTES", "_is_remote_target"]
