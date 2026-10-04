"""The per-render state every mixin shares, and the primitives on it.

``GfmDocxRenderer`` is one object assembled from concern-specific mixins
(see the package docstring). They share state by living on the same
instance; this base class, which every mixin derives from, is where that
state is defined. ``_reset`` gives every attribute its fresh value at the
start of each ``render``; ``_warn`` and the ``_gost``/``_russian``/
``_stamp_runs`` switches, which every concern needs, live next to it.

Which module writes what (any module may read anything):

- ``renderer.py`` (``render``): ``document``, ``_front_matter``,
  ``_language``, ``_numbering``, ``_footnotes``, ``_line`` per block.
- ``source.py`` (pre-pass): ``_heading_bookmark_at``, ``_slug_to_bookmark``,
  ``_headings``.
- ``document.py``: ``_text_width``, ``_text_height``.
- ``blocks.py`` / ``special.py``: the open containers ``_lists``,
  ``_quotes``, ``_definition_depth``, ``_in_term``, the current
  ``_heading_level``/``_current_heading_bookmark``, ``_figure_count``,
  ``_table_count``, ``_caption_context``; ``_last_table`` is consumed by a
  caption written after its table.
- ``inline.py`` / ``raw_html.py``: ``_link``, ``_warned_html``; the source
  ``_line`` advances on soft and hard breaks; inside a table row the inline
  children are only stored in ``_table_cell_children``.
- ``tables.py``: ``_table_rows``, ``_table_row``, ``_table_header``,
  ``_table_line``, ``_table_cell_children``, ``_table_columns``,
  ``_table_row_alignment``, ``_cell_context``, ``_last_table``, and
  ``_line`` per cell while the cells render.
- ``footnotes.py``: ``_footnote_target``, ``_footnote_section_open``,
  ``_pending_footnote_label`` (cleared by ``_add_paragraph`` once written),
  ``_nested_footnote_parent``, ``_footnote_word_ids``, ``_skip_footnote``,
  ``_saved_containers``.
- ``toc.py``: ``_toc_inserted``.
- everyone: ``_paragraph`` -- the paragraph inline content currently goes
  into, ``None`` between blocks.

Tables and footnotes swap the container state out (``_lists``, ``_quotes``,
``_definition_depth``) while they render and restore it afterwards.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.shared import Cm, Length

from ..errors import ConversionWarning
from .helpers import _Cell, _ListState

if TYPE_CHECKING:
    from pathlib import Path

    from docx.document import Document as DocumentType
    from docx.shared import Pt

    from ..ooxml import FootnoteManager, ListNumbering
    from ..options import DocumentOptions


class RendererState:
    """State shared by all mixins of ``GfmDocxRenderer``; see the module docstring."""

    # Declarations only (no class attributes are created): the attributes
    # set outside _reset, so every mixin module can see where they come from.
    # Set once by GfmDocxRenderer.__init__.
    options: DocumentOptions
    font_name: str
    font_size: Pt
    footnotes_heading: str
    allow_remote_images: bool
    _resolved_image_roots: list[Path] | None
    # Set by render() before the first block is rendered.
    document: DocumentType
    _language: str
    _numbering: ListNumbering
    _footnotes: FootnoteManager | None
    # Set on demand: the containers a footnote put aside, and the alignment
    # of the table cell being collected.
    _saved_containers: tuple[list[_ListState], list[str | None], int]
    _table_row_alignment: str | None

    def _reset(self) -> None:
        self.warnings = []
        self._paragraph = None
        self._lists: list[_ListState] = []
        self._quotes: list[str | None] = []
        self._definition_depth = 0
        self._in_term = False
        self._heading_level: int | None = None
        self._table_rows: list[list[_Cell]] | None = None
        self._table_row: list[_Cell] | None = None
        self._table_header = False
        self._table_line: int | None = None
        self._cell_context = False
        self._last_table: Any = None
        self._footnote_target: int | None = None
        self._footnote_section_open = False
        self._pending_footnote_label: str | None = None
        self._nested_footnote_parent: dict[int, int] = {}
        self._link: Any = None
        self._heading_bookmark_at: dict[int, str] = {}
        self._current_heading_bookmark: str | None = None
        self._caption_context = False
        self._footnote_word_ids: dict[int, list[int]] = {}
        self._skip_footnote = False
        self._line: int | None = None
        self._front_matter: dict[str, str | list[str]] = {}
        self._headings: list[tuple[int, str, str]] = []
        self._slug_to_bookmark: dict[str, str] = {}
        self._figure_count = 0
        self._table_count = 0
        self._toc_inserted = False
        self._warned_html: set[str] = set()
        self._text_width: Length = Cm(16)
        self._text_height: Length = Cm(24)

    @property
    def _gost(self) -> bool:
        return self.options.preset == "gost"

    @property
    def _russian(self) -> bool:
        return self._language.lower().startswith("ru")

    @property
    def _stamp_runs(self) -> bool:
        """Should runs carry explicit font/size? Not under a template's styles."""
        return self.options.template is None

    def _warn(self, message: str, code: str = "general", line: int | None = None) -> None:
        self.warnings.append(
            ConversionWarning(message, code=code, line=line if line is not None else self._line)
        )
