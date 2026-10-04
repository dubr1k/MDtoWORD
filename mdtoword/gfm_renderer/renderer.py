"""``GfmDocxRenderer``: the public class, assembled from the mixins.

The class itself only holds the constructor, the markdown-it parser
configuration and ``render`` -- the single pass over the token stream.
Everything ``render`` calls lives in a mixin, one concern per module; they
all derive from :class:`~.state.RendererState` and share one instance, so a
mixin calls another's methods simply through ``self``.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

from docx.document import Document as DocumentType
from docx.shared import Pt
from markdown_it import MarkdownIt
from mdit_py_plugins.amsmath import amsmath_plugin
from mdit_py_plugins.deflist import deflist_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.subscript import sub_plugin

from ..md_extensions import mark_plugin, parse_front_matter, sup_plugin
from ..ooxml import FootnoteManager, ListNumbering
from ..options import DocumentOptions
from .blocks import BlockMixin
from .containers import ContainerMixin
from .document import DocumentSetupMixin
from .equations import MathMixin
from .footnotes import FootnoteMixin
from .images import ImageMixin
from .inline import InlineMixin
from .raw_html import RawHtmlMixin
from .source import SourceMixin
from .special import SpecialParagraphMixin
from .state import RendererState
from .tables import TableMixin
from .toc import TitleTocMixin


class GfmDocxRenderer(
    SourceMixin,
    DocumentSetupMixin,
    TitleTocMixin,
    BlockMixin,
    SpecialParagraphMixin,
    ContainerMixin,
    InlineMixin,
    RawHtmlMixin,
    ImageMixin,
    MathMixin,
    TableMixin,
    FootnoteMixin,
    RendererState,
):
    """Render a GFM token stream into a Word document."""

    def __init__(
        self,
        font_name: str,
        font_size: Pt,
        footnotes_heading: str = "Footnotes",
        allow_remote_images: bool = True,
        image_roots: Sequence[Path] | None = None,
        *,
        document_options: DocumentOptions | None = None,
    ):
        self.options = document_options or DocumentOptions()
        self.font_name = font_name
        self.font_size = font_size
        self.footnotes_heading = footnotes_heading
        self.allow_remote_images = allow_remote_images
        self.image_roots = image_roots
        # Resolved once here rather than per image: __init__ runs once per
        # render, while _append_image runs once per image in the document.
        # None means unrestricted (the GUI's default -- see app.py, which
        # never passes image_roots).
        self._resolved_image_roots: list[Path] | None = (
            None if image_roots is None else [Path(root).resolve() for root in image_roots]
        )
        self.document: DocumentType
        self.warnings: list[str]
        self._paragraph: Any = None

    @staticmethod
    def _parser() -> MarkdownIt:
        return (
            MarkdownIt("js-default", {"breaks": False, "html": True, "linkify": True})
            .enable("linkify")
            .use(front_matter_plugin)
            .use(footnote_plugin)
            .use(deflist_plugin)
            .use(sub_plugin)
            .use(sup_plugin)
            .use(mark_plugin)
            .use(dollarmath_plugin, allow_digits=False, allow_blank_lines=False, double_inline=True)
            .use(amsmath_plugin)
        )

    def render(
        self, markdown: str, source_path: Path | None = None
    ) -> tuple[DocumentType, list[str]]:
        self._reset()
        markdown = self._prepare_source(markdown)
        environment: dict[str, Any] = {}
        tokens = self._parser().parse(markdown, environment)
        self._warn_unreferenced_footnotes(markdown, environment)
        if tokens and tokens[0].type == "front_matter":
            self._front_matter, skipped = parse_front_matter(tokens[0].content)
            if skipped:
                self._warn(
                    "Front matter entries not understood and ignored: "
                    + ", ".join(skipped[:5]) + ("…" if len(skipped) > 5 else ""),
                    "front_matter_ignored",
                    line=1,
                )
        self._language = self._resolve_language(markdown)
        self.document = self._new_document()
        self._configure_document()
        self._numbering = ListNumbering(self.document)
        self._footnotes = (
            FootnoteManager(self.document) if self.options.footnotes == "native" else None
        )
        self._prepare_heading_anchors(tokens)

        if self.options.toc and (not tokens or tokens[0].type != "front_matter"):
            self._insert_toc()

        index = 0
        while index < len(tokens):
            token = tokens[index]
            if token.map:
                self._line = token.map[0] + 1
            index = self._render_block(tokens, index, source_path)

        return self.document, self.warnings
