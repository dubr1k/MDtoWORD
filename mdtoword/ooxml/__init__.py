"""Low-level WordprocessingML (OOXML) primitives for the Markdown renderer.

python-docx 1.1.2 covers paragraphs, runs, tables, sections and raster
pictures, but has no API for real Word numbering, footnotes, bookmarks, fields
or SVG pictures. This package fills those gaps with small, focused helpers that
write schema-valid markup: every child element is inserted at the position the
ECMA-376 schema requires (``w:pPr``, ``w:rPr``, ``w:settings`` and friends are
ordered sequences, and Word reports a file as corrupt when they are not).

Modules, one concern each:

- ``_xml`` -- schema element-order sequences and ordered insertion helpers;
- ``_styles`` -- style lookup by id or name;
- ``numbering`` -- real Word numbering for lists;
- ``footnotes`` -- the footnotes story part and :class:`FootnoteManager`;
- ``links`` -- heading slugs, bookmarks and hyperlinks;
- ``fields`` -- complex fields, the table of contents, update-on-open;
- ``page`` -- page setup, text area, proofing language, page numbers;
- ``images`` -- raster picture sizing, alt text, format normalisation;
- ``svg`` -- SVG parsing, sanitising and embedding;
- ``formatting`` -- shading, borders, keep options, table-row flags.

Importing this package registers :class:`FootnotesPart` with python-docx's
``PartFactory`` (via ``footnotes``) so documents that already carry a
footnotes part (a saved result, or a user template) load it as a story part.
"""

from __future__ import annotations

# Not in __all__, but kept importable from the package (``X as X`` marks an
# explicit re-export): the tests assert element order against the schema
# sequences and probe the hardened SVG parser.
from ._xml import _PBDR_SEQUENCE as _PBDR_SEQUENCE
from ._xml import _PPR_SEQUENCE as _PPR_SEQUENCE
from ._xml import _RPR_SEQUENCE as _RPR_SEQUENCE
from ._xml import _SETTINGS_SEQUENCE as _SETTINGS_SEQUENCE
from .fields import (
    add_field,
    add_toc,
    append_toc,
    insert_toc_before,
    request_field_update_on_open,
)
from .footnotes import (
    FOOTNOTE_REFERENCE_STYLE_NAME,
    FOOTNOTE_TEXT_STYLE_NAME,
    FootnoteManager,
    FootnotesPart,
)
from .formatting import BorderSpec as BorderSpec
from .formatting import (
    set_cant_split,
    set_keep_lines,
    set_paragraph_borders,
    set_paragraph_shading,
    set_run_shading,
    set_table_header_repeat,
)
from .images import normalize_raster, picture_size_to_fit, set_picture_description
from .links import (
    add_bookmark,
    add_external_hyperlink,
    add_hyperlink_run,
    add_internal_hyperlink,
    bookmark_name,
    github_slug,
)
from .numbering import ListNumbering
from .page import (
    PAGE_SIZES_MM,
    add_page_number_footer,
    apply_page_setup,
    detect_language,
    set_document_language,
    text_height,
    text_width,
)
from .svg import _parse_svg_root as _parse_svg_root
from .svg import add_svg_picture, sanitize_svg, svg_intrinsic_size

__all__ = [
    "FOOTNOTE_REFERENCE_STYLE_NAME",
    "FOOTNOTE_TEXT_STYLE_NAME",
    "FootnoteManager",
    "FootnotesPart",
    "ListNumbering",
    "PAGE_SIZES_MM",
    "add_bookmark",
    "add_external_hyperlink",
    "add_field",
    "add_hyperlink_run",
    "add_internal_hyperlink",
    "add_page_number_footer",
    "add_svg_picture",
    "add_toc",
    "append_toc",
    "apply_page_setup",
    "bookmark_name",
    "detect_language",
    "github_slug",
    "insert_toc_before",
    "normalize_raster",
    "picture_size_to_fit",
    "request_field_update_on_open",
    "set_cant_split",
    "set_document_language",
    "set_keep_lines",
    "set_paragraph_borders",
    "set_paragraph_shading",
    "set_picture_description",
    "set_run_shading",
    "set_table_header_repeat",
    "sanitize_svg",
    "svg_intrinsic_size",
    "text_height",
    "text_width",
]
