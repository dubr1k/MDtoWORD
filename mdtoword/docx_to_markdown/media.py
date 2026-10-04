"""Pictures: image parts extracted to ``<output stem>_media/`` and linked as Markdown images.

DrawingML pictures (an SVG original preferred over its PNG fallback) and
VML image data both become ``![alt](path "title")``; a figure caption that
was paired with the picture becomes its title. Shapes, charts and text
boxes have no Markdown form and are reported instead.
"""

from __future__ import annotations

import posixpath
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .escaping import _escape_text, _link_destination
from .inline import _Raw
from .wordml import (
    A_BLIP,
    ASVG_SVGBLIP,
    O_HR,
    O_TITLE,
    R_EMBED,
    R_ID,
    R_LINK,
    V_IMAGEDATA,
    V_RECT,
    W_TXBXCONTENT,
    WP_DOCPR,
    _q,
)

if TYPE_CHECKING:
    from .diagnostics import _Diagnostics

_EXTENSIONS = {
    "image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif",
    "image/bmp": "bmp", "image/tiff": "tiff", "image/svg+xml": "svg",
    "image/webp": "webp", "image/x-emf": "emf", "image/x-wmf": "wmf",
}


@dataclass
class _Media:
    """Extracted images: one file per image part, named in document order."""

    directory: Path | None
    prefix: str
    names: dict[str, str] = field(default_factory=dict)
    count: int = 0

    def store(self, part: Any) -> str | None:
        key = str(part.partname)
        if key in self.names:
            return self.names[key]
        if self.directory is None:
            return None
        extension = posixpath.splitext(key)[1].lstrip(".").lower()
        if not extension:
            extension = _EXTENSIONS.get(getattr(part, "content_type", ""), "bin")
        self.count += 1
        name = f"image{self.count}.{extension}"
        self.directory.mkdir(parents=True, exist_ok=True)
        (self.directory / name).write_bytes(part.blob)
        link = f"{self.prefix}/{name}" if self.prefix else name
        self.names[key] = link
        return link


class _Pictures:
    """Markdown images for ``w:drawing`` and ``w:pict`` elements."""

    def __init__(self, media: _Media, diagnostics: _Diagnostics) -> None:
        self.media = media
        self.diagnostics = diagnostics
        # Caption text waiting for the next picture: a figure caption is
        # folded into the image's title instead of staying a paragraph.
        self.caption: str | None = None

    def drawing(self, element: Any, part: Any) -> list[_Raw | None]:
        """The pictures of a DrawingML drawing, in order."""
        blips = list(element.iter(A_BLIP))
        textboxes = next(element.iter(W_TXBXCONTENT), None) is not None
        if textboxes:
            self.diagnostics.warn("textbox_skipped")
        if not blips:
            if not textboxes:
                self.diagnostics.warn("image_unsupported")
            return []
        properties = next(element.iter(WP_DOCPR), None)
        alt = ""
        if properties is not None:
            alt = properties.get("descr") or properties.get("title") or ""
        images = []
        for blip in blips:
            embed = blip.get(R_EMBED)
            vector = next(blip.iter(ASVG_SVGBLIP), None)
            if vector is not None and vector.get(R_EMBED):
                embed = vector.get(R_EMBED)
            images.append(self._image(part, embed, blip.get(R_LINK), alt))
        return images

    def vml(self, element: Any, part: Any) -> list[_Raw | None]:
        """The picture of a VML ``w:pict``, if it is one."""
        data = next(element.iter(V_IMAGEDATA), None)
        if data is not None and (data.get(R_ID) or data.get(_q("r:href"))):
            shape = data.getparent()
            alt = (shape.get("alt") if shape is not None else None) or data.get(O_TITLE) or ""
            return [self._image(part, data.get(R_ID), data.get(_q("r:href")), alt)]
        if any(rect.get(O_HR) in ("t", "true") for rect in element.iter(V_RECT)):
            return []  # a horizontal rule; the paragraph classifier handles it
        if next(element.iter(W_TXBXCONTENT), None) is not None:
            self.diagnostics.warn("textbox_skipped")
        else:
            self.diagnostics.warn("image_unsupported")
        return []

    def _image(self, part: Any, embed: str | None, linked: str | None, alt: str) -> _Raw | None:
        alt = " ".join(alt.split())
        caption, self.caption = self.caption, None
        title = " ".join((caption or "").split())
        alt = alt or title
        if title == alt:
            # Markdown renderers caption a picture with its title, else its alt.
            title = ""
        alt_markdown = _escape_text(alt)
        title_markdown = ""
        if title:
            escaped = title.replace("\\", "\\\\").replace('"', '\\"')
            title_markdown = f' "{escaped}"'
        if embed:
            target = part.related_parts.get(embed)
            if target is None:
                self.diagnostics.warn("image_missing")
                return None
            path = self.media.store(target)
            if path is None:
                self.diagnostics.warn("image_not_extracted")
                text = " — ".join(part for part in (alt, title) if part)
                return _Raw(_escape_text(text), text) if text else None
            destination = _link_destination(path)
            return _Raw(f"![{alt_markdown}]({destination}{title_markdown})", alt)
        if linked:
            relationship = part.rels.get(linked)
            if relationship is not None:
                destination = _link_destination(relationship.target_ref)
                return _Raw(f"![{alt_markdown}]({destination}{title_markdown})", alt)
        self.diagnostics.warn("image_missing")
        return None
