"""Raster pictures: fit-to-page sizing, alt text and format normalisation."""

from __future__ import annotations

import math
from io import BytesIO
from typing import Any

from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from docx.shape import InlineShape
from docx.shared import Emu, Length
from docx.text.run import Run
from PIL import Image as PILImage

_EMU_PER_INCH = 914400
_DEFAULT_DPI = 96.0
_DOCX_NATIVE_FORMATS = {
    "PNG": "png", "JPEG": "jpeg", "MPO": "jpeg", "GIF": "gif", "BMP": "bmp", "TIFF": "tiff",
}


def _usable_dpi(value: Any) -> float | None:
    try:
        dpi = float(value)
    except (TypeError, ValueError):
        return None
    return dpi if math.isfinite(dpi) and dpi > 0 else None


def picture_size_to_fit(
    image_px: tuple[int, int],
    dpi: tuple[float, float] | None,
    max_width: Length,
    max_height: Length,
) -> tuple[Length, Length]:
    """Natural picture size (pixels at ``dpi``, 96 when unknown), scaled DOWN only
    to fit ``max_width`` x ``max_height`` with the aspect ratio preserved."""
    px_w, px_h = image_px
    if px_w <= 0 or px_h <= 0:
        raise ValueError(f"image has no area: {image_px!r}")
    if max_width <= 0 or max_height <= 0:
        raise ValueError("max_width and max_height must be positive")
    dpi_x, dpi_y = (dpi if dpi is not None else (None, None))
    dpi_x, dpi_y = _usable_dpi(dpi_x), _usable_dpi(dpi_y)
    dpi_x = dpi_x or dpi_y or _DEFAULT_DPI
    dpi_y = dpi_y or dpi_x
    width = px_w / dpi_x * _EMU_PER_INCH
    height = px_h / dpi_y * _EMU_PER_INCH
    scale = min(1.0, int(max_width) / width, int(max_height) / height)
    return Emu(max(1, round(width * scale))), Emu(max(1, round(height * scale)))


def set_picture_description(
    inline_shape_or_run: Any, description: str, title: str | None = None
) -> None:
    """Set the alt text (``wp:docPr/@descr``, and ``@title``) of a picture.

    Accepts the InlineShape returned by ``run.add_picture()``, the Run that
    holds the picture (its last drawing is used) or a raw drawing element.
    """
    target = inline_shape_or_run
    if isinstance(target, InlineShape):
        element = target._inline
    elif isinstance(target, Run):
        element = target._r
    else:
        element = target
    doc_prs = element.xpath(".//wp:docPr") if element.tag != qn("wp:docPr") else [element]
    if not doc_prs:
        raise ValueError("no picture (wp:docPr) found")
    doc_pr = doc_prs[-1]
    doc_pr.set("descr", description)
    if title is not None:
        doc_pr.set("title", title)


def _python_docx_accepts(image_bytes: bytes) -> bool:
    try:
        DocxImage.from_blob(image_bytes)
    except Exception:  # python-docx header parsers raise assorted types on bad input
        return False
    return True


def _to_png(image: Any) -> bytes:
    image.load()
    if image.mode == "P":
        image = image.convert("RGBA" if "transparency" in image.info else "RGB")
    elif image.mode not in ("1", "L", "LA", "RGB", "RGBA"):
        image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
    out = BytesIO()
    dpi = image.info.get("dpi")
    if dpi:
        image.save(out, "PNG", dpi=dpi)
    else:
        image.save(out, "PNG")
    return out.getvalue()


def normalize_raster(image_bytes: bytes) -> tuple[bytes, str]:
    """Return ``(bytes, format)`` python-docx can embed.

    PNG/JPEG/GIF/BMP/TIFF pass through unchanged; anything else Pillow can
    read (WEBP, ICO, AVIF…) is converted to PNG. Raises
    ``ValueError("unsupported image format")`` when Pillow cannot read it.
    """
    try:
        with PILImage.open(BytesIO(image_bytes)) as image:
            native = _DOCX_NATIVE_FORMATS.get((image.format or "").upper())
            if native is not None and _python_docx_accepts(image_bytes):
                return image_bytes, native
            return _to_png(image), "png"
    except PILImage.DecompressionBombError as exc:
        raise ValueError("image is too large to decode") from exc
    except (OSError, ValueError, SyntaxError, EOFError) as exc:
        raise ValueError("unsupported image format") from exc
