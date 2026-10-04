"""SVG pictures: safe parsing, sanitising, intrinsic size and Word 2016+ embedding."""

from __future__ import annotations

import hashlib
import math
import re
from io import BytesIO
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.parts.image import ImagePart
from docx.shape import InlineShape
from docx.shared import Emu, Length
from docx.text.run import Run
from lxml import etree
from PIL import Image as PILImage

_EMU_PER_PX = 9525  # at 96 dpi
_SVG_CONTENT_TYPE = "image/svg+xml"
_SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
_ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
# CSS's default size for a replaced element without intrinsic dimensions.
_SVG_FALLBACK_PX = (300.0, 150.0)
_SVG_UNITS_PX = {
    "": 1.0, "px": 1.0, "pt": 96 / 72, "pc": 16.0, "mm": 96 / 25.4,
    "cm": 96 / 2.54, "in": 96.0, "q": 96 / 101.6, "em": 16.0, "ex": 8.0,
}
_SVG_LENGTH = re.compile(
    r"^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*([A-Za-z%]*)\s*$"
)


def _safe_xml_parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False, no_network=True, huge_tree=False, load_dtd=False,
        dtd_validation=False, remove_comments=True, remove_pis=True,
    )


def _parse_svg_root(svg_bytes: bytes) -> Any:
    try:
        root = etree.fromstring(svg_bytes, _safe_xml_parser())
    except (etree.XMLSyntaxError, ValueError) as exc:
        raise ValueError(f"invalid SVG: {exc}") from exc
    if root is None or etree.QName(root).localname != "svg":
        raise ValueError("invalid SVG: root element is not <svg>")
    return root


def _svg_length_px(value: str | None) -> float | None:
    if value is None:
        return None
    match = _SVG_LENGTH.match(value)
    if match is None:
        return None
    factor = _SVG_UNITS_PX.get(match.group(2).lower())
    if factor is None:  # percentages and unknown units are not intrinsic sizes
        return None
    px = float(match.group(1)) * factor
    return px if math.isfinite(px) and px > 0 else None


def _svg_viewbox(value: str | None) -> tuple[float, float] | None:
    if not value:
        return None
    parts = re.split(r"[\s,]+", value.strip())
    if len(parts) != 4:
        return None
    try:
        width, height = float(parts[2]), float(parts[3])
    except ValueError:
        return None
    if not (math.isfinite(width) and math.isfinite(height) and width > 0 and height > 0):
        return None
    return width, height


def _internal_entity_parser() -> etree.XMLParser:
    """Expands *internal* entities only (libxml2 caps the amplification, so
    "billion laughs" input fails to parse); external entities and the DTD
    are never loaded."""
    return etree.XMLParser(
        resolve_entities="internal", no_network=True, huge_tree=False, load_dtd=False,
        dtd_validation=False, remove_comments=True, remove_pis=True,
    )


def sanitize_svg(svg_bytes: bytes) -> bytes | None:
    """Re-serialize an SVG for embedding: UTF-8, no DOCTYPE/internal subset,
    no comments, processing instructions or entity references.

    Word and the Open XML validator reject a DTD inside ``word/media/*.svg``,
    yet standard Illustrator exports carry one and use internal entities for
    the SVG namespace and styles -- often inside attribute values, where a
    reference cannot survive without its DTD. Internal entities are therefore
    expanded (amplification-capped); external ones are never fetched, and an
    SVG that needs them yields ``None`` (callers embed the PNG only).
    Raises ValueError for input that is not an SVG document at all.
    """
    _parse_svg_root(svg_bytes)
    try:
        root = etree.fromstring(svg_bytes, _internal_entity_parser())
        cleaned = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
        # Without the DTD any leftover reference makes this parse fail.
        reparsed = etree.fromstring(cleaned, _safe_xml_parser())
    except (etree.XMLSyntaxError, ValueError, TypeError):
        return None
    if (
        etree.QName(reparsed).localname != "svg"
        or reparsed.getroottree().docinfo.doctype
        or next(reparsed.iter(etree.Entity), None) is not None
    ):
        return None
    return cleaned


def svg_intrinsic_size(svg_bytes: bytes) -> tuple[float, float] | None:
    """SVG size in CSS px (96 dpi) from width/height (px, pt, pc, mm, cm, in…) or
    the viewBox. ``None`` when it cannot be determined or the SVG is invalid."""
    try:
        root = _parse_svg_root(svg_bytes)
    except ValueError:
        return None
    width = _svg_length_px(root.get("width"))
    height = _svg_length_px(root.get("height"))
    if width and height:
        return width, height
    viewbox = _svg_viewbox(root.get("viewBox"))
    if viewbox is None:
        return None
    vb_width, vb_height = viewbox
    if width:
        return width, width * vb_height / vb_width
    if height:
        return height * vb_width / vb_height, height
    return vb_width, vb_height


def _svg_display_size(
    svg_bytes: bytes, width: Length | None, height: Length | None
) -> tuple[Length, Length]:
    px_w, px_h = svg_intrinsic_size(svg_bytes) or _SVG_FALLBACK_PX
    if width is None and height is None:
        return Emu(max(1, round(px_w * _EMU_PER_PX))), Emu(max(1, round(px_h * _EMU_PER_PX)))
    if width is None:
        return Emu(max(1, round(int(height) * px_w / px_h))), Emu(int(height))
    if height is None:
        return Emu(int(width)), Emu(max(1, round(int(width) * px_h / px_w)))
    return Emu(int(width)), Emu(int(height))


def _placeholder_png(width: int, height: int) -> bytes:
    """Light-grey PNG with the given aspect ratio (fallback for SVG-unaware readers)."""
    longest = 256
    if width >= height:
        size = (longest, max(1, round(longest * height / width)))
    else:
        size = (max(1, round(longest * width / height)), longest)
    image = PILImage.new("RGB", size, (0xF2, 0xF2, 0xF2))
    if size[0] > 2 and size[1] > 2:
        pixels = image.load()
        border = (0xBF, 0xBF, 0xBF)
        for x in range(size[0]):
            pixels[x, 0] = pixels[x, size[1] - 1] = border
        for y in range(size[1]):
            pixels[0, y] = pixels[size[0] - 1, y] = border
    out = BytesIO()
    image.save(out, "PNG", dpi=(96, 96))
    return out.getvalue()


def _get_or_add_svg_part(package: Any, svg_bytes: bytes) -> ImagePart:
    sha1 = hashlib.sha1(svg_bytes).hexdigest()
    for part in package.image_parts:
        if part.content_type == _SVG_CONTENT_TYPE and part.sha1 == sha1:
            return part
    partname = package.image_parts._next_image_partname("svg")
    part = ImagePart(partname, _SVG_CONTENT_TYPE, svg_bytes)
    package.image_parts.append(part)
    return part


def _attach_svg_blip(blip: Any, svg_rid: str) -> None:
    ext_lst = blip.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = OxmlElement("a:extLst")
        blip.append(ext_lst)
    ext = OxmlElement("a:ext", {"uri": _SVG_BLIP_EXT_URI})
    ext_lst.append(ext)
    svg_blip = etree.SubElement(ext, f"{{{_ASVG_NS}}}svgBlip", nsmap={"asvg": _ASVG_NS})
    svg_blip.set(qn("r:embed"), svg_rid)


def add_svg_picture(
    run: Run,
    svg_bytes: bytes,
    width: Length | None = None,
    height: Length | None = None,
    fallback_png: bytes | None = None,
) -> InlineShape:
    """Embed an SVG the way Word 2016+ does: PNG blip + ``asvg:svgBlip`` extension.

    Size comes from ``width``/``height`` (one is enough; the other follows the
    aspect ratio) or the SVG's intrinsic size. Without ``fallback_png`` a
    light-grey placeholder of the right aspect is used for older readers.
    The SVG is stored re-serialized by :func:`sanitize_svg` (no DOCTYPE, no
    entity references); when that is impossible only the PNG is embedded.
    Raises ValueError for bytes that are not an SVG document.
    """
    sanitized = sanitize_svg(svg_bytes)  # raises ValueError for non-SVG input
    cx, cy = _svg_display_size(svg_bytes, width, height)
    png = fallback_png if fallback_png is not None else _placeholder_png(cx, cy)
    part = run.part
    inline = part.new_pic_inline(BytesIO(png), cx, cy)
    if sanitized is not None:
        svg_rid = part.relate_to(_get_or_add_svg_part(part.package, sanitized), RT.IMAGE)
        _attach_svg_blip(inline.xpath(".//a:blip")[0], svg_rid)
    run._r.add_drawing(inline)
    return InlineShape(inline)
