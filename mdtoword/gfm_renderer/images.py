"""Images: fetching or reading them safely, sizing and embedding them.

Remote images are fetched only when allowed and only through
``safe_fetch`` with a size cap; UNC and protocol-relative targets never
touch the network unless fetching is allowed; local reads are confined to
the allowed image roots and checked *before* the filesystem is asked about
the path. Rasters are normalised and SVGs embedded natively, scaled to fit
the space their container leaves. Every failure falls back to the alt text
plus a warning.

Tests patch ``fetch_image`` and ``Path`` in this module, where they are
looked up.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import unquote

from docx.shared import Emu, Length
from PIL import Image

from ..ooxml import (
    add_svg_picture,
    normalize_raster,
    picture_size_to_fit,
    set_picture_description,
    svg_intrinsic_size,
)
from ..safe_fetch import RemoteFetchError, fetch_image
from .constants import _MAX_LOCAL_IMAGE_BYTES, _MAX_REMOTE_IMAGE_BYTES
from .helpers import _is_remote_target, _looks_like_svg, _plain_formatting
from .state import RendererState


class ImageMixin(RendererState):
    """Embed images into the current paragraph, or fall back to alt text."""

    def _append_image(self, token: Any, source_path: Path | None) -> bool:
        """Embed the image *token* points at into the current paragraph.

        Returns whether a picture was embedded; on any failure the alt text
        goes in instead, with a warning.
        """
        target = token.attrGet("src") or ""
        alt_text = token.content or "image"
        if self._paragraph is None:
            self._paragraph = self._new_paragraph()
        try:
            if _is_remote_target(target):
                if not self.allow_remote_images:
                    self._image_fallback(
                        alt_text,
                        f"Remote image not fetched: {target} (network access is disabled; "
                        "pass fetch_remote_images=true to allow it)",
                        "image_remote_disabled",
                    )
                    return False
                if target.lower().startswith(("http://", "https://")):
                    try:
                        image_bytes = fetch_image(target, max_bytes=_MAX_REMOTE_IMAGE_BYTES)
                    except RemoteFetchError as error:
                        reason = str(error)
                        if "large" in reason.lower() or "exceed" in reason.lower():
                            self._image_fallback(
                                alt_text,
                                f"Remote image too large: {target} (exceeds the "
                                f"{_MAX_REMOTE_IMAGE_BYTES}-byte limit; not embedded)",
                                "image_too_large",
                            )
                        else:
                            self._image_fallback(
                                alt_text,
                                f"Remote image could not be fetched: {target} ({reason})",
                                "image_fetch_failed",
                            )
                        return False
                    return self._insert_picture(image_bytes, target, token)
                # A UNC (``\\host\share``) or protocol-relative (``//host``)
                # target with fetching enabled is not fetched over HTTP -- it
                # falls through to the same filesystem check as any other
                # path, exactly as it did before this gate existed.
            if target.lower().startswith("data:"):
                self._image_fallback(alt_text, "Inline data: images are not supported", "image_unrenderable")
                return False

            if self._resolved_image_roots is not None and target.startswith(("//", "\\\\")):
                # resolve() on a UNC path already talks SMB to the host; with
                # roots enforced such a path can never be inside them anyway.
                self._image_fallback(
                    alt_text,
                    f"Image outside the allowed root: {target} (pass image_root=... to widen it)",
                    "image_outside_root",
                )
                return False
            image_path = Path(unquote(target)) if "%" in target else Path(target)
            if not image_path.is_absolute() and source_path is not None:
                image_path = source_path.parent / image_path
            if self._resolved_image_roots is not None:
                # Containment is checked here, before is_file() below ever
                # touches the filesystem for this path. Reversing the order
                # would leak the same existence oracle a prefix check would:
                # asking the filesystem about a path outside the sandbox is
                # itself the leak, not just what it returns.
                candidate = image_path.resolve()
                if not any(
                    candidate.is_relative_to(root) for root in self._resolved_image_roots
                ):
                    self._image_fallback(
                        alt_text,
                        f"Image outside the allowed root: {target} "
                        "(pass image_root=... to widen it)",
                        "image_outside_root",
                    )
                    return False
            if not image_path.is_file():
                raise FileNotFoundError(target)
            if image_path.stat().st_size > _MAX_LOCAL_IMAGE_BYTES:
                self._image_fallback(
                    alt_text,
                    f"Image too large: {target} (over {_MAX_LOCAL_IMAGE_BYTES} bytes; not embedded)",
                    "image_too_large",
                )
                return False
            return self._insert_picture(image_path.read_bytes(), target, token)
        except FileNotFoundError:
            self._image_fallback(alt_text, f"Image not found: {target}", "image_not_found")
        except Exception as error:
            self._image_fallback(
                alt_text,
                f"Image could not be rendered: {target} ({error or type(error).__name__})",
                "image_unrenderable",
            )
        return False

    def _image_limits(self) -> tuple[Length, Length]:
        # A picture exactly as wide as the text column still needs room for
        # the paragraph mark: LibreOffice then pushes it past the right
        # margin. One percent of slack keeps every viewer inside the margins.
        width = Emu(int(self._available_width()) * 99 // 100)
        if self._cell_context and self._table_columns:
            width = Emu(int(width) // max(1, self._table_columns))
        height = Emu(int(self._text_height * 0.8))
        return width, height

    def _insert_picture(self, data: bytes, target: str, token: Any) -> bool:
        max_width, max_height = self._image_limits()
        requested = token.attrGet("width") if hasattr(token, "attrGet") else None
        if requested and str(requested).rstrip("px").isdigit():
            max_width = Emu(min(int(max_width), int(int(str(requested).rstrip("px")) * 9525)))
        run = self._paragraph.add_run()
        if self._link is not None:
            self._link.append(run._r)
        if _looks_like_svg(target, data):
            size = svg_intrinsic_size(data) or (300.0, 150.0)
            width, height = picture_size_to_fit(
                (round(size[0]), round(size[1])), (96, 96), max_width, max_height
            )
            shape = add_svg_picture(run, data, width, height)
        else:
            data, _ = normalize_raster(data)
            with Image.open(BytesIO(data)) as image:
                pixels = image.size
                dpi = image.info.get("dpi") or (96, 96)
            width, height = picture_size_to_fit(pixels, dpi, max_width, max_height)
            shape = run.add_picture(BytesIO(data), width=width, height=height)
        # python-docx omits the wrap distances; Word reads that as zero but
        # LibreOffice falls back to ~3 mm and shifts the picture sideways.
        for side in ("distT", "distB", "distL", "distR"):
            shape._inline.set(side, "0")
        description = token.content or ""
        title = token.attrGet("title")
        if description or title:
            set_picture_description(shape, description, title)
        return True

    def _image_fallback(self, alt_text: str, warning: str, code: str = "image_unrenderable") -> None:
        """Record *warning* and insert ``[alt_text]`` in place of the image.

        Every path that gives up on embedding an image -- a disabled or
        skipped remote fetch, an oversized response, a missing file, or any
        other render failure -- ends up here, so the fallback text and its
        warning are appended together in exactly one place instead of being
        repeated at each call site.
        """
        self._warn(warning, code)
        self._append_text(f"[{alt_text}]", _plain_formatting(), self._link)
