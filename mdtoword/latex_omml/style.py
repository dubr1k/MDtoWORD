"""The math alphabet runs are written in, and the commands that pick one."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class _Style:
    """The math alphabet runs are written in.

    `face` is None (math italic letters, the default), "bold" (`\\mathbf`:
    bold upright), "bolditalic" (`\\boldsymbol` / `\\bm`), "italic"
    (`\\mathit`) or "upright" (`\\mathrm`).  `script` is an OMML `m:scr`
    value -- "double-struck" for `\\mathbb` and so on -- or None.
    """

    face: Optional[str] = None
    script: Optional[str] = None

    @property
    def bold(self) -> bool:
        return self.face in ("bold", "bolditalic")


_PLAIN = _Style()

# Old-style declarations such as `{\bf x}` and `{\rm d}x`: each replaces
# the alphabet for the rest of its group, exactly like the matching
# `\math..` command would for its argument.
_FONT_SWITCHES = {
    "rm": _Style("upright"), "bf": _Style("bold"), "it": _Style("italic"),
    "cal": _Style(None, "script"), "frak": _Style(None, "fraktur"),
    "sf": _Style(None, "sans-serif"), "tt": _Style(None, "monospace"),
}

# `\mathbf{...}` and its fixed-alphabet siblings: the argument is set in
# exactly this alphabet, whatever the surrounding style.
_MATH_ALPHABETS = {
    "mathbf": _Style("bold"), "mathit": _Style("italic"),
    "mathbfit": _Style("bolditalic"),
    "mathrm": _Style("upright"), "mathup": _Style("upright"),
    "mathnormal": _PLAIN,
}
