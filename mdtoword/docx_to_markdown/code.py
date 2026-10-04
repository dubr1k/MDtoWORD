"""Code recognition: monospace fonts and code styles, verbatim text, fences.

A paragraph is code when its style says so -- by name or by a monospace
font -- or when every run carrying text is monospace. Consecutive code
paragraphs become one fenced block; an italic one-word paragraph right
above one names its language.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from .wordml import (
    W_ASCII,
    W_BR,
    W_CR,
    W_HANSI,
    W_NOBREAKHYPHEN,
    W_PTAB,
    W_RFONTS,
    W_T,
    W_TAB,
    W_TYPE,
    _plain_text,
    has_content,
    text_runs,
)

if TYPE_CHECKING:
    from .classify import _Classifier, _ParaInfo
    from .styles import _Style

_MONOSPACE_FONTS = frozenset({
    "courier new", "courier", "consolas", "menlo", "monaco", "lucida console",
    "lucida sans typewriter", "source code pro", "fira code", "fira mono",
    "jetbrains mono", "dejavu sans mono", "liberation mono", "roboto mono",
    "ubuntu mono", "sf mono", "cascadia code", "cascadia mono", "pt mono",
    "andale mono", "inconsolata", "noto sans mono", "ibm plex mono", "hack",
    "nimbus mono", "freemono", "courier std", "osaka-mono", "ms gothic",
})
_MONOSPACE_WORD = re.compile(r"\b(?:mono|monospace|code|courier)\b", re.IGNORECASE)
_CODE_STYLE = re.compile(
    r"\b(?:code|source|preformatted|verbatim|plain text)\b", re.IGNORECASE
)
_CODE_CHAR_STYLE = re.compile(r"\b(?:code|verbatim)\b", re.IGNORECASE)
_CAPTION_WORD = re.compile(r"[\w+#.-]{1,40}")


def _is_monospace(font: str | None) -> bool:
    if not font:
        return False
    name = font.strip().lower()
    return name in _MONOSPACE_FONTS or bool(_MONOSPACE_WORD.search(name))


def is_code_paragraph(classifier: _Classifier, paragraph: Any, chain: list[_Style],
                      style_id: str | None) -> bool:
    """Whether a paragraph with this style chain is a line of code."""
    if any(_CODE_STYLE.search(style.name) for style in chain):
        return True
    if chain and style_id != classifier.styles.default_paragraph:
        for style in chain:
            fonts = style.rpr.find(W_RFONTS) if style.rpr is not None else None
            font = (fonts.get(W_ASCII) or fonts.get(W_HANSI)) if fonts is not None else None
            if font:
                if _is_monospace(font):
                    return True
                break
    if has_content(paragraph):
        return False
    seen_text = False
    for run in text_runs(paragraph):
        text = "".join(t.text or "" for t in run.iter(W_T))
        if not text.strip():
            continue
        if not classifier.run_format(run, in_link=False).code:
            return False
        seen_text = True
    return seen_text


def code_text(paragraph: Any) -> str:
    """A code paragraph's text, verbatim: tabs and line breaks kept."""
    parts: list[str] = []
    for run in text_runs(paragraph):
        for child in run:
            tag = child.tag
            if tag == W_T:
                parts.append(child.text or "")
            elif tag in (W_TAB, W_PTAB):
                parts.append("\t")
            elif (tag == W_BR and child.get(W_TYPE) in (None, "textWrapping")) or tag == W_CR:
                parts.append("\n")
            elif tag == W_NOBREAKHYPHEN:
                parts.append("-")
    return "".join(parts)


def fenced_code(lines: list[str], language: str) -> str:
    """A fenced block of the code lines, blank lines at either end dropped.

    The fence is longer than any backtick run inside the code.
    """
    lines = list(lines)
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    content = "\n".join(lines)
    longest = max((len(run) for run in re.findall(r"`{3,}", content)), default=2)
    fence = "`" * max(3, longest + 1)
    return f"{fence}{language}\n{content}\n{fence}"


def is_language_caption(classifier: _Classifier, element: Any, info: _ParaInfo) -> bool:
    """One italic word: names the language when a code block follows."""
    if info.heading or info.level is not None or info.quote:
        return False
    text = _plain_text(element).strip()
    if not _CAPTION_WORD.fullmatch(text):
        return False
    runs = [run for run in text_runs(element)
            if "".join(t.text or "" for t in run.iter(W_T)).strip()]
    return bool(runs) and all(
        "italic" in classifier.run_format(run, in_link=False).attrs for run in runs
    )
