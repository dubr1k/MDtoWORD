"""Input preparation and the pre-passes that run before any block renders.

The raw text is normalised before markdown-it sees it, footnote
definitions the parser silently drops are reported, the document language
is resolved (option, front matter, or detection), and every heading's
bookmark is computed up front so links and the table of contents can
point forward.
"""

from __future__ import annotations

from collections.abc import Sequence
import re
from typing import Any

from ..ooxml import bookmark_name, detect_language, github_slug
from .constants import _LANGUAGE_TAG, _XML_INVALID
from .helpers import _inline_plain_text
from .state import RendererState


class SourceMixin(RendererState):
    """Normalise the input and run the pre-passes over the token stream."""

    def _prepare_source(self, markdown: str) -> str:
        """Normalise the raw text before markdown-it sees it.

        - A leading byte-order mark (Windows editors write one) would turn
          the first heading or the front matter into plain text.
        - Characters XML 1.0 forbids (ESC from pasted terminal output, form
          feeds, other C0 controls) cannot be written to a .docx at all;
          they are removed with one warning instead of failing the file.
        - markdown-it-py 4.0 raises IndexError on some constructs that end
          the input without a newline (a quoted table, for one); a final
          newline changes nothing else.
        """
        if markdown.startswith("\ufeff"):
            markdown = markdown[1:]
        cleaned, removed = _XML_INVALID.subn("", markdown)
        if removed:
            self._warn(
                f"Removed {removed} control character(s) that a Word document cannot contain",
                "control_characters_removed",
                line=markdown.count("\n", 0, _XML_INVALID.search(markdown).start()) + 1,
            )
        if cleaned and not cleaned.endswith("\n"):
            cleaned += "\n"
        return cleaned

    def _warn_unreferenced_footnotes(self, markdown: str, environment: dict[str, Any]) -> None:
        """The footnote plugin silently drops definitions nothing refers to."""
        references = (environment.get("footnotes") or {}).get("refs") or {}
        for key, footnote_id in references.items():
            if footnote_id != -1:
                continue
            label = key.removeprefix(":")
            match = re.search(rf"^[ \t]*\[\^{re.escape(label)}\]:", markdown, re.MULTILINE)
            line = markdown.count("\n", 0, match.start()) + 1 if match else None
            self._warn(
                f"Footnote [^{label}] is defined but never referenced; it is not in the document",
                "footnote_unreferenced",
                line=line,
            )

    def _resolve_language(self, markdown: str) -> str:
        if self.options.language != "auto":
            return self.options.language
        declared = self._front_matter.get("lang") or self._front_matter.get("language")
        if isinstance(declared, str) and declared.strip():
            tag = declared.strip().replace("_", "-")
            aliases = {"ru": "ru-RU", "rus": "ru-RU", "russian": "ru-RU", "русский": "ru-RU",
                       "en": "en-US", "eng": "en-US", "english": "en-US"}
            tag = aliases.get(tag.lower(), tag)
            if _LANGUAGE_TAG.match(tag):
                return tag
            self._warn(
                f"Front matter language {declared.strip()!r} is not a language tag such as "
                "'ru-RU'; the language was detected from the text instead",
                "front_matter_ignored",
                line=1,
            )
        return detect_language(markdown)

    def _prepare_heading_anchors(self, tokens: Sequence[Any]) -> None:
        """Compute every heading's bookmark up front so links can point forward."""
        used_slugs: dict[str, int] = {}
        used_names: set[str] = set()
        footnote_depth = 0
        for index, token in enumerate(tokens):
            if token.type == "footnote_open":
                footnote_depth += 1
            elif token.type == "footnote_close":
                footnote_depth -= 1
            if token.type != "heading_open" or index + 1 >= len(tokens):
                continue
            text = _inline_plain_text(tokens[index + 1].children or []).strip()
            slug = github_slug(text, used_slugs)
            name = bookmark_name(slug or "section", used_names)
            self._slug_to_bookmark.setdefault(slug, name)
            # Keyed by token position, not consumed in order: a heading in a
            # skipped footnote must not shift every later heading's bookmark.
            self._heading_bookmark_at[index] = name
            if not footnote_depth:
                self._headings.append((min(int(token.tag[1:]), 9), text, name))
