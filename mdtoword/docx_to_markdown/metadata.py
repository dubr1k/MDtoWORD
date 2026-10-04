"""Document properties: YAML front matter, and the title block it makes redundant."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .wordml import W_P, _plain_text

if TYPE_CHECKING:
    from .classify import _Classifier


def _yaml_quote(value: str) -> str:
    escaped = []
    for ch in value:
        if ch == "\\":
            escaped.append("\\\\")
        elif ch == '"':
            escaped.append('\\"')
        elif ch == "\n":
            escaped.append("\\n")
        elif ch == "\t":
            escaped.append("\\t")
        elif ch == "\r":
            escaped.append("\\r")
        elif ord(ch) < 0x20:
            escaped.append(f"\\x{ord(ch):02x}")
        else:
            escaped.append(ch)
    return '"' + "".join(escaped) + '"'


def front_matter(document: Any) -> str | None:
    """Title, author, subject, keywords and language as YAML, if any is set."""
    properties = document.core_properties
    entries = []
    for key, value in (("title", properties.title), ("author", properties.author),
                       ("subject", properties.subject),
                       ("keywords", properties.keywords),
                       ("lang", properties.language)):
        value = (value or "").strip()
        # python-docx's own template signs every document it creates.
        if not value or (key == "author" and value == "python-docx"):
            continue
        entries.append(f"{key}: {_yaml_quote(value)}")
    if not entries:
        return None
    return "\n".join(["---", *entries, "---"])


def without_title_block(elements: list[Any], document: Any,
                        classifier: _Classifier) -> list[Any]:
    """Drop a leading title block that only repeats the front matter.

    A title page made from metadata -- a ``Title`` paragraph holding the
    document title, then centred lines with the author -- is regenerated
    from the front matter on the way back, so keeping it would print the
    title twice.
    """
    properties = document.core_properties
    title = (properties.title or "").strip()
    author = (properties.author or "").strip()
    index = 0
    if (index < len(elements) and elements[index].tag == W_P and title
            and classifier.paragraph_info(elements[index]).title
            and _plain_text(elements[index]).strip() == title):
        index += 1
        while index < len(elements) and elements[index].tag == W_P:
            info = classifier.paragraph_info(elements[index])
            text = _plain_text(elements[index]).strip()
            if info.subtitle:
                break
            if info.centered and author and text == author:
                index += 1
                continue
            break
    return elements[index:]
