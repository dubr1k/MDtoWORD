"""Paragraph and character styles, looked up by id with ``basedOn`` chains resolved."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .wordml import W_BASEDON, W_DEFAULT, W_NAME, W_PPR, W_RPR, W_STYLE, W_STYLEID, W_TYPE, W_VAL


@dataclass
class _Style:
    style_id: str
    name: str
    kind: str
    based_on: str | None
    ppr: Any
    rpr: Any


class _Styles:
    """Style lookup by id, with ``basedOn`` chains resolved and cached."""

    def __init__(self, styles_element: Any) -> None:
        self._by_id: dict[str, _Style] = {}
        self.default_paragraph: str | None = None
        self._chains: dict[str, list[_Style]] = {}
        if styles_element is None:
            return
        for element in styles_element.iterchildren(W_STYLE):
            style_id = element.get(W_STYLEID) or ""
            name_element = element.find(W_NAME)
            name = name_element.get(W_VAL, "") if name_element is not None else style_id
            based = element.find(W_BASEDON)
            kind = element.get(W_TYPE) or "paragraph"
            self._by_id[style_id] = _Style(
                style_id, (name or style_id).strip().lower(), kind,
                based.get(W_VAL) if based is not None else None,
                element.find(W_PPR), element.find(W_RPR),
            )
            if kind == "paragraph" and element.get(W_DEFAULT) in ("1", "true", "on"):
                self.default_paragraph = style_id

    def chain(self, style_id: str | None) -> list[_Style]:
        if not style_id:
            return []
        cached = self._chains.get(style_id)
        if cached is not None:
            return cached
        chain: list[_Style] = []
        seen: set[str] = set()
        current = self._by_id.get(style_id)
        while current is not None and current.style_id not in seen and len(chain) < 20:
            chain.append(current)
            seen.add(current.style_id)
            current = self._by_id.get(current.based_on or "")
        self._chains[style_id] = chain
        return chain
