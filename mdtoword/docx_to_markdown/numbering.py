"""Word list numbering: level definitions, the running counters and printed labels.

``w:numPr`` (direct or from the style) names a list and a level; the
abstract numbering gives that level's format, start and label text, with
``w:startOverride`` honoured and list styles (``w:numStyleLink``) followed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .styles import _Styles
from .wordml import (
    W_ABSTRACTNUM,
    W_ABSTRACTNUMID,
    W_ILVL,
    W_IND,
    W_LEFT,
    W_LVL,
    W_LVLOVERRIDE,
    W_LVLTEXT,
    W_NUM,
    W_NUMFMT,
    W_NUMID,
    W_NUMPR,
    W_NUMSTYLELINK,
    W_PPR,
    W_START,
    W_STARTOVERRIDE,
    W_VAL,
    _int,
    _on,
    _q,
)


def _format_number(value: int, number_format: str) -> str:
    """A list counter in one of Word's numbering formats."""
    if number_format in ("lowerLetter", "upperLetter"):
        letters = ""
        number = max(value, 1)
        while number > 0:
            number, remainder = divmod(number - 1, 26)
            letters = chr(ord("a") + remainder) + letters
        return letters.upper() if number_format == "upperLetter" else letters
    if number_format in ("lowerRoman", "upperRoman"):
        numerals = (
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
            (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"),
            (4, "iv"), (1, "i"),
        )
        number, roman = max(value, 1), ""
        for amount, numeral in numerals:
            while number >= amount:
                roman += numeral
                number -= amount
        return roman.upper() if number_format == "upperRoman" else roman
    if number_format in ("russianLower", "russianUpper"):
        alphabet = "абвгдежзиклмнопрстуфхцчшщэюя"
        letter = alphabet[(max(value, 1) - 1) % len(alphabet)]
        return letter.upper() if number_format == "russianUpper" else letter
    if number_format == "decimalZero":
        return f"{value:02d}"
    if number_format in ("bullet", "none"):
        return ""
    return str(value)


@dataclass
class _Level:
    num_id: str
    ilvl: int
    number_format: str
    start: int
    text: str
    indent: int | None
    legal: bool = False

    @property
    def ordered(self) -> bool:
        return self.number_format != "bullet"


class _Numbering:
    """Word list numbering: level definitions plus the running counters."""

    def __init__(self, element: Any, styles: _Styles) -> None:
        self._styles = styles
        self._abstract: dict[str, Any] = {}
        self._nums: dict[str, tuple[str, dict[int, tuple[int | None, Any]]]] = {}
        self._counters: dict[str, list[int | None]] = {}
        self._levels: dict[tuple[str, int], _Level | None] = {}
        if element is None:
            return
        for abstract in element.iterchildren(W_ABSTRACTNUM):
            self._abstract[abstract.get(W_ABSTRACTNUMID) or ""] = abstract
        for num in element.iterchildren(W_NUM):
            reference = num.find(W_ABSTRACTNUMID)
            if reference is None:
                continue
            overrides: dict[int, tuple[int | None, Any]] = {}
            for override in num.iterchildren(W_LVLOVERRIDE):
                start = override.find(W_STARTOVERRIDE)
                overrides[_int(override.get(W_ILVL))] = (
                    _int(start.get(W_VAL)) if start is not None else None,
                    override.find(W_LVL),
                )
            self._nums[num.get(W_NUMID) or ""] = (reference.get(W_VAL) or "", overrides)

    def _abstract_of(self, num_id: str, depth: int = 0) -> Any:
        entry = self._nums.get(num_id)
        if entry is None:
            return None
        abstract = self._abstract.get(entry[0])
        if abstract is None:
            return None
        link = abstract.find(W_NUMSTYLELINK)
        if link is not None and depth < 5:
            # A list style: the real levels hang off the style's own list.
            for style in self._styles.chain(link.get(W_VAL)):
                num_pr = style.ppr.find(W_NUMPR) if style.ppr is not None else None
                linked = num_pr.find(W_NUMID) if num_pr is not None else None
                if linked is not None and linked.get(W_VAL) != num_id:
                    resolved = self._abstract_of(linked.get(W_VAL) or "", depth + 1)
                    if resolved is not None:
                        return resolved
        return abstract

    def fresh(self) -> _Numbering:
        """The same definitions with every counter reset."""
        copy = object.__new__(_Numbering)
        copy.__dict__.update(self.__dict__)
        copy._counters = {}
        return copy

    def level(self, num_id: str, ilvl: int) -> _Level | None:
        key = (num_id, ilvl)
        if key in self._levels:
            return self._levels[key]
        level = self._build_level(num_id, ilvl)
        self._levels[key] = level
        return level

    def _build_level(self, num_id: str, ilvl: int) -> _Level | None:
        entry = self._nums.get(num_id)
        abstract = self._abstract_of(num_id)
        if entry is None or abstract is None:
            return None
        start_override, level_override = entry[1].get(ilvl, (None, None))
        definition = level_override
        if definition is None:
            for candidate in abstract.iterchildren(W_LVL):
                if _int(candidate.get(W_ILVL)) == ilvl:
                    definition = candidate
                    break
        if definition is None:
            return None
        number_format = definition.find(W_NUMFMT)
        start = definition.find(_q("w:start"))
        text = definition.find(W_LVLTEXT)
        indent = None
        ppr = definition.find(W_PPR)
        if ppr is not None and ppr.find(W_IND) is not None:
            ind = ppr.find(W_IND)
            value = ind.get(W_LEFT) or ind.get(W_START)
            indent = _int(value) if value is not None else None
        return _Level(
            num_id=num_id,
            ilvl=ilvl,
            number_format=(number_format.get(W_VAL) if number_format is not None else None)
            or "decimal",
            start=start_override if start_override is not None
            else (_int(start.get(W_VAL)) if start is not None else 0),
            text=(text.get(W_VAL) if text is not None else "") or "",
            indent=indent,
            legal=_on(definition.find(_q("w:isLgl"))),
        )

    def advance(self, level: _Level) -> int:
        """Count one more paragraph at ``level``; return its number."""
        counters = self._counters.setdefault(level.num_id, [None] * 9)
        index = min(max(level.ilvl, 0), 8)
        current = counters[index]
        counters[index] = level.start if current is None else current + 1
        for deeper in range(index + 1, 9):
            counters[deeper] = None
        return counters[index] or 0

    def label(self, level: _Level) -> str:
        """The number as Word prints it, e.g. ``1.2.`` from ``%1.%2.``."""
        counters = self._counters.get(level.num_id, [None] * 9)

        def substitute(match: re.Match[str]) -> str:
            index = int(match.group(1)) - 1
            other = self.level(level.num_id, index)
            value = counters[index] if index < 9 else None
            if value is None:
                value = other.start if other is not None else 1
            # Legal numbering ("1.1.1") prints every level as a number.
            number_format = other.number_format if other is not None else "decimal"
            if level.legal and number_format != "bullet":
                number_format = "decimal"
            return _format_number(value, number_format)

        return re.sub(r"%([1-9])", substitute, level.text).strip()
