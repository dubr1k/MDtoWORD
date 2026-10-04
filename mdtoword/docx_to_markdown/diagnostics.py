"""Warnings about what Markdown cannot represent, counted per kind for one document."""

from __future__ import annotations

import re

from ..errors import ConversionWarning

_WARNING_MESSAGES = {
    "table_merged_cells": "Merged table cells were split: Markdown tables have no "
                          "row or column spans ({count} table(s)).",
    "nested_table_flattened": "Tables nested inside table cells were flattened "
                              "to text ({count}).",
    "image_unsupported": "Drawings that are not pictures (shapes, charts, "
                         "SmartArt) were skipped ({count}).",
    "image_not_extracted": "Images were not extracted; only their alt text "
                           "was kept ({count}).",
    "image_missing": "Images whose data is missing from the document were "
                     "skipped ({count}).",
    "textbox_skipped": "Text boxes were skipped ({count}).",
    "object_unsupported": "Embedded objects (OLE, e.g. MathType equations or "
                          "spreadsheets) were skipped ({count}).",
    "toc_skipped": "A table of figures, or a table of contents that could not "
                   "be placed as a whole, was skipped ({count}).",
    "index_skipped": "A generated index or table of authorities was skipped "
                     "({count}).",
    "heading_level_clamped": "Headings deeper than level 6 were written as bold "
                             "paragraphs ({count}).",
    "footnote_missing": "Footnote references without a footnote text "
                        "({count}).",
    "symbol_unsupported": "Symbol-font characters without a Unicode "
                          "equivalent were skipped ({count}).",
    "altchunk_skipped": "Embedded alternative-format content (altChunk) was "
                        "skipped ({count}).",
}


class _Diagnostics:
    """Warning counts by code, reported once per kind, in order of first occurrence."""

    def __init__(self) -> None:
        self._warnings: dict[str, int] = {}
        self._formula_elements: dict[str, int] = {}
        self._formula_count = 0

    def warn(self, code: str, count: int = 1) -> None:
        self._warnings[code] = self._warnings.get(code, 0) + count

    def formula_warnings(self, found: list[str]) -> None:
        """Count the ``formula_partial`` warnings of one formula's conversion."""
        for warning in found:
            self._formula_count += 1
            for name in re.findall(r"m:(\w+)", str(warning)):
                self._formula_elements[name] = self._formula_elements.get(name, 0) + 1

    def warnings(self) -> list[ConversionWarning]:
        result = []
        for code, count in self._warnings.items():
            template = _WARNING_MESSAGES.get(code, code.replace("_", " ") + " ({count})")
            result.append(ConversionWarning(template.format(count=count), code=code))
        if self._formula_count:
            names = ", ".join(f"m:{name}" for name in sorted(self._formula_elements))
            result.append(ConversionWarning(
                f"Equation elements not understood were kept as plain text in "
                f"{self._formula_count} formula(s): {names}.",
                code="formula_partial",
            ))
        return result
