"""Inline collection: a paragraph's runs, links, fields, pictures and notes as inline items.

The walk descends into hyperlinks, simple fields, content controls,
tracked insertions and markup-compatibility choices; deleted text, hidden
runs and suppressed field results are skipped. Complex fields
(``w:fldChar`` begin/separate/end) may span runs and paragraphs, so the
stack of open fields lives here for the whole document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .equations import latex
from .fields import _Field, interpret_field
from .inline import _Link, _Raw, _Text
from .wordml import (
    _TRANSPARENT_INLINE,
    M_OMATH,
    M_OMATHPARA,
    MC_ALTERNATECONTENT,
    MC_CHOICE,
    MC_FALLBACK,
    R_ID,
    W_ANCHOR,
    W_BR,
    W_CHAR,
    W_CR,
    W_DRAWING,
    W_ENDNOTEREFERENCE,
    W_FLDCHAR,
    W_FLDCHARTYPE,
    W_FLDSIMPLE,
    W_FONT,
    W_FOOTNOTEREFERENCE,
    W_HYPERLINK,
    W_INSTR,
    W_INSTRTEXT,
    W_NOBREAKHYPHEN,
    W_OBJECT,
    W_PICT,
    W_PTAB,
    W_R,
    W_RPR,
    W_RUBY,
    W_RUBYBASE,
    W_SDT,
    W_SDTCONTENT,
    W_SYM,
    W_T,
    W_TAB,
    W_TOOLTIP,
    W_TYPE,
)

if TYPE_CHECKING:
    from .classify import _Classifier, _RunFormat
    from .diagnostics import _Diagnostics
    from .media import _Pictures
    from .notes import _Notes

# Symbol-font characters (w:sym with a code in the F0xx Private Use Area):
# the low byte is a position in that font, not a Unicode letter, so only
# fonts with a known layout are translated.
_SYMBOL_FONTS: dict[str, dict[int, str]] = {
    "symbol": {
        **{ord(latin): greek for latin, greek in zip(
            "abgdezhqiklmnxprstufcywjvJVGDQLXPSUFYW",
            "αβγδεζηθικλμνξπρστυφχψωϕϖϑςΓΔΘΛΞΠΣΥΦΨΩ",
            strict=True,
        )},
        0x22: "∀", 0x24: "∃", 0x27: "∋", 0x5E: "⊥", 0xA2: "′", 0xA3: "≤",
        0xA5: "∞", 0xAB: "↔", 0xAC: "←", 0xAD: "↑", 0xAE: "→", 0xAF: "↓",
        0xB0: "°", 0xB1: "±", 0xB2: "″", 0xB3: "≥", 0xB4: "×", 0xB5: "∝",
        0xB6: "∂", 0xB7: "•", 0xB8: "÷", 0xB9: "≠", 0xBA: "≡", 0xBB: "≈",
        0xBC: "…", 0xC4: "⊗", 0xC5: "⊕", 0xC6: "∅", 0xC7: "∩", 0xC8: "∪",
        0xC9: "⊃", 0xCA: "⊇", 0xCC: "⊂", 0xCD: "⊆", 0xCE: "∈", 0xCF: "∉",
        0xD0: "∠", 0xD1: "∇", 0xD6: "√", 0xD7: "⋅", 0xD8: "¬", 0xD9: "∧",
        0xDA: "∨", 0xDB: "⇔", 0xDC: "⇐", 0xDD: "⇑", 0xDE: "⇒", 0xDF: "⇓",
        0xE5: "∑", 0xF2: "∫",
    },
    # Check boxes and marks, as used for task lists.
    "wingdings": {0xA8: "☐", 0x6F: "☐", 0xFB: "✗", 0xFC: "✓", 0xFD: "☒", 0xFE: "☑"},
    "wingdings 2": {0xA3: "☐", 0x2A: "☐", 0x4F: "✗", 0x50: "✓", 0x52: "☑", 0x54: "☒"},
}


class _InlineCollector:
    """Turn paragraphs into inline items; holds the open fields of the document.

    Items are appended to the innermost *sink*: the paragraph's own list, or
    the children of the link (hyperlink or link field) being collected.
    """

    def __init__(self, classifier: _Classifier, diagnostics: _Diagnostics,
                 pictures: _Pictures, notes: _Notes, anchors: dict[str, str]) -> None:
        self.classifier = classifier
        self.diagnostics = diagnostics
        self.pictures = pictures
        self.notes = notes
        # Heading anchors by bookmark name (filled before any text is collected).
        self.anchors = anchors
        self.fields: list[_Field] = []
        # Heading tables of contents met so far; the block writer replaces
        # each with a "[TOC]" marker.
        self.toc_count = 0
        self._sinks: list[list[Any]] = []

    def anchor_target(self, name: str) -> str:
        return self.anchors.get(name, name)

    def suppressed(self) -> bool:
        return any(f.phase == "instr" or f.suppress for f in self.fields)

    def collect(self, paragraph: Any, part: Any) -> list[Any]:
        root: list[Any] = []
        self._sinks = [root]
        self._walk(paragraph, part, in_link=False)
        # A field hyperlink may not end in the paragraph it started in;
        # Markdown links cannot span paragraphs, so close it here.
        for open_field in self.fields:
            if open_field.link is not None:
                self._close_field_link(open_field)
        self._sinks = []
        return root

    def _emit(self, item: Any) -> None:
        if item is None or self.suppressed() or not self._sinks:
            return
        self._sinks[-1].append(item)

    def _walk(self, element: Any, part: Any, *, in_link: bool) -> None:
        for child in element:
            tag = child.tag
            if tag == W_R:
                self._run(child, part, in_link=in_link)
            elif tag == W_HYPERLINK:
                self._hyperlink(child, part)
            elif tag == W_FLDSIMPLE:
                self._simple_field(child, part, in_link=in_link)
            elif tag in _TRANSPARENT_INLINE:
                self._walk(child, part, in_link=in_link)
            elif tag == W_SDT:
                content = child.find(W_SDTCONTENT)
                if content is not None:
                    self._walk(content, part, in_link=in_link)
            elif tag in (M_OMATH, M_OMATHPARA):
                self._math(child)
            elif tag == MC_ALTERNATECONTENT:
                chosen = self._alternate(child)
                if chosen is not None:
                    self._walk(chosen, part, in_link=in_link)

    @staticmethod
    def _alternate(element: Any) -> Any:
        choice = element.find(MC_CHOICE)
        return choice if choice is not None else element.find(MC_FALLBACK)

    def _math(self, element: Any) -> None:
        if self.suppressed():
            return
        text = latex(self.diagnostics, element)
        if not text:
            return
        if element.tag == M_OMATHPARA:
            self._emit(_Raw(f"$${text}$$", text))
        else:
            self._emit(_Raw(f"${text}$", text))

    # -- runs -------------------------------------------------------------------

    def _run(self, run: Any, part: Any, *, in_link: bool) -> None:
        self._run_content(run, run, part, in_link=in_link, run_format=None)

    def _run_content(self, run: Any, container: Any, part: Any, *, in_link: bool,
                     run_format: _RunFormat | None) -> None:
        for child in container:
            tag = child.tag
            if tag == W_RPR:
                continue
            if tag == W_FLDCHAR:
                self._field_char(child)
                continue
            if tag == W_INSTRTEXT:
                if self.fields and self.fields[-1].phase == "instr":
                    self.fields[-1].instr.append(child.text or "")
                continue
            if self.suppressed():
                continue
            if run_format is None:
                linked = in_link or any(f.link is not None for f in self.fields)
                run_format = self.classifier.run_format(run, in_link=linked)
            if run_format.hidden:
                continue
            if tag == W_T:
                self._text(child.text or "", run_format)
            elif tag in (W_TAB, W_PTAB):
                self._text("\t", run_format)
            elif tag == W_BR:
                if child.get(W_TYPE) in (None, "textWrapping"):
                    self._text("\n", run_format)
            elif tag == W_CR:
                self._text("\n", run_format)
            elif tag == W_NOBREAKHYPHEN:
                self._text("-", run_format)
            elif tag == W_SYM:
                self._symbol(child, run_format)
            elif tag == W_DRAWING:
                for image in self.pictures.drawing(child, part):
                    self._emit(image)
            elif tag == W_PICT:
                for image in self.pictures.vml(child, part):
                    self._emit(image)
            elif tag == W_OBJECT:
                self.diagnostics.warn("object_unsupported")
            elif tag == W_FOOTNOTEREFERENCE:
                self._emit(self.notes.reference("footnote", child))
            elif tag == W_ENDNOTEREFERENCE:
                self._emit(self.notes.reference("endnote", child))
            elif tag == MC_ALTERNATECONTENT:
                chosen = self._alternate(child)
                if chosen is not None:
                    self._run_content(run, chosen, part, in_link=in_link,
                                      run_format=run_format)
            elif tag == W_RUBY:
                base = child.find(W_RUBYBASE)
                if base is not None:
                    self._walk(base, part, in_link=in_link)

    def _text(self, text: str, run_format: _RunFormat) -> None:
        if not text:
            return
        if run_format.caps:
            text = text.upper()
        self._emit(_Text(text, run_format.attrs, run_format.code))

    def _symbol(self, element: Any, run_format: _RunFormat) -> None:
        try:
            code = int(element.get(W_CHAR) or "", 16)
        except ValueError:
            return
        if 0xF000 <= code <= 0xF0FF:
            table = _SYMBOL_FONTS.get((element.get(W_FONT) or "").strip().lower(), {})
            character = table.get(code - 0xF000, "")
            if not character:
                self.diagnostics.warn("symbol_unsupported")
                return
        else:
            character = chr(code)
        self._text(character, run_format)

    def _hyperlink(self, element: Any, part: Any) -> None:
        url = ""
        relationship_id = element.get(R_ID)
        if relationship_id:
            relationship = part.rels.get(relationship_id)
            if relationship is not None:
                url = relationship.target_ref
        anchor = element.get(W_ANCHOR)
        if anchor:
            url = f"{url}#{self.anchor_target(anchor)}"
        if not url:
            self._walk(element, part, in_link=True)
            return
        link = _Link(url, element.get(W_TOOLTIP), [])
        self._sinks.append(link.children)
        try:
            self._walk(element, part, in_link=True)
        finally:
            if self._sinks and self._sinks[-1] is link.children:
                self._sinks.pop()
        self._emit(link)

    # -- fields -------------------------------------------------------------------

    def _field_char(self, element: Any) -> None:
        kind = element.get(W_FLDCHARTYPE)
        if kind == "begin":
            self.fields.append(_Field())
        elif kind == "separate" and self.fields:
            current = self.fields[-1]
            if current.phase == "instr":
                current.phase = "result"
                self._start_field(current)
        elif kind == "end" and self.fields:
            self._finish_field(self.fields.pop())

    def _simple_field(self, element: Any, part: Any, *, in_link: bool) -> None:
        simple = _Field(instr=[element.get(W_INSTR) or ""], phase="result")
        self.fields.append(simple)
        self._start_field(simple)
        try:
            self._walk(element, part, in_link=in_link)
        finally:
            if self.fields and self.fields[-1] is simple:
                self.fields.pop()
            self._finish_field(simple)

    def _start_field(self, current: _Field) -> None:
        meaning = interpret_field(current.instr, self.anchor_target)
        if meaning.suppress:
            current.suppress = True
        if meaning.warning:
            self.diagnostics.warn(meaning.warning)
        if meaning.table_of_contents:
            self.toc_count += 1
        if meaning.target:
            current.link = _Link(meaning.target, meaning.title, [])
            self._sinks.append(current.link.children)

    def _finish_field(self, current: _Field) -> None:
        if current.link is not None:
            self._close_field_link(current)

    def _close_field_link(self, current: _Field) -> None:
        link = current.link
        current.link = None
        if link is None:
            return
        for index in range(len(self._sinks) - 1, 0, -1):
            if self._sinks[index] is link.children:
                del self._sinks[index:]
                break
        else:
            return
        if link.children:
            self._emit(link)
