"""The ``\\command`` dispatch table.

`_parse_command` looks the command name up in `_HANDLERS` and calls the
handler it finds.  Every handler takes ``(tokens, index, name, style,
stop)`` -- `index` already past the command token, `style` the current
`_Style`, `stop` the terminator of the enclosing sequence -- and returns
``(elements, index)`` like every other parse function.  The handlers live
in the ``commands_*`` modules, grouped by concern, and in
:mod:`.environments`.

`_REGISTRY` lists the command families in precedence order: where two
families share a name, the earlier entry wins, exactly as the first
matching branch of a chain of ``if name in ...`` tests would.  Constructs
come first, then the refusals for commands met out of place, and the
plain character tables last.
"""

from __future__ import annotations

from typing import Callable, Iterable, NoReturn

from . import (
    commands_alphabets as alphabets,
    commands_decorations as decorations,
    commands_delimiters as delimiters,
    commands_operators as operators,
    commands_spacing as spacing,
    environments,
)
from .errors import UnsupportedLatexError
from .style import _FONT_SWITCHES, _MATH_ALPHABETS, _PLAIN, _Style
from .symbols import _ESCAPED, _ITALIC_SYMBOLS, _SPACING, _SYMBOLS
from .tables import (
    _ACCENTS,
    _BIG_DELIMITERS,
    _BINOMIALS,
    _BOLD_ITALIC_STYLE,
    _CANCELS,
    _ENVIRONMENT_ONLY,
    _EXTENSIBLE_ARROWS,
    _FRACTIONS,
    _GROUP_CHARACTERS,
    _INFIX,
    _LIMIT_OPERATORS,
    _MATH_CLASSES,
    _NARY,
    _NO_OP_SWITCHES,
    _NUMBERING,
    _PHANTOMS,
    _SCRIPT_ALPHABETS,
    _TEXT_COMMANDS,
    _UPRIGHT_FUNCTIONS,
)
from .tokenizer import Stop

# (tokens, index past the command, command name, style, stop)
#   -> (elements, index after the command and its arguments)
Handler = Callable[[list, int, str, _Style, Stop], tuple]


# Declarations and numbering are handled by `_parse_lines`, which sees
# them wherever they can act on the rest of a sequence.  Reaching them
# here means they turned up as the argument of some command or script
# -- `x^\displaystyle`, `\frac\tag{1}` -- where there is nothing for
# them to act on.


def _refuse_numbering(tokens: list, index: int, name: str, style: _Style,
                      stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        f"\\{name} belongs to a whole formula line, not to the argument "
        "of another command; move it to the end of the line"
    )


def _refuse_declaration(tokens: list, index: int, name: str, style: _Style,
                        stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        f"\\{name} cannot stand in for an argument here; put it inside "
        f"the braces, as in {{\\{name} ...}}"
    )


# `_parse_lines` intercepts both of these wherever they can carry
# meaning, so reaching them here means they turned up somewhere that
# takes a single atom -- `x^\\`, `\frac\over x` -- where TeX has nothing
# to attach them to either. Say which half is missing rather than
# falling through to the generic "unsupported command" below.


def _refuse_line_break(tokens: list, index: int, name: str, style: _Style,
                       stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        "A line break \\\\ has nothing to break here"
    )


def _refuse_infix(tokens: list, index: int, name: str, style: _Style,
                  stop: Stop) -> NoReturn:
    raise UnsupportedLatexError(
        f"\\{name} needs an expression on both sides of it"
    )


_REGISTRY: tuple[tuple[Iterable[str], Handler], ...] = (
    # Constructs.
    (_FRACTIONS, operators._parse_fraction),
    (("sqrt",), operators._parse_sqrt),
    (_TEXT_COMMANDS, alphabets._parse_text),
    (("operatorname",), operators._parse_operatorname),
    (("mathop",), operators._parse_mathop),
    (_MATH_CLASSES, alphabets._parse_math_class),
    (_MATH_ALPHABETS, alphabets._parse_math_alphabet),
    (_SCRIPT_ALPHABETS, alphabets._parse_script_alphabet),
    (_BOLD_ITALIC_STYLE, alphabets._parse_bold_italic),
    (_NARY, operators._parse_nary),
    (_LIMIT_OPERATORS, operators._parse_limit_operator),
    (_ACCENTS, decorations._parse_accent),
    (("overline", "underline"), decorations._parse_overline),
    (_GROUP_CHARACTERS, decorations._parse_group_character),
    (_EXTENSIBLE_ARROWS, decorations._parse_extensible_arrow),
    (("overset", "stackrel", "underset"), decorations._parse_overset),
    (_BINOMIALS, operators._parse_binomial),
    (("substack",), operators._parse_substack),
    (("left",), delimiters._parse_left),
    (("right",), delimiters._refuse_right),
    (("middle",), delimiters._refuse_middle),
    (_BIG_DELIMITERS, delimiters._parse_big_delimiter),
    (("begin",), environments._parse_begin),
    (("end",), environments._refuse_end),
    (("boxed",), decorations._parse_boxed),
    (("fbox", "framebox"), decorations._parse_fbox),
    (_CANCELS, decorations._parse_cancel),
    (_PHANTOMS, decorations._parse_phantom),
    (("smash",), decorations._parse_smash),
    (("mathstrut", "strut"), decorations._parse_strut),
    (("hspace",), spacing._parse_hspace),
    (("kern", "mkern", "hskip", "mskip", "mspace"), spacing._parse_kern),
    (("textcolor",), spacing._parse_textcolor),
    (("bmod",), operators._parse_bmod),
    (("pmod", "pod", "mod"), operators._parse_mod),
    (("not",), operators._parse_negation),
    (("ket",), delimiters._parse_ket),
    (("bra",), delimiters._parse_bra),
    (("braket",), delimiters._parse_braket),
    # Commands that only mean something somewhere else.
    (("sideset",), operators._refuse_sideset),
    (("hline", "cline", "vline", "hdashline"), environments._refuse_ruled_line),
    (("limits", "nolimits", "displaylimits"), operators._refuse_limits),
    (_NUMBERING, _refuse_numbering),
    ((*_NO_OP_SWITCHES, *_FONT_SWITCHES, "color"), _refuse_declaration),
    (("\\",), _refuse_line_break),
    (_INFIX, _refuse_infix),
    # Checked before the remaining fallback tables (_SYMBOLS, _SPACING,
    # _ESCAPED, _UPRIGHT_FUNCTIONS) -- not just "the symbol table" -- so a
    # construct this package handles only in another form never silently
    # degrades into something that merely looks plausible. Everything above
    # this point is a construct entry for something already implemented;
    # _ENVIRONMENT_ONLY only needs to stay disjoint from the tables it
    # precedes.
    (_ENVIRONMENT_ONLY, environments._refuse_environment_only),
    # Plain characters and named functions.
    (_SYMBOLS, alphabets._parse_symbol),
    (_ITALIC_SYMBOLS, alphabets._parse_italic_symbol),
    (_SPACING, spacing._parse_spacing),
    (_ESCAPED, alphabets._parse_escaped),
    (_UPRIGHT_FUNCTIONS, operators._parse_upright_function),
)


def _dispatch_table(
        registry: tuple[tuple[Iterable[str], Handler], ...]) -> dict[str, Handler]:
    """Flatten `registry` into one name -> handler map; the first entry
    that names a command keeps it."""
    table: dict[str, Handler] = {}
    for names, handler in registry:
        for name in names:
            table.setdefault(name, handler)
    return table


_HANDLERS = _dispatch_table(_REGISTRY)


def _parse_command(tokens: list, index: int, style: _Style = _PLAIN,
                   stop: Stop = None) -> tuple:
    """Parse one `\\command` and its arguments. Returns (elements, index)."""
    name = tokens[index][1][1:]
    handler = _HANDLERS.get(name)
    if handler is None:
        raise UnsupportedLatexError(f"Unsupported LaTeX command: \\{name}")
    return handler(tokens, index + 1, name, style, stop)
