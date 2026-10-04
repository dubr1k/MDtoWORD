"""Convert Office Math Markup Language (OMML) back into LaTeX.

This is the inverse of :mod:`mdtoword.latex_omml`. The guiding property is a
round trip: for any formula that module produces, ``latex_to_omml(
omml_to_latex(omml))`` rebuilds a structurally identical OMML tree. That
decides several otherwise arbitrary choices:

* ``m:r`` text runs are mapped back character by character -- Greek letters
  and operators to their commands, LaTeX-special characters escaped -- and the
  run style picks the wrapper: ``m:nor`` is ``\\text{...}`` (or a function
  name such as ``\\sin`` when the text is one), ``m:sty`` ``b`` is
  ``\\mathbf``, ``bi`` is ``\\boldsymbol``, ``p`` is upright (``\\mathrm`` for
  Latin words).
* Composite elements (fractions, scripts, radicals, n-ary operators,
  delimiters, accents, matrices ...) each map to one LaTeX construct, with the
  matrix environments recognised from the delimiter that wraps them.
* An n-ary operator swallows the rest of its group when parsed back, so one
  that is followed by more content is braced to keep its operand where Word
  had it.

Documents written by Word itself use more of OMML than the forward converter
does (``m:func``, ``m:groupChr``, Unicode mathematical alphanumerics ...);
those are converted to the closest standard LaTeX. Anything not understood is
not dropped: its text is kept and a ``formula_partial`` warning names the
element, so the caller can report it.

Package layout:

* :mod:`.converter` -- the tree walker (traversal, text runs, warnings), the
  element dispatch table and the public functions;
* :mod:`.structures` -- handlers for fractions, scripts, radicals, n-ary
  operators, accents, bars, functions, limits, group characters and boxes;
* :mod:`.matrices` -- delimiters and the matrix, ``cases`` and equation-array
  environments hidden behind them;
* :mod:`.text` -- characters and run text to math-mode LaTeX;
* :mod:`.symbols` -- the character tables (Unicode to LaTeX commands);
* :mod:`.names` -- OMML/WordprocessingML tag names and OOXML helpers.
"""

from __future__ import annotations

from .converter import equation_rows, omml_to_latex
from .matrices import equation_environment

__all__ = ["equation_environment", "equation_rows", "omml_to_latex"]
