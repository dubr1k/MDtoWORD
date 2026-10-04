"""The error raised for LaTeX that has no OMML equivalent here."""

from __future__ import annotations


class UnsupportedLatexError(ValueError):
    """Raised when a LaTeX construct has no OMML equivalent here."""

    # Reported under the package it is raised from and re-exported by --
    # `mdtoword.latex_omml.UnsupportedLatexError` -- rather than under this
    # private submodule, so tracebacks and pickles name it as they always did.
    __module__ = "mdtoword.latex_omml"
