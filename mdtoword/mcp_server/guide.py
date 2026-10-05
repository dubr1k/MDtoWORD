"""Ресурс с руководством для агента и промпт prepare_markdown_for_word.

Руководство лежит в корне пакета ``mdtoword`` как agent_guide.md (package data).
"""

from __future__ import annotations

from importlib import resources

from .server import GUIDE_URI, mcp


def _guide_text() -> str:
    return (
        resources.files("mdtoword").joinpath("agent_guide.md").read_text(encoding="utf-8")
    )


@mcp.resource(
    GUIDE_URI,
    name="markdown_guide",
    title="Writing Markdown that converts cleanly to Word",
    description=(
        "Supported Markdown features, how each maps to Word, and the "
        "recommended check-convert-verify workflow for this server"
    ),
    mime_type="text/markdown",
)
def markdown_guide() -> str:
    """Руководство для агента, лежит в пакете как agent_guide.md."""
    return _guide_text()


@mcp.prompt(
    name="prepare_markdown_for_word",
    title="Prepare Markdown for Word",
    description=(
        "Write or fix a Markdown document so it converts to Word without "
        "losses, then convert and verify it"
    ),
)
def prepare_markdown_for_word(topic_or_path: str = "") -> str:
    """Порядок работы над документом плюс само руководство целиком.

    Руководство вложено в текст, а не дано ссылкой на ресурс: не каждый
    клиент даёт модели читать ресурсы.
    """
    target = topic_or_path.strip()
    if target:
        task = (
            f"Target: {target}\n"
            "If this is the path of an existing Markdown file, verify and fix "
            "it in place. Otherwise write a new Markdown document on this topic "
            "and save it as a .md file next to its images."
        )
    else:
        task = "Target: the Markdown document we are working on."
    steps = (
        "Prepare a Markdown document for conversion to Word with the mdtoword tools.\n\n"
        f"{task}\n\n"
        "Steps:\n"
        "1. Write or check the Markdown against the guide below: use only "
        "the supported constructs, keep images as local paths relative to "
        "the .md file, and put each display formula in $$…$$ or an amsmath "
        "environment.\n"
        "2. Collect every formula and run `check_latex` on them. Rewrite any "
        "formula with ok=false using supported LaTeX and check it again.\n"
        "3. Run `preview_markdown` on the file. Fix every warning at its "
        "reported `line` (use `code` to see what kind of problem it is) and "
        "preview again until no warnings remain or the remaining ones are "
        "intended.\n"
        "4. Run `markdown_to_word` with the options the reader needs (for "
        "example preset=\"gost\" for a Russian GOST 7.32 report, or "
        "preset=\"gost_user\" for the custom 12 pt adaptation without "
        "headers/footers, toc=true for "
        "a table of contents).\n"
        "5. Run `inspect_docx` on the output and confirm the outline, the "
        "equation/table/image counts and the page setup match what you "
        "expect. If something is missing, fix the Markdown and repeat from "
        "step 3.\n\n"
        "--- Guide ---\n\n"
    )
    return steps + _guide_text()
