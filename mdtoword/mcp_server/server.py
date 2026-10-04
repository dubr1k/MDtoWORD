"""Экземпляр FastMCP, инструкции для агента и точка входа stdio.

Инструменты, ресурс и промпт регистрируются на ``mcp`` модулями
``tools_*`` и ``guide``; их импортирует ``__init__`` пакета.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

GUIDE_URI = "mdtoword://guide/markdown"

mcp = FastMCP(
    "mdtoword",
    instructions=(
        "Converts between Markdown and Word (.docx) by file path; document "
        "content never crosses the protocol. Recommended flow: read the "
        f"resource {GUIDE_URI} (or use the prepare_markdown_for_word prompt), "
        "run check_latex on the formulas, run preview_markdown and fix every "
        "warning at its reported line, convert with markdown_to_word, then "
        "verify the .docx with inspect_docx."
    ),
)


def main() -> None:
    """Запустить сервер на транспорте stdio."""
    mcp.run()
