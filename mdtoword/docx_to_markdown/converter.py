"""The public Word → Markdown converter: files in, Markdown and warnings out."""

from __future__ import annotations

from pathlib import Path

from docx import Document

from ..errors import ConversionError
from .document import _Conversion
from .media import _Media


class WordToMarkdownConverter:
    """Convert a Word document into GitHub-flavoured Markdown.

    ``extract_media``
        Write embedded pictures next to the Markdown file, into
        ``<output stem>_media/``, and link them. When false only the alt
        text is kept and an ``image_not_extracted`` warning is raised.
    ``front_matter``
        Emit the document's title, author, subject, keywords and language
        (``lang``) as YAML front matter when any of them is set.

    Contract, as for :class:`~mdtoword.converters.MarkdownToWordConverter`:
    success returns the list of warnings (``ConversionWarning`` instances
    with a ``code``), failure raises :class:`~mdtoword.errors.ConversionError`.
    """

    def __init__(self, *, extract_media: bool = True, front_matter: bool = True) -> None:
        self.extract_media = extract_media
        self.front_matter = front_matter

    def convert_file(self, input_path: str | Path, output_path: str | Path) -> list[str]:
        """Convert ``input_path`` and write the Markdown to ``output_path``."""
        output = Path(output_path)
        media_directory = output.parent / f"{output.stem}_media"
        markdown, warnings = self.convert_to_string(
            input_path,
            media_dir=media_directory if self.extract_media else None,
            media_link_prefix=media_directory.name,
        )
        try:
            output.write_text(markdown, encoding="utf-8")
        except Exception as error:
            raise ConversionError(str(error)) from error
        return warnings

    def convert_to_string(
        self,
        input_path: str | Path,
        *,
        media_dir: str | Path | None = None,
        media_link_prefix: str = "",
    ) -> tuple[str, list[str]]:
        """Convert a document and return ``(markdown, warnings)``.

        Images are written to ``media_dir`` and linked as
        ``<media_link_prefix>/imageN.ext``; with no ``media_dir`` they are
        not extracted.
        """
        try:
            document = Document(str(input_path))
            media = _Media(Path(media_dir) if media_dir is not None else None,
                           media_link_prefix.strip("/"))
            conversion = _Conversion(document, media, front_matter=self.front_matter)
            markdown = conversion.convert()
            return markdown, list(conversion.warnings())
        except ConversionError:
            raise
        except Exception as error:
            raise ConversionError(str(error)) from error
