"""Пути инструментов: исходные файлы, песочница изображений, каталог вывода, шаблон.

Шаблон .dotx пересохраняется во временный .docx, макрошаблоны отклоняются.
"""

from __future__ import annotations

from contextlib import ExitStack
from pathlib import Path
import tempfile
import zipfile

from docx import Document

from ..workflow import discover_sources


def resolve_inputs(inputs: list[str], mode: str) -> list[Path]:
    """Развернуть переданные пути в отсортированный список исходных файлов."""
    if not inputs:
        raise ValueError(
            "inputs must not be empty; pass at least one file or directory path"
        )
    return discover_sources([Path(item).expanduser() for item in inputs], mode)


def resolve_image_roots(inputs: list[str]) -> list[Path]:
    """Derive the local-image sandbox from *inputs*: one root per item.

    A directory input allows filesystem images anywhere under it; a file
    input allows them only next to it, in its parent directory -- naming
    one file should not license wandering into sibling directories too.
    An input that resolves to neither (a typo, a path that does not exist)
    contributes no root, same as it contributes nothing to
    ``discover_sources``.

    Deliberately not a common ancestor of all inputs: with inputs drawn
    from unrelated trees, the common ancestor can collapse to the
    filesystem root, silently discarding the restriction this function
    exists to build.
    """
    roots: list[Path] = []
    for item in inputs:
        candidate = Path(item).expanduser().resolve()
        if candidate.is_dir():
            roots.append(candidate)
        elif candidate.is_file():
            roots.append(candidate.parent)
    return roots


def prepare_output_dir(output_dir: str | None) -> Path | None:
    """Подготовить каталог назначения; None означает «рядом с исходником»."""
    if output_dir is None:
        return None
    # resolve() — иначе относительный output_dir оставит ConvertedFile.output
    # относительным, вопреки его же Field(description="Absolute path ...")
    directory = Path(output_dir).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory


_DOCUMENT_MAIN_TYPE = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
_TEMPLATE_MAIN_TYPE = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
)
_MACRO_MAIN_TYPES = (
    b"application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    b"application/vnd.ms-word.document.macroEnabled.main+xml",
)


def _content_types(path: Path) -> bytes | None:
    """Прочитать [Content_Types].xml пакета или None, если это не OOXML-zip."""
    try:
        with zipfile.ZipFile(path) as archive:
            return archive.read("[Content_Types].xml")
    except (zipfile.BadZipFile, KeyError, OSError):
        return None


def _docx_from_dotx(source: Path, target: Path) -> None:
    """Пересохранить .dotx как .docx: у них отличается лишь тип главной части.

    python-docx открывает только пакеты с типом document.main, а шаблон Word
    помечен как template.main — при одинаковом содержимом. Замена строки
    типа в [Content_Types].xml делает из шаблона обычный документ.
    """
    with zipfile.ZipFile(source) as reader, zipfile.ZipFile(target, "w") as writer:
        for item in reader.infolist():
            data = reader.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_TEMPLATE_MAIN_TYPE, _DOCUMENT_MAIN_TYPE)
            writer.writestr(item, data)


def resolve_template(template: str | None, scratch: ExitStack) -> Path | None:
    """Проверить шаблон до начала батча, а не отказом на каждом файле.

    Шаблон обязан открываться python-docx — рендерер открывает его так же.
    Шаблон .dotx пересохраняется как .docx во временный каталог, который
    живёт, пока открыт *scratch* (то есть до конца батча). Шаблоны с
    макросами (.dotm/.docm) отклоняются: макросы в .docx недопустимы.
    """
    if template is None:
        return None
    path = Path(template).expanduser().resolve()
    if not path.is_file():
        raise ValueError(
            f"template must be an existing .docx or .dotx file, but {path} "
            f"{'is not a file' if path.exists() else 'does not exist'}"
        )
    content_types = _content_types(path) or b""
    if any(macro_type in content_types for macro_type in _MACRO_MAIN_TYPES):
        raise ValueError(
            f"template {path} is a macro-enabled Word file (.dotm/.docm), which "
            "is not supported; save it as a .docx (or .dotx) first"
        )
    usable = path
    if _TEMPLATE_MAIN_TYPE in content_types:
        directory = Path(
            scratch.enter_context(tempfile.TemporaryDirectory(prefix="mdtoword-template-"))
        )
        usable = directory / f"{path.stem}.docx"
        _docx_from_dotx(path, usable)
    try:
        Document(str(usable))
    except Exception as error:
        raise ValueError(
            f"template {path} is not a readable Word document "
            f"({type(error).__name__}: {error}); pass a .docx or .dotx file"
        ) from error
    return usable
