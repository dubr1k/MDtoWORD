from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path


_SUPPORTED_SUFFIXES = {
    "md_to_word": frozenset({".md", ".markdown"}),
    "word_to_md": frozenset({".docx"}),
}


def supported_suffixes(mode: str) -> frozenset[str]:
    """Return the input filename extensions accepted by a conversion mode."""
    try:
        return _SUPPORTED_SUFFIXES[mode]
    except KeyError as error:
        raise ValueError(f"Unsupported conversion mode: {mode}") from error


def discover_sources(paths: Iterable[Path], mode: str) -> list[Path]:
    """Recursively collect unique, supported source files in canonical order."""
    suffixes = supported_suffixes(mode)
    sources: set[Path] = set()

    for path in paths:
        if path.is_file():
            candidates = (path,)
        elif path.is_dir():
            candidates = path.rglob("*")
        else:
            continue

        for candidate in candidates:
            if candidate.is_file() and candidate.suffix.lower() in suffixes:
                sources.add(candidate.resolve())

    return sorted(sources)


def resolve_output_paths(
    inputs: Sequence[Path], output_directory: Path | None, suffix: str
) -> dict[Path, Path]:
    """Allocate one output path for every input without batch name collisions."""
    allocated: set[Path] = set()
    outputs: dict[Path, Path] = {}

    for source in inputs:
        directory = output_directory if output_directory is not None else source.parent
        candidate = directory / f"{source.stem}{suffix}"
        index = 2

        while candidate in allocated:
            candidate = directory / f"{source.stem} ({index}){suffix}"
            index += 1

        allocated.add(candidate)
        outputs[source] = candidate

    return outputs
