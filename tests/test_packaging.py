"""Guards against ``environment.yml`` and ``requirements.txt`` drifting apart.

The conda environment installs its Python dependencies through a pip block
inside ``environment.yml``. That block is hand-maintained separately from
``requirements.txt`` (the source of truth used by ``pip install -r``), so
nothing stops the two lists from silently disagreeing -- which is exactly
what happened when ``markdown-it-py``, ``mdit-py-plugins`` and
``linkify-it-py`` were added to ``requirements.txt`` without touching
``environment.yml``: a conda-created environment could no longer import the
application. This test parses both files and asserts the pinned package
sets are identical.
"""

from pathlib import Path
import unittest

_REPO_ROOT = Path(__file__).resolve().parent.parent
_ENVIRONMENT_YML = _REPO_ROOT / "environment.yml"
_REQUIREMENTS_TXT = _REPO_ROOT / "requirements.txt"


def _parse_requirements(path: Path) -> dict[str, str]:
    """Map package name -> pinned version from a ``pip``-style requirements file."""
    packages: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        name, separator, version = line.partition("==")
        if not separator:
            raise ValueError(f"Unpinned or unsupported requirement line: {raw_line!r}")
        packages[name.strip()] = version.strip()
    return packages


def _parse_environment_pip_block(path: Path) -> dict[str, str]:
    """Map package name -> pinned version from the ``pip:`` sub-list of a conda env file.

    Written by hand rather than via PyYAML, since the test environment does
    not have PyYAML installed and ``environment.yml``'s structure is simple
    enough not to need it: find the ``- pip:`` line, then collect every
    more-indented ``- name==version`` line that follows it.
    """
    lines = path.read_text(encoding="utf-8").splitlines()
    packages: dict[str, str] = {}
    in_pip_block = False
    pip_indent = 0
    for raw_line in lines:
        stripped = raw_line.strip()
        if not in_pip_block:
            if stripped == "- pip:":
                in_pip_block = True
                pip_indent = len(raw_line) - len(raw_line.lstrip(" "))
            continue
        if not stripped:
            continue
        indent = len(raw_line) - len(raw_line.lstrip(" "))
        if indent <= pip_indent:
            break
        item = stripped[1:].strip() if stripped.startswith("-") else stripped
        name, separator, version = item.partition("==")
        if not separator:
            raise ValueError(f"Unpinned or unsupported pip entry: {raw_line!r}")
        packages[name.strip()] = version.strip()
    return packages


class PackagingTests(unittest.TestCase):
    def test_environment_yml_pip_block_matches_requirements_txt(self):
        requirements = _parse_requirements(_REQUIREMENTS_TXT)
        environment_pip = _parse_environment_pip_block(_ENVIRONMENT_YML)

        # Sanity check the parsers themselves found something, so a parsing
        # bug that returns an empty dict on both sides can't masquerade as
        # a passing test.
        self.assertTrue(requirements)

        self.assertEqual(
            environment_pip,
            requirements,
            "environment.yml's pip block must list the exact same packages "
            "and pinned versions as requirements.txt",
        )


if __name__ == "__main__":
    unittest.main()
