from pathlib import Path
import unittest

from mdtoword.workflow import discover_sources, resolve_output_paths


class ConversionWorkflowTests(unittest.TestCase):
    def test_discovery_recurses_deduplicates_and_filters_markdown_sources(self):
        with self.subTest("setup"):
            root = Path(self._tmpdir.name)
            nested = root / "nested"
            nested.mkdir()
            first = root / "first.md"
            second = nested / "second.MD"
            ignored = nested / "ignored.txt"
            for path in (first, second, ignored):
                path.write_text("# test", encoding="utf-8")

        discovered = discover_sources([root, first], "md_to_word")

        self.assertEqual(discovered, [first.resolve(), second.resolve()])

    def test_selected_output_directory_allocates_collision_suffixes(self):
        root = Path(self._tmpdir.name)
        left = root / "left" / "report.md"
        right = root / "right" / "report.md"
        output = root / "output"
        left.parent.mkdir(parents=True)
        right.parent.mkdir(parents=True)
        output.mkdir()

        paths = resolve_output_paths([left, right], output, ".docx")

        self.assertEqual(paths[left], output / "report.docx")
        self.assertEqual(paths[right], output / "report (2).docx")

    def test_automatic_output_paths_stay_beside_each_input(self):
        root = Path(self._tmpdir.name)
        source = root / "nested" / "notes.md"
        source.parent.mkdir(parents=True)

        paths = resolve_output_paths([source], None, ".docx")

        self.assertEqual(paths[source], source.with_suffix(".docx"))

    def setUp(self):
        import tempfile

        self._tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self._tmpdir.cleanup()


if __name__ == "__main__":
    unittest.main()
