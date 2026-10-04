"""Тесты MCP-сервера через in-memory клиент из SDK.

Клиент и сервер соединяются напрямую в одном процессе, без подпроцесса
и без stdio, — проверяется реальный путь вызова инструмента вместе со
схемами и валидацией аргументов.

Рендерер и конвертеры развиваются отдельно от сервера, поэтому всё, что
касается самого слоя MCP (передача опций, форма отчётов, прогресс, потоки),
проверяется на подменённом конвертере, а не по деталям готового документа.
"""

from io import BytesIO
from pathlib import Path
import os
import tempfile
import threading
import typing
import unittest
from unittest.mock import MagicMock, patch
import zipfile

from docx import Document
from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

from mdtoword import latex_omml, options
from mdtoword.errors import ConversionError, ConversionWarning
from mdtoword.options import DocumentOptions

try:
    from mcp.shared.memory import (
        create_connected_server_and_client_session as client_session,
    )
    from pydantic import AnyUrl
except ImportError:  # pragma: no cover
    client_session = None
    mcp_server = None
    server_factory = None
    server = None
else:
    # Пропуск — только когда нет самого SDK. Ошибка импорта сервера или
    # рендерера — это сломанный код, и прогон должен об этом сказать.
    from mdtoword import mcp_server
    from mdtoword.mcp_server import factory as server_factory

    server = mcp_server.mcp


# The smallest well-formed PNG, for images the renderer must embed.
_MINIMAL_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
    b"\x00\x00\x05\x00\x01\x05-\xb4\x00\x00\x00\x00\x00IEND\xaeB`\x82"
)

# Сервер создаёт конвертеры только в mcp_server.factory — подменять там.
_CONVERTER = "mdtoword.mcp_server.factory.MarkdownToWordConverter"

_DOCX_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_DOTX_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
_DOTM_MAIN = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"


def _save_word_package(path: Path, main_content_type: str) -> None:
    """Сохранить пустой документ с заданным типом главной части (.docx/.dotx/.dotm)."""
    buffer = BytesIO()
    Document().save(buffer)
    with zipfile.ZipFile(BytesIO(buffer.getvalue())) as reader, \
            zipfile.ZipFile(path, "w") as writer:
        for item in reader.infolist():
            data = reader.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_DOCX_MAIN.encode(), main_content_type.encode())
            writer.writestr(item, data)


def _messages(warnings: list[dict]) -> list[str]:
    return [warning["message"] for warning in warnings]


def setUpModule() -> None:
    """Пропустить набор целиком, если SDK mcp не установлен.

    Пропуск объявлен здесь, а не прямо на уровне модуля: ``python -m unittest``
    не перехватывает ``SkipTest``, выброшенный во время импорта, и обрывает
    весь прогон, а README документирует именно эту команду. ``setUpModule``
    же корректно понимают оба раннера.
    """
    if server is None:  # pragma: no cover
        raise unittest.SkipTest("SDK mcp не установлен; см. requirements-mcp.txt")


class McpServerTestCase(unittest.IsolatedAsyncioTestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self.root = Path(self._tmpdir.name)

    async def call(self, tool: str, arguments: dict, **kwargs):
        async with client_session(server._mcp_server) as client:
            return await client.call_tool(tool, arguments, **kwargs)

    async def list_tools(self) -> dict:
        async with client_session(server._mcp_server) as client:
            listed = await client.list_tools()
        return {tool.name: tool for tool in listed.tools}

    def write_markdown(self, *names: str) -> None:
        for name in names:
            (self.root / name).write_text("# Заголовок", encoding="utf-8")


class ToolRegistrationTests(McpServerTestCase):
    async def test_all_tools_are_advertised_with_descriptions_and_titles(self) -> None:
        tools = await self.list_tools()

        self.assertEqual(
            set(tools),
            {
                "markdown_to_word",
                "word_to_markdown",
                "preview_markdown",
                "check_latex",
                "inspect_docx",
            },
        )
        for tool in tools.values():
            self.assertTrue(tool.description)
            self.assertTrue(tool.title)

    async def test_word_to_markdown_describes_its_losses_and_media_folder(self) -> None:
        tools = await self.list_tools()
        tool = tools["word_to_markdown"]
        description = tool.description.lower()

        self.assertIn("merged table cells", description)
        self.assertIn("_media/", description)
        self.assertEqual(tool.inputSchema["properties"]["extract_media"]["default"], True)

    async def test_word_to_markdown_passes_extract_media_through(self) -> None:
        with patch.object(server_factory, "WordToMarkdownConverter") as converter_class:
            converter_class.return_value.convert_file.return_value = []
            with tempfile.TemporaryDirectory() as directory:
                source = Path(directory) / "a.docx"
                Document().save(str(source))
                await self.call(
                    "word_to_markdown", {"inputs": [str(source)], "extract_media": False}
                )

        converter_class.assert_called_once_with(extract_media=False)

    async def test_annotations_describe_side_effects(self) -> None:
        tools = await self.list_tools()

        expected = {
            "markdown_to_word": (False, True, True, True),
            "word_to_markdown": (False, True, True, False),
            "preview_markdown": (True, None, True, True),
            "check_latex": (True, None, True, False),
            "inspect_docx": (True, None, True, False),
        }
        for name, (read_only, destructive, idempotent, open_world) in expected.items():
            annotations = tools[name].annotations
            with self.subTest(tool=name):
                self.assertIsNotNone(annotations)
                self.assertIs(annotations.readOnlyHint, read_only)
                self.assertIs(annotations.destructiveHint, destructive)
                self.assertIs(annotations.idempotentHint, idempotent)
                self.assertIs(annotations.openWorldHint, open_world)
                self.assertEqual(annotations.title, tools[name].title)


class SchemaTests(McpServerTestCase):
    async def test_context_parameter_is_hidden_from_every_input_schema(self) -> None:
        tools = await self.list_tools()

        for name, tool in tools.items():
            with self.subTest(tool=name):
                self.assertNotIn("ctx", tool.inputSchema["properties"])

    async def test_document_options_are_enums_in_the_schema(self) -> None:
        tools = await self.list_tools()

        for name in ("markdown_to_word", "preview_markdown"):
            properties = tools[name].inputSchema["properties"]
            with self.subTest(tool=name):
                self.assertEqual(properties["preset"]["enum"], ["default", "gost"])
                self.assertEqual(properties["line_breaks"]["enum"], ["soft", "preserve"])
                self.assertEqual(properties["footnotes"]["enum"], ["native", "section"])
                page_size_enums = [
                    option["enum"]
                    for option in properties["page_size"]["anyOf"]
                    if "enum" in option
                ]
                self.assertEqual(page_size_enums, [["A4", "Letter"]])
                self.assertEqual(properties["toc"]["type"], "boolean")
                self.assertIn("template", properties)
                self.assertIn("language", properties)

    async def test_font_size_is_optional_and_bounded(self) -> None:
        tools = await self.list_tools()

        font_size = tools["markdown_to_word"].inputSchema["properties"]["font_size"]
        self.assertIsNone(font_size["default"])
        number = next(o for o in font_size["anyOf"] if o.get("type") == "number")
        self.assertEqual(number["exclusiveMinimum"], 0)
        self.assertEqual(number["maximum"], 400)
        font_name = tools["markdown_to_word"].inputSchema["properties"]["font_name"]
        self.assertEqual(font_name["minLength"], 1)

    async def test_out_of_range_font_arguments_are_rejected(self) -> None:
        self.write_markdown("doc.md")

        for arguments in ({"font_size": 0}, {"font_size": 401}, {"font_name": ""}, {"font_name": "   "}):
            with self.subTest(arguments=arguments):
                result = await self.call(
                    "markdown_to_word", {"inputs": [str(self.root)], **arguments}
                )
                self.assertTrue(result.isError)
        self.assertFalse((self.root / "doc.docx").exists())

    async def test_warnings_are_structured_in_the_output_schema(self) -> None:
        tools = await self.list_tools()

        for name in ("markdown_to_word", "preview_markdown"):
            definitions = tools[name].outputSchema["$defs"]
            with self.subTest(tool=name):
                self.assertEqual(
                    set(definitions["Diagnostic"]["properties"]), {"message", "code", "line"}
                )

    def test_literal_choices_match_the_options_module(self) -> None:
        self.assertEqual(typing.get_args(mcp_server.Preset), options.PRESETS)
        self.assertEqual(typing.get_args(mcp_server.PageSize), options.PAGE_SIZES)
        self.assertEqual(typing.get_args(mcp_server.LineBreaks), options.LINE_BREAK_MODES)
        self.assertEqual(typing.get_args(mcp_server.FootnoteMode), options.FOOTNOTE_MODES)


class MarkdownToWordTests(McpServerTestCase):
    async def test_a_directory_is_converted_recursively_with_one_output_each(self) -> None:
        nested = self.root / "nested"
        nested.mkdir()
        (self.root / "first.md").write_text("# Первый", encoding="utf-8")
        (nested / "second.markdown").write_text("# Второй", encoding="utf-8")
        (self.root / "ignored.txt").write_text("не markdown", encoding="utf-8")

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        self.assertFalse(result.isError)
        report = result.structuredContent
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["converted"]), 2)
        self.assertEqual(report["failed"], [])
        for entry in report["converted"]:
            self.assertTrue(Path(entry["output"]).is_file())

    async def test_output_dir_is_created_and_used(self) -> None:
        (self.root / "doc.md").write_text("# Заголовок", encoding="utf-8")
        destination = self.root / "out" / "deep"

        result = await self.call(
            "markdown_to_word",
            {"inputs": [str(self.root / "doc.md")], "output_dir": str(destination)},
        )

        report = result.structuredContent
        # .resolve() с обеих сторон: на macOS временный каталог лежит под /var,
        # который является симлинком на /private/var, а _prepare_output_dir
        # теперь резолвит output_dir, поэтому путь в отчёте — каноническая
        # форма destination.
        self.assertEqual(
            Path(report["converted"][0]["output"]).parent, destination.resolve()
        )
        self.assertTrue((destination / "doc.docx").is_file())

    async def test_relative_output_dir_is_resolved_to_an_absolute_path(self) -> None:
        self.addCleanup(os.chdir, os.getcwd())
        (self.root / "doc.md").write_text("# Заголовок", encoding="utf-8")
        os.chdir(self.root)

        result = await self.call(
            "markdown_to_word",
            {"inputs": ["doc.md"], "output_dir": "out"},
        )

        report = result.structuredContent
        output = Path(report["converted"][0]["output"])
        self.assertTrue(output.is_absolute())
        self.assertTrue(output.is_file())

    async def test_nonfatal_warnings_are_reported_per_file(self) -> None:
        (self.root / "doc.md").write_text("![diagram](missing.png)", encoding="utf-8")

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        warnings = result.structuredContent["converted"][0]["warnings"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("missing.png", warnings[0]["message"])

    async def test_font_arguments_reach_the_document(self) -> None:
        (self.root / "doc.md").write_text("Текст", encoding="utf-8")

        result = await self.call(
            "markdown_to_word",
            {"inputs": [str(self.root)], "font_name": "Georgia", "font_size": 14},
        )

        output = Path(result.structuredContent["converted"][0]["output"])
        document = Document(str(output))
        self.assertEqual(document.styles["Normal"].font.name, "Georgia")

    async def test_paths_matching_nothing_report_zero_sources_found(self) -> None:
        result = await self.call(
            "markdown_to_word", {"inputs": [str(self.root / "нет-такой-папки")]}
        )

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 0)
        self.assertEqual(report["converted"], [])
        self.assertEqual(report["failed"], [])

    async def test_output_dir_is_not_created_when_nothing_matches(self) -> None:
        destination = self.root / "out" / "deep"

        result = await self.call(
            "markdown_to_word",
            {
                "inputs": [str(self.root / "нет-такой-папки")],
                "output_dir": str(destination),
            },
        )

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 0)
        self.assertFalse(destination.exists())

    async def test_empty_inputs_is_an_error_not_an_empty_success(self) -> None:
        result = await self.call("markdown_to_word", {"inputs": []})

        self.assertTrue(result.isError)

    async def test_default_does_not_fetch_remote_images(self) -> None:
        (self.root / "doc.md").write_text(
            "![diagram](https://example.invalid/x.png)", encoding="utf-8"
        )

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        warnings = result.structuredContent["converted"][0]["warnings"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("fetch_remote_images", warnings[0]["message"])

    async def test_image_outside_inputs_tree_is_refused_and_names_image_root(self) -> None:
        with tempfile.TemporaryDirectory() as outside_dir:
            outside_image = Path(outside_dir) / "secret.png"
            outside_image.write_bytes(_MINIMAL_PNG)
            (self.root / "doc.md").write_text(
                f"![diagram]({outside_image})", encoding="utf-8"
            )

            result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        warnings = _messages(result.structuredContent["converted"][0]["warnings"])
        self.assertEqual(len(warnings), 1)
        self.assertIn("outside the allowed root", warnings[0])
        self.assertIn("image_root", warnings[0])

    async def test_shared_assets_layout_with_root_as_input_embeds(self) -> None:
        # The common ../images/logo.png layout: the caller passed the
        # containing folder, so a sibling directory one level up from the
        # document is still inside the allowed root.
        (self.root / "guide").mkdir()
        (self.root / "images").mkdir()
        (self.root / "images" / "logo.png").write_bytes(_MINIMAL_PNG)
        (self.root / "guide" / "x.md").write_text(
            "![logo](../images/logo.png)", encoding="utf-8"
        )

        result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 1)
        self.assertEqual(report["converted"][0]["warnings"], [])

    async def test_image_under_second_of_multiple_inputs_is_allowed(self) -> None:
        # Regression guard: roots must stay a list, one per input. Collapsing
        # them into a single common ancestor (or keeping only the first) is
        # the wrong "simplification" this test exists to catch.
        with tempfile.TemporaryDirectory() as first_dir, \
                tempfile.TemporaryDirectory() as second_dir:
            first = Path(first_dir)
            second = Path(second_dir)
            (second / "diagram.png").write_bytes(_MINIMAL_PNG)
            (first / "doc.md").write_text(
                f"![diagram]({second / 'diagram.png'})", encoding="utf-8"
            )

            result = await self.call(
                "markdown_to_word", {"inputs": [str(first), str(second)]}
            )

        report = result.structuredContent
        self.assertEqual(report["converted"][0]["warnings"], [])

    async def test_image_root_override_widens_access_to_a_refused_path(self) -> None:
        with tempfile.TemporaryDirectory() as outside_dir:
            outside_image = Path(outside_dir) / "secret.png"
            outside_image.write_bytes(_MINIMAL_PNG)
            (self.root / "doc.md").write_text(
                f"![diagram]({outside_image})", encoding="utf-8"
            )

            result = await self.call(
                "markdown_to_word",
                {"inputs": [str(self.root)], "image_root": str(outside_dir)},
            )

        report = result.structuredContent
        self.assertEqual(report["converted"][0]["warnings"], [])


class ConverterArgumentsTests(McpServerTestCase):
    """Аргументы инструмента доходят до конвертера: опции — как DocumentOptions."""

    async def converter_call(self, tool: str, arguments: dict) -> tuple[MagicMock, object]:
        self.write_markdown("doc.md")
        with patch(_CONVERTER) as converter_class:
            converter_class.return_value.convert_file.return_value = []
            converter_class.return_value.preview_file.return_value = []
            result = await self.call(tool, {"inputs": [str(self.root)], **arguments})
        self.assertFalse(result.isError, result.content)
        converter_class.assert_called_once()
        return converter_class, result

    async def test_defaults_build_default_options_and_12pt(self) -> None:
        converter_class, _ = await self.converter_call("markdown_to_word", {})

        kwargs = converter_class.call_args.kwargs
        self.assertEqual(kwargs["document_options"], DocumentOptions())
        self.assertEqual(kwargs["font_size"], Pt(12))
        self.assertEqual(kwargs["font_name"], "Times New Roman")
        self.assertFalse(kwargs["allow_remote_images"])

    async def test_every_option_is_passed_through(self) -> None:
        template = self.root / "reference.docx"
        Document().save(str(template))

        for tool in ("markdown_to_word", "preview_markdown"):
            with self.subTest(tool=tool):
                converter_class, _ = await self.converter_call(
                    tool,
                    {
                        "preset": "gost",
                        "page_size": "Letter",
                        "language": "ru-RU",
                        "line_breaks": "preserve",
                        "template": str(template),
                        "toc": True,
                        "footnotes": "section",
                    },
                )

                kwargs = converter_class.call_args.kwargs
                self.assertEqual(
                    kwargs["document_options"],
                    DocumentOptions(
                        preset="gost",
                        page_size="Letter",
                        language="ru-RU",
                        line_breaks="preserve",
                        template=template.resolve(),
                        toc=True,
                        footnotes="section",
                    ),
                )

    async def test_remote_fetching_is_off_unless_requested(self) -> None:
        for tool in ("markdown_to_word", "preview_markdown"):
            with self.subTest(tool=tool):
                default_class, _ = await self.converter_call(tool, {})
                enabled_class, _ = await self.converter_call(
                    tool, {"fetch_remote_images": True}
                )

                self.assertIs(default_class.call_args.kwargs["allow_remote_images"], False)
                self.assertIs(enabled_class.call_args.kwargs["allow_remote_images"], True)

    async def test_image_roots_come_from_inputs_or_image_root(self) -> None:
        derived_class, _ = await self.converter_call("markdown_to_word", {})
        override_class, _ = await self.converter_call(
            "markdown_to_word", {"image_root": str(self.root / "assets")}
        )

        self.assertEqual(derived_class.call_args.kwargs["image_roots"], [self.root.resolve()])
        self.assertEqual(
            override_class.call_args.kwargs["image_roots"], [(self.root / "assets").resolve()]
        )

    async def test_gost_preset_defaults_the_font_size_to_14(self) -> None:
        converter_class, _ = await self.converter_call(
            "markdown_to_word", {"preset": "gost"}
        )

        self.assertEqual(converter_class.call_args.kwargs["font_size"], Pt(14))

    async def test_explicit_font_size_wins_over_the_preset(self) -> None:
        converter_class, _ = await self.converter_call(
            "preview_markdown", {"preset": "gost", "font_size": 13}
        )

        self.assertEqual(converter_class.call_args.kwargs["font_size"], Pt(13))

    async def test_missing_template_is_a_clear_error_before_anything_is_written(self) -> None:
        self.write_markdown("doc.md")
        destination = self.root / "out"

        for tool in ("markdown_to_word", "preview_markdown"):
            arguments = {"inputs": [str(self.root)], "template": str(self.root / "нет.docx")}
            if tool == "markdown_to_word":
                arguments["output_dir"] = str(destination)
            with self.subTest(tool=tool):
                result = await self.call(tool, arguments)

                self.assertTrue(result.isError)
                self.assertIn("template", result.content[0].text)
                self.assertIn("нет.docx", result.content[0].text)
        self.assertFalse(destination.exists())

    async def test_dotx_template_is_usable_for_the_batch_then_cleaned_up(self) -> None:
        dotx = self.root / "reference.dotx"
        _save_word_package(dotx, _DOTX_MAIN)
        with self.assertRaises(ValueError):
            Document(str(dotx))  # python-docx itself refuses a .dotx
        seen_templates: list[Path] = []

        def convert(source, output):
            template = converter_class.call_args.kwargs["document_options"].template
            Document(str(template))  # opens while the batch runs
            seen_templates.append(template)
            return []

        self.write_markdown("doc.md")
        with patch(_CONVERTER) as converter_class:
            converter_class.return_value.convert_file.side_effect = convert
            result = await self.call(
                "markdown_to_word", {"inputs": [str(self.root)], "template": str(dotx)}
            )

        self.assertFalse(result.isError, result.content)
        self.assertEqual(len(seen_templates), 1)
        self.assertEqual(seen_templates[0].suffix, ".docx")
        self.assertFalse(seen_templates[0].exists())

    async def test_dotx_template_converts_end_to_end(self) -> None:
        dotx = self.root / "reference.dotx"
        _save_word_package(dotx, _DOTX_MAIN)
        self.write_markdown("doc.md")

        for tool in ("markdown_to_word", "preview_markdown"):
            with self.subTest(tool=tool):
                result = await self.call(
                    tool, {"inputs": [str(self.root / "doc.md")], "template": str(dotx)}
                )

                self.assertFalse(result.isError, result.content)
                self.assertEqual(result.structuredContent["failed"], [])

    async def test_unusable_templates_are_rejected_before_anything_is_written(self) -> None:
        dotm = self.root / "macros.dotm"
        _save_word_package(dotm, _DOTM_MAIN)
        garbage = self.root / "garbage.docx"
        garbage.write_text("не документ", encoding="utf-8")
        self.write_markdown("doc.md")
        destination = self.root / "out"

        for template, expected in ((dotm, "macro-enabled"), (garbage, "not a readable Word document")):
            with self.subTest(template=template.name):
                result = await self.call(
                    "markdown_to_word",
                    {
                        "inputs": [str(self.root / "doc.md")],
                        "template": str(template),
                        "output_dir": str(destination),
                    },
                )

                self.assertTrue(result.isError)
                self.assertIn(expected, result.content[0].text)
        self.assertFalse(destination.exists())

    async def test_invalid_language_tag_is_rejected(self) -> None:
        self.write_markdown("doc.md")

        result = await self.call(
            "markdown_to_word", {"inputs": [str(self.root)], "language": "not a tag"}
        )

        self.assertTrue(result.isError)


class BatchExecutionTests(McpServerTestCase):
    async def test_progress_is_reported_after_each_file(self) -> None:
        self.write_markdown("a.md", "b.md")
        events: list[tuple[float, float | None, str | None]] = []

        async def on_progress(progress, total, message):
            events.append((progress, total, message))

        for tool, method in (("markdown_to_word", "convert_file"), ("preview_markdown", "preview_file")):
            events.clear()
            with self.subTest(tool=tool), patch(_CONVERTER) as converter_class:
                getattr(converter_class.return_value, method).return_value = []
                result = await self.call(
                    tool, {"inputs": [str(self.root)]}, progress_callback=on_progress
                )

                self.assertFalse(result.isError)
                self.assertEqual([(p, t) for p, t, _ in events], [(1, 2), (2, 2)])
                self.assertIn("a.md", events[0][2])
                self.assertIn("b.md", events[1][2])

    async def test_failed_progress_notification_does_not_lose_the_report(self) -> None:
        self.write_markdown("a.md", "b.md")

        async def broken_report_progress(self, *args, **kwargs):
            raise RuntimeError("client went away")

        async def on_progress(progress, total, message):
            pass

        with patch(_CONVERTER) as converter_class, patch.object(
            mcp_server.Context, "report_progress", broken_report_progress
        ), self.assertLogs("mdtoword.mcp_server", level="WARNING") as logs:
            converter_class.return_value.convert_file.return_value = []
            result = await self.call(
                "markdown_to_word", {"inputs": [str(self.root)]}, progress_callback=on_progress
            )

        self.assertFalse(result.isError, result.content)
        self.assertEqual(len(result.structuredContent["converted"]), 2)
        self.assertEqual(len(logs.records), 2)

    async def test_conversion_runs_off_the_event_loop_thread(self) -> None:
        self.write_markdown("a.md")
        threads: list[int] = []

        def convert(source, output):
            threads.append(threading.get_ident())
            return []

        with patch(_CONVERTER) as converter_class:
            converter_class.return_value.convert_file.side_effect = convert
            await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        self.assertEqual(len(threads), 1)
        self.assertNotEqual(threads[0], threading.get_ident())

    async def test_one_failure_does_not_stop_the_batch(self) -> None:
        self.write_markdown("a.md", "b.md", "c.md")

        def convert(source, output):
            if source.name == "a.md":
                raise ConversionError("broken")
            if source.name == "b.md":
                raise RuntimeError("bug")
            return []

        with patch(_CONVERTER) as converter_class:
            converter_class.return_value.convert_file.side_effect = convert
            result = await self.call("markdown_to_word", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertFalse(result.isError)
        self.assertEqual(report["sources_found"], 3)
        self.assertEqual([Path(e["source"]).name for e in report["converted"]], ["c.md"])
        errors = {Path(e["source"]).name: e["error"] for e in report["failed"]}
        self.assertEqual(errors["a.md"], "broken")
        self.assertIn("RuntimeError", errors["b.md"])


class DiagnosticTests(McpServerTestCase):
    async def test_warning_code_and_line_are_passed_through(self) -> None:
        self.write_markdown("doc.md")
        warnings = [
            ConversionWarning("Formula kept as text", code="formula_unsupported", line=7),
            "A plain string warning",
            ConversionWarning("No line known", code="image_not_found"),
        ]
        expected = [
            {"message": "Formula kept as text", "code": "formula_unsupported", "line": 7},
            {"message": "A plain string warning", "code": "general", "line": None},
            {"message": "No line known", "code": "image_not_found", "line": None},
        ]

        with patch(_CONVERTER) as converter_class:
            converter_class.return_value.convert_file.return_value = warnings
            converter_class.return_value.preview_file.return_value = warnings
            converted = await self.call("markdown_to_word", {"inputs": [str(self.root)]})
            previewed = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        self.assertEqual(converted.structuredContent["converted"][0]["warnings"], expected)
        self.assertEqual(previewed.structuredContent["previews"][0]["warnings"], expected)


class WordToMarkdownTests(McpServerTestCase):
    def write_docx(self, name: str) -> Path:
        path = self.root / name
        document = Document()
        document.add_heading("Раздел", level=1)
        document.save(str(path))
        return path

    async def test_documents_are_converted_to_markdown_files(self) -> None:
        self.write_docx("report.docx")

        result = await self.call("word_to_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 1)
        output = Path(report["converted"][0]["output"])
        self.assertEqual(output.suffix, ".md")
        self.assertIn("# Раздел", output.read_text(encoding="utf-8"))

    async def test_a_broken_file_fails_alone_without_stopping_the_batch(self) -> None:
        self.write_docx("good.docx")
        (self.root / "broken.docx").write_text("это не zip-контейнер", encoding="utf-8")

        result = await self.call("word_to_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertFalse(result.isError)
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["converted"]), 1)
        self.assertEqual(len(report["failed"]), 1)
        self.assertTrue(report["failed"][0]["source"].endswith("broken.docx"))
        self.assertTrue(report["failed"][0]["error"])

    async def test_output_dir_is_not_created_when_nothing_matches(self) -> None:
        destination = self.root / "out" / "deep"

        result = await self.call(
            "word_to_markdown",
            {
                "inputs": [str(self.root / "нет-такой-папки")],
                "output_dir": str(destination),
            },
        )

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 0)
        self.assertFalse(destination.exists())


class PreviewTests(McpServerTestCase):
    async def test_preview_reports_warnings_and_writes_no_files(self) -> None:
        (self.root / "doc.md").write_text("![diagram](missing.png)", encoding="utf-8")
        before = sorted(path.name for path in self.root.iterdir())

        result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 1)
        warnings = report["previews"][0]["warnings"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("missing.png", warnings[0]["message"])
        self.assertEqual(sorted(path.name for path in self.root.iterdir()), before)

    async def test_preview_reports_unreadable_files_in_failed(self) -> None:
        (self.root / "good.md").write_text("# Заголовок", encoding="utf-8")
        broken = self.root / "broken.md"
        broken.write_bytes(b"\xff\xfe\x00 invalid utf-8")

        result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 2)
        self.assertEqual(len(report["previews"]), 1)
        self.assertEqual(len(report["failed"]), 1)
        self.assertTrue(report["failed"][0]["source"].endswith("broken.md"))

    async def test_paths_matching_nothing_report_zero_sources_found(self) -> None:
        result = await self.call(
            "preview_markdown", {"inputs": [str(self.root / "нет-такой-папки")]}
        )

        report = result.structuredContent
        self.assertEqual(report["sources_found"], 0)
        self.assertEqual(report["previews"], [])
        self.assertEqual(report["failed"], [])

    async def test_default_does_not_fetch_remote_images(self) -> None:
        (self.root / "doc.md").write_text(
            "![diagram](https://example.invalid/x.png)", encoding="utf-8"
        )

        result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        warnings = result.structuredContent["previews"][0]["warnings"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("fetch_remote_images", warnings[0]["message"])

    async def test_image_outside_inputs_tree_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as outside_dir:
            outside_image = Path(outside_dir) / "secret.png"
            outside_image.write_bytes(_MINIMAL_PNG)
            (self.root / "doc.md").write_text(
                f"![diagram]({outside_image})", encoding="utf-8"
            )

            result = await self.call("preview_markdown", {"inputs": [str(self.root)]})

        warnings = _messages(result.structuredContent["previews"][0]["warnings"])
        self.assertEqual(len(warnings), 1)
        self.assertIn("outside the allowed root", warnings[0])


class CheckLatexTests(McpServerTestCase):
    async def check(self, formulas: list[str]) -> dict:
        result = await self.call("check_latex", {"formulas": formulas})
        self.assertFalse(result.isError, result.content)
        return result.structuredContent

    async def test_supported_and_unsupported_formulas(self) -> None:
        report = await self.check([r"\frac{a}{b}", r"\notacommand{x}"])

        good, bad = report["results"]
        self.assertEqual(good, {"formula": r"\frac{a}{b}", "ok": True, "error": None})
        self.assertEqual(bad["formula"], r"\notacommand{x}")
        self.assertFalse(bad["ok"])
        self.assertIn("notacommand", bad["error"])
        self.assertFalse(report["all_ok"])

    async def test_all_ok_when_every_formula_converts(self) -> None:
        report = await self.check([r"x^2", r"\sqrt{y}"])

        self.assertTrue(report["all_ok"])

    async def test_display_delimiters_are_stripped_before_parsing(self) -> None:
        seen: list[str] = []

        def fake_latex_to_omml(latex):
            seen.append(latex)
            return MagicMock()

        formulas = ["$$ b $$", r"\[d\]", "  e  "]
        with patch.object(latex_omml, "latex_to_omml", side_effect=fake_latex_to_omml):
            report = await self.check(formulas)

        self.assertEqual(seen, ["b", "d", "e"])
        self.assertEqual([r["formula"] for r in report["results"]], formulas)
        self.assertTrue(report["all_ok"])

    async def test_inline_delimiters_are_rendered_as_inline_math(self) -> None:
        rendered: list[str] = []
        document = MagicMock()
        document.element.body.xpath.return_value = [object()]

        def render(markdown, source_path):
            rendered.append(markdown)
            return document, []

        with patch(_CONVERTER) as converter_class:
            converter_class.return_value._render.side_effect = render
            report = await self.check(["$a$", r"\(c\)"])

        self.assertEqual(rendered, ["$a$\n", "$c$\n"])
        self.assertTrue(report["all_ok"])

    async def test_inline_math_is_judged_like_the_renderer_judges_it(self) -> None:
        # latex_to_omml would accept the first two, but inside a paragraph the
        # renderer keeps them as text (they read as prose, not formulas).
        report = await self.check(
            ["$x = длина$", "$PATH and HOME$", r"$\foo$", "$x^2$", r"\(\frac{a}{b}\)"]
        )

        verdicts = [(r["ok"], r["error"]) for r in report["results"]]
        self.assertEqual([ok for ok, _ in verdicts], [False, False, False, True, True])
        self.assertIn("prose", verdicts[0][1])
        self.assertIn("prose", verdicts[1][1])
        self.assertIn("foo", verdicts[2][1])
        self.assertIsNone(verdicts[3][1])
        self.assertFalse(report["all_ok"])

    async def test_equation_tag_is_split_off_when_latex_omml_supports_it(self) -> None:
        seen: list[str] = []

        def fake_latex_to_omml(latex):
            seen.append(latex)
            return MagicMock()

        with patch.object(
            latex_omml,
            "split_equation_tag",
            create=True,
            side_effect=lambda latex: ("x = 1", "1"),
        ), patch.object(latex_omml, "latex_to_omml", side_effect=fake_latex_to_omml):
            report = await self.check([r"x = 1 \tag{1}"])

        self.assertEqual(seen, ["x = 1"])
        self.assertTrue(report["all_ok"])

    async def test_empty_formula_is_not_ok(self) -> None:
        report = await self.check(["$$  $$"])

        self.assertFalse(report["results"][0]["ok"])
        self.assertTrue(report["results"][0]["error"])

    async def test_amsmath_environment_is_checked_like_the_renderer_treats_it(self) -> None:
        # A top-level amsmath environment must be judged the way
        # markdown_to_word treats it, not rejected as unknown input.
        report = await self.check([r"\begin{align} a &= b \\ c &= d \end{align}"])

        self.assertTrue(report["all_ok"], report)

    async def test_empty_list_is_rejected(self) -> None:
        result = await self.call("check_latex", {"formulas": []})

        self.assertTrue(result.isError)


_FOOTNOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
    "<w:continuationSeparator/></w:r></w:p></w:footnote>"
    '<w:footnote w:id="1"><w:p><w:r><w:t>Сноска</w:t></w:r></w:p></w:footnote>'
    "</w:footnotes>"
).encode("utf-8")


def _build_sample_docx(path: Path) -> None:
    """Документ со всем, что считает inspect_docx, по одной штуке."""
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin, section.right_margin = Mm(30), Mm(15)
    section.top_margin, section.bottom_margin = Mm(20), Mm(20)

    document.core_properties.title = "Отчёт о работе"
    document.core_properties.author = "Лаборатория"
    document.core_properties.keywords = "отчёт, тест"
    lang = document.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr/w:lang")[0]
    lang.set(qn("w:val"), "ru-RU")

    document.add_heading("Отчёт", level=0)
    toc = document.add_paragraph()
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), 'TOC \\o "1-3" \\h')
    toc._p.append(field)
    document.add_heading("Введение", level=1)
    document.add_paragraph("Обычный абзац.")
    document.add_heading("Метод", level=2)
    document.add_paragraph()._p.append(latex_omml.latex_to_omml("x^2"))
    document.add_picture(BytesIO(_MINIMAL_PNG))
    document.add_paragraph("Пункт списка", style="List Bullet")

    linked = document.add_paragraph()
    relationship_id = document.part.relate_to(
        "https://example.com", RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "ссылка"
    run.append(text)
    hyperlink.append(run)
    linked._p.append(hyperlink)

    table = document.add_table(rows=2, cols=2)
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            cell.text = f"{row_index}{column_index}"

    footnotes = Part(
        PackURI("/word/footnotes.xml"),
        CONTENT_TYPE.WML_FOOTNOTES,
        _FOOTNOTES_XML,
        document.part.package,
    )
    document.part.relate_to(footnotes, RELATIONSHIP_TYPE.FOOTNOTES)
    document.save(str(path))


class InspectDocxTests(McpServerTestCase):
    async def inspect(self, path: Path) -> dict:
        result = await self.call("inspect_docx", {"path": str(path)})
        self.assertFalse(result.isError, result.content)
        return result.structuredContent

    async def test_structure_of_a_known_document(self) -> None:
        path = self.root / "sample.docx"
        _build_sample_docx(path)

        report = await self.inspect(path)

        self.assertEqual(report["path"], str(path.resolve()))
        page = report["page"]
        self.assertEqual(page["paper"], "A4")
        self.assertEqual(page["orientation"], "portrait")
        self.assertAlmostEqual(page["width_mm"], 210.0, places=0)
        self.assertAlmostEqual(page["height_mm"], 297.0, places=0)
        self.assertAlmostEqual(page["margin_left_mm"], 30.0, places=0)
        self.assertAlmostEqual(page["margin_right_mm"], 15.0, places=0)
        self.assertAlmostEqual(page["margin_top_mm"], 20.0, places=0)
        self.assertAlmostEqual(page["margin_bottom_mm"], 20.0, places=0)
        self.assertEqual(report["section_count"], 1)
        self.assertEqual(
            report["counts"],
            {
                # Title, H1, абзац, H2, формула, рисунок, пункт, ссылка, 4 ячейки.
                "paragraphs": 12,
                "headings": 3,
                "tables": 1,
                "images": 1,
                "equations": 1,
                "footnotes": 1,
                "hyperlinks": 1,
                "list_paragraphs": 1,
            },
        )
        self.assertTrue(report["has_toc"])
        self.assertEqual(
            report["outline"],
            [
                {"level": 0, "text": "Отчёт"},
                {"level": 1, "text": "Введение"},
                {"level": 2, "text": "Метод"},
            ],
        )
        self.assertEqual(report["properties"]["title"], "Отчёт о работе")
        self.assertEqual(report["properties"]["author"], "Лаборатория")
        self.assertEqual(report["properties"]["keywords"], "отчёт, тест")
        self.assertIsNone(report["properties"]["subject"])
        self.assertEqual(report["language"], "ru-RU")
        self.assertTrue(report["text_preview"].startswith("Отчёт\nВведение\nОбычный абзац."))

    async def test_text_preview_is_truncated(self) -> None:
        path = self.root / "long.docx"
        document = Document()
        for _ in range(20):
            document.add_paragraph("слово " * 20)
        document.save(str(path))

        report = await self.inspect(path)

        self.assertLessEqual(len(report["text_preview"]), 501)
        self.assertTrue(report["text_preview"].endswith("…"))
        self.assertFalse(report["has_toc"])
        self.assertEqual(report["counts"]["footnotes"], 0)

    async def test_missing_file_is_a_clear_error(self) -> None:
        result = await self.call("inspect_docx", {"path": str(self.root / "нет.docx")})

        self.assertTrue(result.isError)
        self.assertIn("not found", result.content[0].text)

    async def test_non_docx_file_is_a_clear_error(self) -> None:
        path = self.root / "fake.docx"
        path.write_text("это не docx", encoding="utf-8")

        result = await self.call("inspect_docx", {"path": str(path)})

        self.assertTrue(result.isError)
        self.assertIn("Not a readable .docx", result.content[0].text)


class GuideResourceAndPromptTests(McpServerTestCase):
    async def test_guide_resource_is_listed_and_readable(self) -> None:
        async with client_session(server._mcp_server) as client:
            listed = await client.list_resources()
            contents = await client.read_resource(AnyUrl(mcp_server.GUIDE_URI))

        uris = [str(resource.uri) for resource in listed.resources]
        self.assertIn(mcp_server.GUIDE_URI, uris)
        self.assertEqual(len(contents.contents), 1)
        guide = contents.contents[0]
        self.assertEqual(guide.mimeType, "text/markdown")
        self.assertGreater(len(guide.text), 1000)
        for feature in ("[TOC]", "[!NOTE]", "\\tag", "front matter"):
            self.assertIn(feature, guide.text)

    async def test_prompt_is_listed_with_an_optional_argument(self) -> None:
        async with client_session(server._mcp_server) as client:
            listed = await client.list_prompts()

        prompt = next(p for p in listed.prompts if p.name == "prepare_markdown_for_word")
        self.assertEqual([a.name for a in prompt.arguments], ["topic_or_path"])
        self.assertFalse(prompt.arguments[0].required)

    async def test_prompt_renders_the_workflow_and_the_guide(self) -> None:
        async with client_session(server._mcp_server) as client:
            targeted = await client.get_prompt(
                "prepare_markdown_for_word", {"topic_or_path": "/data/report.md"}
            )
            generic = await client.get_prompt("prepare_markdown_for_word", {})

        text = targeted.messages[0].content.text
        self.assertIn("/data/report.md", text)
        positions = [
            text.index(tool)
            for tool in ("check_latex", "preview_markdown", "markdown_to_word", "inspect_docx")
        ]
        self.assertEqual(positions, sorted(positions))
        self.assertIn("# Markdown that converts cleanly to Word", text)
        self.assertTrue(generic.messages[0].content.text)


class StdioProtocolTests(unittest.TestCase):
    def test_importing_the_server_writes_nothing_to_stdout(self) -> None:
        # stdout — это канал stdio-протокола: одна лишняя строка при импорте
        # рвёт JSON-RPC сессию, и клиент видит нечитаемую ошибку парсинга.
        import subprocess
        import sys

        repo_root = Path(__file__).resolve().parent.parent
        result = subprocess.run(
            [sys.executable, "-c", "import mdtoword.mcp_server"],
            capture_output=True,
            text=True,
            cwd=repo_root,
        )

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, "")


if __name__ == "__main__":
    unittest.main()
