"""Тексты интерфейса на двух языках и форматирование предупреждений."""

from __future__ import annotations

TRANSLATIONS: dict[str, dict[str, str]] = {
    "ru": {
        "title_md": "Markdown → Word", "title_word": "Word → Markdown",
        "settings": "Оформление документа", "font": "Шрифт", "size": "Размер",
        "drop_md": "Перетащите файлы или папки Markdown сюда",
        "drop_word": "Перетащите файлы или папки Word сюда",
        "add_files": "Добавить файлы", "add_folder": "Добавить папку",
        "remove": "Удалить выбранные", "clear": "Очистить очередь",
        "files_tab": "Файлы", "text_tab": "Текст", "text_label": "Введите Markdown-текст",
        "output": "Место сохранения", "output_auto": "Рядом с исходными файлами",
        "choose_output": "Выбрать папку", "reset_output": "Сбросить",
        "ready": "Готово к конвертации", "queued": "В очереди: {count}",
        "converting": "Конвертация: {filename}",
        "finished": "Конвертация завершена", "convert": "Конвертировать",
        "toggle_md": "Режим: MD → Word", "toggle_word": "Режим: Word → MD",
        "theme_dark": "Тёмная тема · Переключить на светлую",
        "theme_light": "Светлая тема · Переключить на тёмную",
        "no_files": "Добавьте файлы или папку для конвертации",
        "empty_text": "Введите текст для конвертации", "save_as": "Сохранить как",
        "errors": "Конвертация завершена с ошибками", "result": "Готово: {success}\nОшибок: {errors}",
        "footnotes_heading": "Сноски",
        "preset": "Стиль", "preset_default": "Обычный", "preset_gost": "ГОСТ 7.32",
        "toc": "Оглавление",
        "converted_ok": "Успешно конвертировано",
        "convert_failed": "Ошибка при конвертации: {error}",
    },
    "en": {
        "title_md": "Markdown → Word", "title_word": "Word → Markdown",
        "settings": "Document appearance", "font": "Font", "size": "Size",
        "drop_md": "Drop Markdown files or folders here",
        "drop_word": "Drop Word files or folders here",
        "add_files": "Add files", "add_folder": "Add folder",
        "remove": "Remove selected", "clear": "Clear queue",
        "files_tab": "Files", "text_tab": "Text", "text_label": "Enter Markdown text",
        "output": "Save location", "output_auto": "Next to each source file",
        "choose_output": "Choose folder", "reset_output": "Reset",
        "ready": "Ready to convert", "queued": "In queue: {count}",
        "converting": "Converting: {filename}",
        "finished": "Conversion finished", "convert": "Convert",
        "toggle_md": "Mode: MD → Word", "toggle_word": "Mode: Word → MD",
        "theme_dark": "Dark theme · Switch to light",
        "theme_light": "Light theme · Switch to dark",
        "no_files": "Add files or a folder to convert",
        "empty_text": "Enter text to convert", "save_as": "Save as",
        "errors": "Conversion completed with errors", "result": "Complete: {success}\nErrors: {errors}",
        "footnotes_heading": "Footnotes",
        "preset": "Style", "preset_default": "Standard", "preset_gost": "GOST 7.32",
        "toc": "Table of contents",
        "converted_ok": "Converted successfully",
        "convert_failed": "Conversion failed: {error}",
    },
}


def describe_warning(prefix: str, warning: str) -> str:
    """«файл.md:12: текст» — с номером строки, если рендерер его знает."""
    line = getattr(warning, "line", None)
    location = f"{prefix}:{line}" if line else prefix
    return f"{location}: {warning}" if location else str(warning)
