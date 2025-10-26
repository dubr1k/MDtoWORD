import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import ttk
import os
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageTk

if sys.platform == "darwin":  # 'darwin' — это Mac
    from tkmacosx import Button
else:
    from tkinter import Button

class MarkdownToWordConverter:
    """Класс для конвертации Markdown в Word"""

    def __init__(self):
        self.doc = None
        self.default_font_name = 'Times New Roman'
        self.default_font_size = Pt(12)

    def create_document(self):
        """Создает новый документ Word"""
        self.doc = Document()
        # Устанавливаем стандартный стиль
        style = self.doc.styles['Normal']
        font = style.font
        font.name = self.default_font_name
        font.size = self.default_font_size
        font.color.rgb = RGBColor(0, 0, 0)  # Черный цвет

    def add_table_borders(self, table):
        """Добавляет границы к таблице"""
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)

        # Создаем элемент границ
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)

        tblPr.append(tblBorders)

    def parse_inline_formatting(self, text):
        """Парсит встроенное форматирование (жирный, курсив, код)"""
        # Возвращает список кортежей (текст, форматирование)
        # форматирование: {'bold': bool, 'italic': bool, 'code': bool}
        parts = []

        # Регулярное выражение для поиска форматирования
        pattern = r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|`.*?`|___.*?___|__.*?__|_.*?_)'

        last_end = 0
        for match in re.finditer(pattern, text):
            # Добавляем текст до совпадения
            if match.start() > last_end:
                parts.append((text[last_end:match.start()], {}))

            matched_text = match.group()
            formatting = {}
            clean_text = matched_text

            # Проверяем тип форматирования
            if matched_text.startswith('***') or matched_text.startswith('___'):
                formatting = {'bold': True, 'italic': True}
                clean_text = matched_text[3:-3]
            elif matched_text.startswith('**') or matched_text.startswith('__'):
                formatting = {'bold': True}
                clean_text = matched_text[2:-2]
            elif matched_text.startswith('*') or matched_text.startswith('_'):
                formatting = {'italic': True}
                clean_text = matched_text[1:-1]
            elif matched_text.startswith('`'):
                formatting = {'code': True}
                clean_text = matched_text[1:-1]

            parts.append((clean_text, formatting))
            last_end = match.end()

        # Добавляем оставшийся текст
        if last_end < len(text):
            parts.append((text[last_end:], {}))

        return parts if parts else [(text, {})]

    def add_formatted_text(self, paragraph, text):
        """Добавляет текст с форматированием в параграф"""
        parts = self.parse_inline_formatting(text)

        for part_text, formatting in parts:
            run = paragraph.add_run(part_text)
            run.font.name = self.default_font_name
            run.font.size = self.default_font_size
            run.font.color.rgb = RGBColor(0, 0, 0)

            if formatting.get('bold'):
                run.bold = True
            if formatting.get('italic'):
                run.italic = True
            if formatting.get('code'):
                run.font.name = 'Courier New'
                run.font.size = Pt(10)

    def process_table(self, lines, start_idx):
        """Обрабатывает таблицу из markdown"""
        table_lines = []
        idx = start_idx

        # Собираем все строки таблицы
        while idx < len(lines) and '|' in lines[idx]:
            table_lines.append(lines[idx])
            idx += 1

        if len(table_lines) < 2:
            return idx

        # Парсим таблицу
        rows = []
        for line in table_lines:
            # Пропускаем разделительную строку (---|---|---)
            if re.match(r'^\|[\s\-:|]+\|$', line.strip()):
                continue

            # Разбиваем строку на ячейки
            cells = [cell.strip() for cell in line.split('|')]
            # Удаляем пустые ячейки в начале и конце
            cells = [c for c in cells if c or cells.index(c) not in [0, len(cells) - 1]]
            if cells:
                rows.append(cells)

        if not rows:
            return idx

        # Создаем таблицу в документе
        table = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = 'Table Grid'
        self.add_table_borders(table)

        # Заполняем таблицу
        for i, row_data in enumerate(rows):
            for j, cell_text in enumerate(row_data):
                if j < len(table.rows[i].cells):
                    cell = table.rows[i].cells[j]
                    # Очищаем ячейку и добавляем форматированный текст
                    cell.text = ''
                    paragraph = cell.paragraphs[0]
                    self.add_formatted_text(paragraph, cell_text)
                    # Устанавливаем выравнивание по ширине для текста в ячейках
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                    # Форматирование для заголовка таблицы (первая строка)
                    if i == 0:
                        for run in paragraph.runs:
                            run.bold = True

        return idx

    def process_list(self, lines, start_idx):
        """Обрабатывает списки (маркированные и нумерованные)"""
        idx = start_idx
        list_items = []

        # Определяем тип списка
        first_line = lines[start_idx].strip()
        is_ordered = bool(re.match(r'^\d+\.', first_line))

        # Собираем элементы списка
        while idx < len(lines):
            line = lines[idx].strip()
            if not line:
                break

            # Проверяем маркированный список
            if re.match(r'^[-*+]\s', line):
                list_items.append(('unordered', line[2:]))
                idx += 1
            # Проверяем нумерованный список
            elif re.match(r'^\d+\.\s', line):
                match = re.match(r'^\d+\.\s(.*)', line)
                list_items.append(('ordered', match.group(1)))
                idx += 1
            else:
                break

        # Добавляем элементы списка в документ
        for list_type, item_text in list_items:
            paragraph = self.doc.add_paragraph(style='List Bullet' if list_type == 'unordered' else 'List Number')
            self.add_formatted_text(paragraph, item_text)

        return idx

    def convert_file(self, input_path, output_path):
        """Конвертирует markdown файл в Word"""
        try:
            # Читаем файл
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Создаем документ
            self.create_document()

            # Разбиваем на строки
            lines = content.split('\n')

            i = 0
            while i < len(lines):
                line = lines[i]
                stripped = line.strip()

                # Пропускаем пустые строки
                if not stripped:
                    i += 1
                    continue

                # Заголовки
                if stripped.startswith('#'):
                    level = 0
                    while level < len(stripped) and stripped[level] == '#':
                        level += 1

                    title_text = stripped[level:].strip()
                    heading_style = f'Heading {min(level, 9)}'

                    paragraph = self.doc.add_paragraph(style=heading_style)
                    self.add_formatted_text(paragraph, title_text)
                    paragraph.paragraph_format.space_after = Pt(6)
                    i += 1

                # Горизонтальная линия
                elif stripped == '---' or stripped == '***' or stripped == '___':
                    self.doc.add_paragraph('_' * 50)
                    i += 1

                # Таблицы
                elif '|' in line:
                    i = self.process_table(lines, i)

                # Списки
                elif re.match(r'^[-*+]\s', stripped) or re.match(r'^\d+\.\s', stripped):
                    i = self.process_list(lines, i)

                # Блок кода (```)
                elif stripped.startswith('```'):
                    i += 1
                    code_lines = []
                    while i < len(lines) and not lines[i].strip().startswith('```'):
                        code_lines.append(lines[i])
                        i += 1

                    if code_lines:
                        paragraph = self.doc.add_paragraph()
                        run = paragraph.add_run('\n'.join(code_lines))
                        run.font.name = 'Courier New'
                        run.font.size = Pt(10)
                        run.font.color.rgb = RGBColor(0, 0, 0)

                    i += 1

                # Обычный текст
                else:
                    paragraph = self.doc.add_paragraph()
                    self.add_formatted_text(paragraph, line)
                    # Устанавливаем выравнивание по ширине для обычного текста
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    i += 1

            # Сохраняем документ
            self.doc.save(output_path)
            return True, "Успешно конвертировано"

        except Exception as e:
            return False, f"Ошибка при конвертации: {str(e)}"


class WordToMarkdownConverter:
    """Класс для конвертации Word в Markdown"""

    def __init__(self):
        pass

    def convert_file(self, input_path, output_path):
        """Конвертирует Word файл в Markdown"""
        try:
            doc = Document(input_path)

            markdown_lines = []

            for paragraph in doc.paragraphs:
                text = paragraph.text

                # Пропускаем пустые параграфы
                if not text.strip():
                    markdown_lines.append('')
                    continue

                # Определяем уровень заголовка по стилю
                heading_level = 0
                if paragraph.style.name.startswith('Heading '):
                    try:
                        heading_level = int(paragraph.style.name.split()[-1])
                        if 1 <= heading_level <= 6:
                            text = '#' * heading_level + ' ' + text
                    except ValueError:
                        pass  # Если не удалось определить уровень, оставляем как есть

                # Если не заголовок, обрабатываем как обычный текст
                if heading_level == 0:
                    # Простое форматирование текста (жирный, курсив) через run'ы
                    formatted_text = ""
                    for run in paragraph.runs:
                        run_text = run.text
                        if run.bold and run.italic:
                            formatted_text += f'***{run_text}***'
                        elif run.bold:
                            formatted_text += f'**{run_text}**'
                        elif run.italic:
                            formatted_text += f'*{run_text}*'
                        elif run_text.strip().startswith('`') and run_text.strip().endswith('`'):
                            formatted_text += f'`{run_text}`'  # Код (если run весь в кавычках)
                        else:
                            formatted_text += run_text

                    text = formatted_text

                    # Обработка выравнивания (если нужно)
                    # paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY - это для Word, в MD выравнивание не поддерживается стандартно

                markdown_lines.append(text)

            # Обработка таблиц (упрощённо)
            for table in doc.tables:
                if table.rows:
                    markdown_lines.append('') # Пустая строка перед таблицей
                    # Заголовки (первый ряд)
                    header_cells = [cell.text.strip() for cell in table.rows[0].cells]
                    markdown_lines.append('| ' + ' | '.join(header_cells) + ' |')
                    # Разделитель
                    markdown_lines.append('| ' + ' | '.join(['---'] * len(header_cells)) + ' |')
                    # Остальные строки
                    for row in table.rows[1:]:
                        row_cells = [cell.text.strip() for cell in row.cells]
                        markdown_lines.append('| ' + ' | '.join(row_cells) + ' |')
                    markdown_lines.append('') # Пустая строка после таблицы

            # Записываем результат в файл
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(markdown_lines))

            return True, "Успешно конвертировано"

        except Exception as e:
            return False, f"Ошибка при конвертации: {str(e)}"


class ConverterGUI:
    """GUI для конвертера"""

    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Markdown в Word")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        try:
            pil_image = Image.open('ico.png')
            icon = ImageTk.PhotoImage(pil_image)
            self.root.iconphoto(True, icon)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")

        self.selected_files = []
        self.output_directory = ""
        self.converter = MarkdownToWordConverter() # Конвертер по умолчанию
        self.current_converter_type = "md_to_word" # Тип конвертера

        # Популярные шрифты (для режима MD->Word)
        self.fonts = [
            "Arial", "Times New Roman", "Courier New", "Helvetica", "Verdana",
            "Georgia", "Palatino", "Garamond", "Comic Sans MS", "Trebuchet MS",
            "Arial Black", "Impact", "Calibri", "Cambria", "Candara",
            "Consolas", "Segoe UI", "Roboto", "Open Sans", "Lato",
            "Montserrat", "Source Sans Pro", "Noto Sans", "Ubuntu", "Merriweather",
            "Lora", "Poppins", "Inter", "Fira Sans", "PT Sans",
            "Droid Sans", "Nunito", "Raleway", "Oswald", "Roboto Condensed",
            "Titillium Web", "Exo 2", "Crimson Text", "Playfair Display", "Quicksand",
            "Comfortaa", "Pacifico", "Indie Flower", "Dancing Script", "Kaushan Script",
            "Courgette", "Satisfy", "Handlee", "Bangers", "Chewy",
            "Fredoka One", "Architects Daughter", "Nova Square", "Orbitron", "Press Start 2P",
            "Monoton", "VT323", "Cutive Mono", "Inconsolata", "Space Mono",
            "Fira Mono", "Source Code Pro", "Roboto Mono", "IBM Plex Mono", "JetBrains Mono",
            "Cascadia Code", "Ubuntu Mono", "PT Mono", "Noto Sans Mono", "Roboto Slab",
            "Crimson Pro", "Barlow", "Nunito Sans", "Work Sans", "Public Sans",
            "Red Hat Display", "Red Hat Text", "Atkinson Hyperlegible", "Literata", "Charter",
            "Avenir Next", "Proxima Nova", "Myriad Pro", "Gill Sans", "Franklin Gothic",
            "Optima", "Futura", "Univers", "Helvetica Neue", "Avenir",
            "Geneva", "Tahoma", "Trebuchet MS", "Lucida Grande", "Lucida Sans Unicode",
            "Bitstream Vera Sans", "DejaVu Sans", "Liberation Sans", "Noto Sans CJK",
            "Source Han Sans", "Noto Serif CJK", "Source Han Serif", "SimSun", "SimHei",
            "Microsoft YaHei", "Microsoft JhengHei", "Meiryo", "Yu Gothic",
            "Hiragino Kaku Gothic Pro", "Apple Gothic", "Malgun Gothic", "Batang", "Dotum",
            "Arial Unicode MS", "Lucida Sans", "Segoe UI Symbol", "Symbol", "Webdings",
            "Wingdings", "Wingdings 2", "Wingdings 3",
        ]

        # --- Добавление словаря переводов ---
        self.translations = {
            "ru": {
                "title": "Конвертер",
                "title_md_to_word": "Конвертер Markdown в Word",
                "title_word_to_md": "Конвертер Word в Markdown",
                "settings_frame": "Настройки документа",
                "font_label": "Шрифт:",
                "width_label": "Ширина текста (пт):",
                "files_frame": "Выбор файлов",
                "files_frame_md": "Выбор файлов Markdown",
                "files_frame_word": "Выбор файлов Word",
                "select_button": "Выбрать файлы .md",
                "select_button_md": "Выбрать файлы .md",
                "select_button_word": "Выбрать файлы .docx",
                "remove_button": "Удалить выбранные",
                "remove_all_button": "Удалить все",
                "output_frame": "Место сохранения",
                "output_label_default": "Папка не выбрана",
                "output_button": "Выбрать папку",
                "status_ready": "Готов к конвертации",
                "convert_button": "Конвертировать",
                "toggle_button_md_to_word": "Режим: MD -> Word",
                "toggle_button_word_to_md": "Режим: Word -> MD",
                "warning_no_files": "Выберите файлы для конвертации",
                "warning_no_dir": "Выберите папку для сохранения",
                "success_message": "Все файлы успешно конвертированы!\nВсего: {count}\nШрифт: {font}, Размер: {size} pt",
                "success_message_word": "Все файлы успешно конвертированы!\nВсего: {count}",
                "status_converting": "Конвертация: {filename}",
                "status_finished": "Конвертация завершена",
                "error_title": "Конвертация завершена с ошибками",
                "error_message": "Успешно: {success}\nОшибок: {errors}\n\n{details}",
                "font_changed": "Шрифт изменен на: {font}",
                "size_changed": "Размер шрифта изменен на: {size} pt",
                "invalid_size": "Некорректное значение ширины (размера шрифта)",
                "files_selected": "Выбрано файлов: {count}",
                "icon_failed": "Не удалось установить иконку: {error}",
                "logo_failed": "Не удалось загрузить логотип: {error}",
            },
            "en": {
                "title": "Converter",
                "title_md_to_word": "Markdown to Word Converter",
                "title_word_to_md": "Word to Markdown Converter",
                "settings_frame": "Document Settings",
                "font_label": "Font:",
                "width_label": "Text Width (pt):",
                "files_frame": "Select Files",
                "files_frame_md": "Select Markdown Files",
                "files_frame_word": "Select Word Files",
                "select_button": "Select .md Files",
                "select_button_md": "Select .md Files",
                "select_button_word": "Select .docx Files",
                "remove_button": "Remove Selected",
                "remove_all_button": "Remove All",
                "output_frame": "Save Location",
                "output_label_default": "Folder not selected",
                "output_button": "Select Folder",
                "status_ready": "Ready to convert",
                "convert_button": "Convert",
                "toggle_button_md_to_word": "Mode: MD -> Word",
                "toggle_button_word_to_md": "Mode: Word -> MD",
                "warning_no_files": "Select files for conversion",
                "warning_no_dir": "Select a folder to save",
                "success_message": "All files converted successfully!\nTotal: {count}\nFont: {font}, Size: {size} pt",
                "success_message_word": "All files converted successfully!\nTotal: {count}",
                "status_converting": "Converting: {filename}",
                "status_finished": "Conversion finished",
                "error_title": "Conversion completed with errors",
                "error_message": "Success: {success}\nErrors: {errors}\n\n{details}",
                "font_changed": "Font changed to: {font}",
                "size_changed": "Font size changed to: {size} pt",
                "invalid_size": "Invalid width (font size) value",
                "files_selected": "Files selected: {count}",
                "icon_failed": "Failed to set icon: {error}",
                "logo_failed": "Failed to load logo: {error}",
            }
        }
        self.current_language = "ru"  # Язык по умолчанию

        self.create_widgets()

    def create_widgets(self):
        """Создает виджеты интерфейса"""
        # Заголовок
        self.title_label = tk.Label(
            self.root,
            text=self.translations[self.current_language]["title_md_to_word"],
            font=("Arial", 16, "bold")
        )
        self.title_label.pack(pady=10)

        # --- ФРЕЙМ ДЛЯ НАСТРОЕК (теперь выше) ---
        self.settings_frame = tk.LabelFrame(self.root, text=self.translations[self.current_language]["settings_frame"],
                                            padx=10, pady=10)
        # settings_frame.pack(padx=10, pady=5, fill=tk.X) # Пока не пакуем, сделаем видимым/невидимым позже

        # Настройка шрифта
        self.font_label = tk.Label(self.settings_frame, text=self.translations[self.current_language]["font_label"])
        self.font_label.grid(row=0, column=0, sticky="w", padx=(0, 5))

        self.font_var = tk.StringVar(value=self.converter.default_font_name)
        self.font_combobox = ttk.Combobox(
            self.settings_frame,
            textvariable=self.font_var,
            values=self.fonts,
            state="readonly",
            width=20
        )
        self.font_combobox.grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.font_combobox.bind("<<ComboboxSelected>>", self.on_font_change)

        # Настройка ширины текста (через Spinbox)
        self.width_label = tk.Label(self.settings_frame, text=self.translations[self.current_language]["width_label"])
        self.width_label.grid(row=0, column=2, sticky="w", padx=(10, 5))

        self.width_var = tk.IntVar(value=int(self.converter.default_font_size.pt))
        self.width_spinbox = tk.Spinbox(
            self.settings_frame,
            from_=1,  # Минимальный размер шрифта
            to=100,  # Максимальный размер шрифта
            textvariable=self.width_var,
            width=5,
            command=self.on_width_change  # Вызывается при изменении значения
        )
        self.width_spinbox.grid(row=0, column=3, sticky="w", padx=(0, 10))

        # --- ФРЕЙМ ДЛЯ ВЫБОРА ФАЙЛОВ (теперь ниже настроек) ---
        self.files_frame = tk.LabelFrame(self.root, text=self.translations[self.current_language]["files_frame_md"],
                                         padx=10, pady=10)
        self.files_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Кнопка выбора файлов
        self.select_button = Button(
            self.files_frame,
            text=self.translations[self.current_language]["select_button_md"],
            command=self.select_files,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        self.select_button.pack(pady=5)

        # Список выбранных файлов
        list_frame = tk.Frame(self.files_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.files_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 9),
            selectmode=tk.EXTENDED
        )
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_listbox.yview)


        # Кнопки удаления
        remove_buttons_frame = tk.Frame(self.files_frame)
        remove_buttons_frame.pack(pady=5)  # Отступ от списка файлов

        # Кнопка удаления выбранных файлов
        self.remove_button = Button(
            remove_buttons_frame,
            text=self.translations[self.current_language]["remove_button"],
            command=self.remove_selected_files,
            bg="#f44336",
            fg="white",
            font=("Arial", 9),
            padx=5,
            pady=3
        )
        self.remove_button.pack(side=tk.LEFT, padx=(0, 5))

        # Кнопка удаления всех файлов
        self.remove_button_all = Button(
            remove_buttons_frame,
            text=self.translations[self.current_language]["remove_all_button"],
            command=self.remove_all_files,
            bg="#d32f2f",
            fg="white",
            font=("Arial", 9),
            padx=5,
            pady=3
        )
        self.remove_button_all.pack(side=tk.LEFT, padx=(0, 0))  # Рядом с первой кнопкой
        # Фрейм для выбора папки сохранения
        self.output_frame = tk.LabelFrame(self.root, text=self.translations[self.current_language]["output_frame"],
                                          padx=10, pady=10)
        self.output_frame.pack(padx=10, pady=5, fill=tk.X)

        output_inner_frame = tk.Frame(self.output_frame)
        output_inner_frame.pack(fill=tk.X)

        self.output_label = tk.Label(
            output_inner_frame,
            text=self.translations[self.current_language]["output_label_default"],
            anchor="w",
            bg="gray",
            relief=tk.SUNKEN,
            padx=5,
            pady=5
        )
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.output_button = Button(
            output_inner_frame,
            text=self.translations[self.current_language]["output_button"],
            command=self.select_output_directory,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        self.output_button.pack(side=tk.RIGHT)

        # Прогресс бар
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(padx=10, pady=5, fill=tk.X)

        self.progress = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            length=100,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X)

        self.status_label = tk.Label(
            self.root,
            text=self.translations[self.current_language]["status_ready"],
            font=("Arial", 9),
            fg="gray"
        )
        self.status_label.pack(pady=5)

        # Кнопка конвертации
        self.convert_button = Button(
            self.root,
            text=self.translations[self.current_language]["convert_button"],
            command=self.convert_files,
            bg="#FF9800",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10
        )
        self.convert_button.pack(pady=10)


        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=5, side=tk.BOTTOM, fill=tk.X)

        # Кнопка переключения режима
        self.toggle_button = Button(
            bottom_frame,
            text=self.translations[self.current_language]["toggle_button_md_to_word"],
            command=self.toggle_converter_type,
            bg="#670067",
            fg="white",
            font=("Arial", 12),
            padx=10,
            pady=3
        )
        self.toggle_button.pack(side=tk.LEFT, padx=(0, 5))

        # Кнопка переключения языка
        self.language_button = Button(
            bottom_frame,
            text="EN",
            command=self.toggle_language,
            bg="#9E9E9E",
            fg="white",
            font=("Arial", 12),
            padx=10,
            pady=3
        )
        self.language_button.pack(side=tk.RIGHT, padx=(5, 0))
        # Инициализируем видимость элементов в зависимости от режима
        self.update_ui_for_converter_type()

    def toggle_converter_type(self):
        """Переключает между режимами MD->Word и Word->MD"""
        if self.current_converter_type == "md_to_word":
            self.current_converter_type = "word_to_md"
            self.converter = WordToMarkdownConverter() # Меняем экземпляр конвертера
        else:
            self.current_converter_type = "md_to_word"
            self.converter = MarkdownToWordConverter() # Меняем экземпляр конвертера

        self.update_ui_for_converter_type()

    def update_ui_for_converter_type(self):
        """Обновляет интерфейс в зависимости от текущего режима конвертации"""
        t = self.translations[self.current_language]

        if self.current_converter_type == "md_to_word":
            self.root.title(t["title_md_to_word"])
            self.title_label.config(text=t["title_md_to_word"])
            self.toggle_button.config(text=t["toggle_button_md_to_word"])
            self.files_frame.config(text=t["files_frame_md"])
            self.select_button.config(text=t["select_button_md"])
            # Показываем настройки
            self.settings_frame.pack(padx=10, pady=5, fill=tk.X, before=self.files_frame)  # pack с before
        else:  # word_to_md
            self.root.title(t["title_word_to_md"])
            self.title_label.config(text=t["title_word_to_md"])
            self.toggle_button.config(text=t["toggle_button_word_to_md"])
            self.files_frame.config(text=t["files_frame_word"])
            self.select_button.config(text=t["select_button_word"])
            # Скрываем настройки (убираем из менеджера упаковки)
            self.settings_frame.pack_forget()


    def toggle_language(self):
        """Переключает язык интерфейса между русским и английским"""
        if self.current_language == "ru":
            self.current_language = "en"
            self.language_button.config(text="RU")
        else:
            self.current_language = "ru"
            self.language_button.config(text="EN")

        self.update_ui_text()

    def update_ui_text(self):
        """Обновляет текст всех виджетов в соответствии с текущим языком"""
        t = self.translations[self.current_language]
        # Обновляем заголовки и текст в зависимости от режима
        if self.current_converter_type == "md_to_word":
             self.root.title(t["title_md_to_word"])
             self.title_label.config(text=t["title_md_to_word"])
             self.files_frame.config(text=t["files_frame_md"])
             self.select_button.config(text=t["select_button_md"])
             # Обновляем текст кнопки переключения режима на "Mode: Word -> MD"
             self.toggle_button.config(text=t["toggle_button_md_to_word"])
        else: # word_to_md
             self.root.title(t["title_word_to_md"])
             self.title_label.config(text=t["title_word_to_md"])
             self.files_frame.config(text=t["files_frame_word"])
             self.select_button.config(text=t["select_button_word"])
             # Обновляем текст кнопки переключения режима на "Mode: MD -> Word"
             self.toggle_button.config(text=t["toggle_button_word_to_md"])

        # Обновляем текст кнопки смены языка (в зависимости от текущего языка)
        if self.current_language == "ru":
            self.language_button.config(text="EN")
        else: # en
            self.language_button.config(text="RU")

        # Кнопки удаления
        self.remove_button.config(text=t["remove_button"])
        self.remove_button_all.config(text=t["remove_all_button"])


        self.output_frame.config(text=t["output_frame"])
        # Если папка не выбрана, обновляем текст на "Folder not selected" / "Папка не выбрана"
        if self.output_directory == "":
             self.output_label.config(text=t["output_label_default"])
        self.output_button.config(text=t["output_button"])
        # Обновляем текст статуса, если он соответствует одному из стандартных
        current_status = self.status_label.cget("text")
        if current_status == self.translations["ru"]["status_ready"] or current_status == self.translations["en"]["status_ready"]:
            self.status_label.config(text=t["status_ready"])
        elif current_status == self.translations["ru"]["status_finished"] or current_status == self.translations["en"]["status_finished"]:
            self.status_label.config(text=t["status_finished"])
        self.convert_button.config(text=t["convert_button"])


    def on_font_change(self, event):
        """Обновляет шрифт конвертера при выборе из Combobox (только для MD->Word)"""
        if self.current_converter_type == "md_to_word":
            selected_font = self.font_var.get()
            self.converter.default_font_name = selected_font
            t = self.translations[self.current_language]
            print(t["font_changed"].format(font=selected_font))

    def on_width_change(self):
        """Обновляет размер шрифта конвертера при изменении Spinbox (только для MD->Word)"""
        if self.current_converter_type == "md_to_word":
            try:
                selected_size = self.width_var.get()
                self.converter.default_font_size = Pt(selected_size)
                t = self.translations[self.current_language]
                print(t["size_changed"].format(size=selected_size))
            except tk.TclError:
                # Обработка случая, когда введено некорректное значение
                t = self.translations[self.current_language]
                print(t["invalid_size"])
                # Восстанавливаем предыдущее корректное значение
                self.width_var.set(int(self.converter.default_font_size.pt))

    def select_files(self):
        """Выбор файлов (в зависимости от режима)"""
        t = self.translations[self.current_language]
        if self.current_converter_type == "md_to_word":
            filetypes = [("Markdown files", "*.md"), ("All files", "*.*")]
            title = t["select_button_md"]
        else: # word_to_md
            filetypes = [("Word files", "*.docx"), ("All files", "*.*")]
            title = t["select_button_word"]

        files = filedialog.askopenfilenames(
            title=title,
            filetypes=filetypes
        )

        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))

        self.status_label.config(text=t["files_selected"].format(count=len(self.selected_files)))

    def remove_selected_files(self):
        """Удаление выбранных файлов из списка"""
        selected_indices = self.files_listbox.curselection()

        # Удаляем в обратном порядке, чтобы индексы не сбивались
        for index in reversed(selected_indices):
            self.files_listbox.delete(index)
            del self.selected_files[index]

        t = self.translations[self.current_language]
        self.status_label.config(text=t["files_selected"].format(count=len(self.selected_files)))

    def remove_all_files(self):
        """Удаляет все файлы из списка"""
        # Очищаем список в интерфейсе
        self.files_listbox.delete(0, tk.END)
        # Очищаем внутренний список
        self.selected_files.clear()
        # Обновляем статус
        t = self.translations[self.current_language]
        self.status_label.config(text=t["files_selected"].format(count=0))

    def select_output_directory(self):
        """Выбор папки для сохранения"""
        t = self.translations[self.current_language]
        directory = filedialog.askdirectory(title=t["output_button"])

        if directory:
            self.output_directory = directory
            self.output_label.config(text=directory)

    def convert_files(self):
        """Конвертация выбранных файлов"""
        t = self.translations[self.current_language]
        if not self.selected_files:
            messagebox.showwarning("Предупреждение", t["warning_no_files"])
            return

        if not self.output_directory:
            messagebox.showwarning("Предупреждение", t["warning_no_dir"])
            return

        # Сброс прогресса
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.selected_files)

        success_count = 0
        errors = []

        for i, input_file in enumerate(self.selected_files):
            # Обновляем статус
            filename = os.path.basename(input_file)
            self.status_label.config(text=t["status_converting"].format(filename=filename))
            self.root.update()

            # Формируем имя выходного файла в зависимости от режима
            input_path_obj = Path(input_file)
            if self.current_converter_type == "md_to_word":
                output_filename = input_path_obj.stem + ".docx"
            else: # word_to_md
                output_filename = input_path_obj.stem + ".md"

            output_path = os.path.join(self.output_directory, output_filename)

            # Конвертируем
            success, message = self.converter.convert_file(input_file, output_path)

            if success:
                success_count += 1
            else:
                errors.append(f"{filename}: {message}")

            # Обновляем прогресс
            self.progress['value'] = i + 1
            self.root.update()

        # Показываем результат
        if errors:
            error_text = "\n".join(errors)
            messagebox.showwarning(
                t["error_title"],
                t["error_message"].format(success=success_count, errors=len(errors), details=error_text)
            )
        else:
            # Выводим разные сообщения в зависимости от режима
            if self.current_converter_type == "md_to_word":
                msg = t["success_message"].format(count=success_count, font=self.converter.default_font_name, size=self.converter.default_font_size.pt)
            else: # word_to_md
                msg = t["success_message_word"].format(count=success_count)
            messagebox.showinfo(
                "Успех",
                msg
            )

        self.status_label.config(text=t["status_finished"])


def main():
    """Главная функция"""
    root = tk.Tk()
    app = ConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()