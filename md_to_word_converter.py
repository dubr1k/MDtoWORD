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
        self.converter = MarkdownToWordConverter()

        # Популярные шрифты
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

        self.create_widgets()

    def create_widgets(self):
        """Создает виджеты интерфейса"""
        # Заголовок
        title_label = tk.Label(
            self.root,
            text="Конвертер Markdown в Word",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=10)

        # Фрейм для настроек шрифта и ширины
        settings_frame = tk.LabelFrame(self.root, text="Настройки документа", padx=10, pady=10)
        settings_frame.pack(padx=10, pady=5, fill=tk.X)

        # Настройка шрифта
        font_label = tk.Label(settings_frame, text="Шрифт:")
        font_label.grid(row=0, column=0, sticky="w", padx=(0, 5))

        self.font_var = tk.StringVar(value=self.converter.default_font_name)
        self.font_combobox = ttk.Combobox(
            settings_frame,
            textvariable=self.font_var,
            values=self.fonts,
            state="readonly",
            width=20
        )
        self.font_combobox.grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.font_combobox.bind("<<ComboboxSelected>>", self.on_font_change)

        # Настройка ширины текста (через Spinbox)
        width_label = tk.Label(settings_frame, text="Ширина текста (пт):")
        width_label.grid(row=0, column=2, sticky="w", padx=(10, 5))

        self.width_var = tk.IntVar(value=int(self.converter.default_font_size.pt))
        self.width_spinbox = tk.Spinbox(
            settings_frame,
            from_=8,  # Минимальный размер шрифта
            to=72,  # Максимальный размер шрифта
            textvariable=self.width_var,
            width=5,
            command=self.on_width_change  # Вызывается при изменении значения
        )
        self.width_spinbox.grid(row=0, column=3, sticky="w", padx=(0, 10))

        # Фрейм для выбора файлов
        files_frame = tk.LabelFrame(self.root, text="Выбор файлов Markdown", padx=10, pady=10)
        files_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Кнопка выбора файлов
        select_button = Button(
            files_frame,
            text="Выбрать файлы .md",
            command=self.select_files,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        select_button.pack(pady=5)

        # Список выбранных файлов
        list_frame = tk.Frame(files_frame)
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

        # Кнопка удаления выбранных файлов
        remove_button = Button(
            files_frame,
            text="Удалить выбранные",
            command=self.remove_selected_files,
            bg="#f44336",
            fg="white",
            font=("Arial", 9),
            padx=5,
            pady=3
        )
        remove_button.pack(pady=5)

        # Фрейм для выбора папки сохранения
        output_frame = tk.LabelFrame(self.root, text="Место сохранения", padx=10, pady=10)
        output_frame.pack(padx=10, pady=5, fill=tk.X)

        output_inner_frame = tk.Frame(output_frame)
        output_inner_frame.pack(fill=tk.X)

        self.output_label = tk.Label(
            output_inner_frame,
            text="Папка не выбрана",
            anchor="w",
            bg="gray",
            relief=tk.SUNKEN,
            padx=5,
            pady=5
        )
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        output_button = Button(
            output_inner_frame,
            text="Выбрать папку",
            command=self.select_output_directory,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        output_button.pack(side=tk.RIGHT)

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
            text="Готов к конвертации",
            font=("Arial", 9),
            fg="gray"
        )
        self.status_label.pack(pady=5)

        # Кнопка конвертации
        convert_button = Button(
            self.root,
            text="Конвертировать",
            command=self.convert_files,
            bg="#FF9800",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10
        )
        convert_button.pack(pady=10)

    def on_font_change(self, event):
        """Обновляет шрифт конвертера при выборе из Combobox"""
        selected_font = self.font_var.get()
        self.converter.default_font_name = selected_font
        print(f"Шрифт изменен на: {selected_font}")

    def on_width_change(self):
        """Обновляет размер шрифта конвертера при изменении Spinbox"""
        try:
            selected_size = self.width_var.get()
            self.converter.default_font_size = Pt(selected_size)
            print(f"Размер шрифта изменен на: {selected_size} pt")
        except tk.TclError:
            # Обработка случая, когда введено некорректное значение
            print("Некорректное значение ширины (размера шрифта)")
            # Восстанавливаем предыдущее корректное значение
            self.width_var.set(int(self.converter.default_font_size.pt))

    def select_files(self):
        """Выбор markdown файлов"""
        files = filedialog.askopenfilenames(
            title="Выберите файлы Markdown",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")]
        )

        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))

        self.status_label.config(text=f"Выбрано файлов: {len(self.selected_files)}")

    def remove_selected_files(self):
        """Удаление выбранных файлов из списка"""
        selected_indices = self.files_listbox.curselection()

        # Удаляем в обратном порядке, чтобы индексы не сбивались
        for index in reversed(selected_indices):
            self.files_listbox.delete(index)
            del self.selected_files[index]

        self.status_label.config(text=f"Выбрано файлов: {len(self.selected_files)}")

    def select_output_directory(self):
        """Выбор папки для сохранения"""
        directory = filedialog.askdirectory(title="Выберите папку для сохранения")

        if directory:
            self.output_directory = directory
            self.output_label.config(text=directory)

    def convert_files(self):
        """Конвертация выбранных файлов"""
        if not self.selected_files:
            messagebox.showwarning("Предупреждение", "Выберите файлы для конвертации")
            return

        if not self.output_directory:
            messagebox.showwarning("Предупреждение", "Выберите папку для сохранения")
            return

        # Сброс прогресса
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.selected_files)

        success_count = 0
        errors = []

        for i, input_file in enumerate(self.selected_files):
            # Обновляем статус
            filename = os.path.basename(input_file)
            self.status_label.config(text=f"Конвертация: {filename}")
            self.root.update()

            # Формируем имя выходного файла
            output_filename = Path(input_file).stem + ".docx"
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
                "Конвертация завершена с ошибками",
                f"Успешно: {success_count}\nОшибок: {len(errors)}\n\n{error_text}"
            )
        else:
            messagebox.showinfo(
                "Успех",
                f"Все файлы успешно конвертированы!\nВсего: {success_count}\nШрифт: {self.converter.default_font_name}, Размер: {self.converter.default_font_size.pt} pt"
            )

        self.status_label.config(text="Конвертация завершена")


def main():
    """Главная функция"""
    root = tk.Tk()
    app = ConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()