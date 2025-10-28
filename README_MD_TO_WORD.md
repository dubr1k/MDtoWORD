# RU
• [🇺🇸 EN](#EN)
---
# Конвертер Markdown в Word

Приложение с графическим интерфейсом для конвертации файлов Markdown (.md) в формат Microsoft Word (.docx) с полным сохранением форматирования.

## Возможности

- **Графический интерфейс**: Удобный интерфейс на Tkinter
- **Пакетная обработка**: Выбор и конвертация нескольких файлов одновременно
- **Полная поддержка форматирования**:
  - Заголовки всех уровней (H1-H9)
  - **Жирный текст**
  - *Курсив*
  - ***Жирный курсив***
  - Маркированные списки
  - Нумерованные списки
  - Таблицы с границами
  - Блоки кода
  - Горизонтальные линии
- **Черный шрифт по умолчанию**: Times New Roman, 12pt
- **Прогресс-бар**: Отслеживание процесса конвертации

## Установка

1. Убедитесь, что у вас установлен Python 3.7 или выше

2. Установите зависимости:
```bash
pip install -r requirements.txt
```

или напрямую:
```bash
pip install python-docx
```

## Использование

1. Запустите скрипт:
```bash
python md_to_word_converter.py
```

2. В открывшемся окне:
   - Нажмите **"Выбрать файлы .md"** и выберите один или несколько markdown файлов
   - Нажмите **"Выбрать папку"** и укажите, куда сохранить результаты
   - Нажмите **"Конвертировать"**

3. Готовые файлы .docx будут сохранены в выбранной папке с теми же именами

## Поддерживаемый синтаксис Markdown

### Заголовки
```markdown
# Заголовок 1
## Заголовок 2
### Заголовок 3
```

### Форматирование текста
```markdown
**жирный текст**
*курсив*
***жирный курсив***
`код`
```

### Списки
```markdown
- Элемент маркированного списка
- Еще один элемент

1. Нумерованный список
2. Второй элемент
```

### Таблицы
```markdown
| Заголовок 1 | Заголовок 2 |
|-------------|-------------|
| Ячейка 1    | Ячейка 2    |
```

### Блоки кода
```markdown
```
код на нескольких
строках
```
```

### Горизонтальные линии
```markdown
---
```

## Структура проекта

- `md_to_word_converter.py` - Основной скрипт с GUI и логикой конвертации
- `requirements.txt` - Зависимости проекта
- `README_MD_TO_WORD.md` - Документация

## Примечания

- Скрипт автоматически создает таблицы с границами
- Все текст форматируется черным цветом (RGB: 0, 0, 0)
- Код отображается шрифтом Courier New
- Поддерживается вложенное форматирование в таблицах

## Требования

- Python 3.7+
- python-docx 1.1.2+
- tkinter (входит в стандартную поставку Python)

## Лицензия

Свободное использование

## Автор

Создано для конвертации аналитических отчетов и документации
---


# EN
• [🇷🇺 RU](#RU)
---
Markdown file converter:

```markdown
# Markdown to Word Converter

A graphical application for converting Markdown (`.md`) files into Microsoft Word (`.docx`) format while fully preserving formatting.

## Features

- **Graphical User Interface**: Built with Tkinter for ease of use  
- **Batch Processing**: Convert multiple files at once  
- **Full Formatting Support**:
  - Headings (H1–H9)
  - Bold text
  - Italic text
  - Bold italic text
  - Bulleted lists
  - Numbered lists
  - Tables with borders
  - Code blocks
  - Horizontal rules
- **Default Styling**: Black text in Times New Roman, 12pt  
- **Progress Bar**: Real-time conversion progress tracking  

## Installation

1. Ensure you have Python 3.7 or higher installed.  
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

   Or install directly:

   ```bash
   pip install python-docx
   ```

## Usage

1. Run the script:

   ```bash
   python md_to_word_converter.py
   ```

2. In the opened window:
   - Click **"Select .md Files"** and choose one or more Markdown files  
   - Click **"Select Output Folder"** to specify where to save the results  
   - Click **"Convert"**

The resulting `.docx` files will be saved in the selected folder with the same base filenames.

## Supported Markdown Syntax

### Headings

```markdown
# Heading 1
## Heading 2
### Heading 3
```

### Text Formatting

```markdown
**bold text**  
*italic text*  
***bold italic text***  
`inline code`
```

### Lists

```markdown
- Bullet list item
- Another item

1. Numbered list item
2. Second item
```

### Tables

```markdown
| Header 1    | Header 2    |
|-------------|-------------|
| Cell 1      | Cell 2      |
```

### Code Blocks

```
multi-line
code block
```

### Horizontal Rules

```markdown
---
```

## Project Structure

- `md_to_word_converter.py` — Main script with GUI and conversion logic  
- `requirements.txt` — Project dependencies  
- `README_MD_TO_WORD.md` — Documentation  

## Notes

- Tables are automatically created with borders  
- All text uses black color (RGB: 0, 0, 0)  
- Code is displayed in Courier New font  
- Nested formatting within tables is supported  

## Requirements

- Python 3.7+  
- python-docx 1.1.2+  
- tkinter (included in standard Python distribution)  

## License

Free to use.

## Author

Created for converting analytical reports and documentation.
```
---
