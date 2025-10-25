# RU
# 📄 MDtoDOCX - Конвертер Markdown в Word

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**Профессиональный конвертер Markdown файлов в Microsoft Word с графическим интерфейсом**

[🚀 Быстрый старт](#-быстрый-старт) • [📋 Возможности](#-возможности) • [💻 Установка](#-установка) • [🎯 Использование](#-использование) • [🇺🇸 EN](#EN)

</div>

---

## 🎯 Возможности

### ✨ Основные функции
- **🖥️ Графический интерфейс** - Удобный GUI на Tkinter
- **📦 Пакетная обработка** - Конвертация нескольких файлов одновременно
- **📊 Прогресс-бар** - Отслеживание процесса конвертации
- **🎨 Полное форматирование** - Сохранение всех элементов Markdown

### 📝 Поддерживаемые элементы

| Элемент | Markdown | Результат в Word |
|---------|----------|------------------|
| **Заголовки** | `# ## ###` | Стили Heading 1-9 |
| **Жирный текст** | `**текст**` | Bold formatting |
| **Курсив** | `*текст*` | Italic formatting |
| **Жирный курсив** | `***текст***` | Bold + Italic |
| **Код** | `` `код` `` | Courier New, 10pt |
| **Таблицы** | `\| \|` | Таблицы с границами |
| **Списки** | `- 1.` | Bullet/Numbered lists |
| **Горизонтальные линии** | `---` | Разделители |

---

## 🚀 Быстрый старт

### Установка
```bash
# 1. Клонируйте репозиторий
git clone https://github.com/dubr1k/MDtoWORD.git
cd MDtoWORD

# 2. Установите зависимости
pip install -r requirements.txt

# 3. Запустите программу
python md_to_word_converter.py
```

### Windows (Быстрый запуск)
Для Windows пользователей доступен батник `запуск_конвертера.bat` (только локально):
1. Дважды кликните по `запуск_конвертера.bat`
2. Программа автоматически установит зависимости и запустится

---

## 💻 Установка

### Требования
- **Python 3.7+** 
- **python-docx 1.1.2+**
- **tkinter** (входит в стандартную поставку Python)

### Установка зависимостей
```bash
# Автоматическая установка
pip install -r requirements.txt

# Или вручную
pip install python-docx
```

---

## 🎯 Использование

### Пошаговая инструкция

1. **Запустите программу**
   ```bash
   python md_to_word_converter.py
   ```

2. **Выберите файлы**
   - Нажмите **"Выбрать файлы .md"**
   - Выберите один или несколько Markdown файлов

3. **Выберите папку сохранения**
   - Нажмите **"Выбрать папку"**
   - Укажите директорию для сохранения .docx файлов

4. **Конвертируйте**
   - Нажмите **"Конвертировать"**
   - Следите за прогрессом в статус-баре

### Примеры использования

#### Конвертация одного файла
```
Выберите: report.md
Результат: report.docx
```

#### Пакетная конвертация
```
Выберите: file1.md, file2.md, file3.md
Результат: file1.docx, file2.docx, file3.docx
```

---

## 📋 Примеры Markdown

### Заголовки
```markdown
# Главный заголовок
## Подзаголовок
### Мелкий заголовок
```

### Форматирование текста
```markdown
**Жирный текст**
*Курсив*
***Жирный курсив***
`Код в строке`
```

### Таблицы
```markdown
| Заголовок 1 | Заголовок 2 | Заголовок 3 |
|-------------|-------------|-------------|
| Ячейка 1    | Ячейка 2    | Ячейка 3    |
| Данные 1    | Данные 2    | Данные 3    |
```

### Списки
```markdown
- Маркированный список
- Второй элемент
- Третий элемент

1. Нумерованный список
2. Второй элемент
3. Третий элемент
```

### Блоки кода
```python
def hello_world():
    print("Hello, World!")
    return "Успех"
```

---

## 🏗️ Структура проекта

```
MDtoWORD/
├── 📄 md_to_word_converter.py    # Основной скрипт с GUI
├── 📋 requirements.txt           # Зависимости проекта
├── 📖 README.md                  # Документация (этот файл)
├── 📄 ИНСТРУКЦИЯ.txt             # Краткая инструкция
├── 🚀 запуск_конвертера.bat      # Быстрый запуск для Windows (локально)
└── 📝 test_example.md            # Пример Markdown файла (локально)
```

---

## ⚙️ Технические детали

### Форматирование по умолчанию
- **Шрифт**: Times New Roman, 12pt
- **Цвет**: Черный (RGB: 0, 0, 0)
- **Код**: Courier New, 10pt
- **Таблицы**: С границами и форматированием

### Поддерживаемые форматы
- **Входные**: `.md` (Markdown)
- **Выходные**: `.docx` (Microsoft Word)

---

## 🔧 Решение проблем

### Частые вопросы

**Q: Программа не запускается**
```bash
# Проверьте версию Python
python --version

# Установите зависимости
pip install python-docx
```

**Q: Не сохраняется форматирование таблиц**
- Убедитесь, что таблица имеет правильный синтаксис Markdown
- Проверьте, что в таблице есть разделительная строка `|---|---|`

**Q: Ошибка кодировки**
- Убедитесь, что файлы .md сохранены в UTF-8

### Логи и отладка
Программа выводит подробные сообщения об ошибках в интерфейсе.
---


# EN

# 📄 MDtoDOCX - Markdown to Word Converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**A professional Markdown to Microsoft Word converter with a graphical interface**

[🚀 Quick Start](#-quick-start) • [📋 Features](#-features) • [💻 Installation](#-installation) • [🎯 Usage](#-usage) • [🇷🇺 RU](#RU)

</div>

---

## 🎯 Features

### ✨ Main Functions
- **🖥️ Graphical Interface** – Convenient GUI built with Tkinter
- **📦 Batch Processing** – Convert multiple files at once
- **📊 Progress Bar** – Track conversion progress
- **🎨 Full Formatting** – Preserves all Markdown elements

### 📝 Supported Elements

| Element | Markdown | Result in Word |
|----------|-----------|----------------|
| **Headings** | `# ## ###` | Heading 1–9 styles |
| **Bold text** | `**text**` | Bold formatting |
| **Italic text** | `*text*` | Italic formatting |
| **Bold + Italic** | `***text***` | Bold + Italic |
| **Inline code** | `` `code` `` | Courier New, 10pt |
| **Tables** | `\| \|` | Tables with borders |
| **Lists** | `- 1.` | Bulleted/Numbered lists |
| **Horizontal lines** | `---` | Section dividers |

---

## 🚀 Quick Start

### Installation
```bash
# 1. Clone the repository
git clone https://github.com/dubr1k/MDtoWORD.git
cd MDtoWORD

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the program
python md_to_word_converter.py
```

### Windows (Quick Launch)
For Windows users, there is a batch file `запуск_конвертера.bat` (local use only):
1. Double-click `запуск_конвертера.bat`
2. The program will automatically install dependencies and start

---

## 💻 Installation

### Requirements
- **Python 3.7+**
- **python-docx 1.1.2+**
- **tkinter** (included with Python)

### Install Dependencies
```bash
# Automatic installation
pip install -r requirements.txt

# Or manually
pip install python-docx
```

---

## 🎯 Usage

### Step-by-step guide

1. **Run the program**
   ```bash
   python md_to_word_converter.py
   ```

2. **Select files**
   - Click **“Select .md files”**
   - Choose one or more Markdown files

3. **Select output folder**
   - Click **“Select folder”**
   - Choose a directory to save the `.docx` files

4. **Convert**
   - Click **“Convert”**
   - Monitor progress in the status bar

### Usage Examples

#### Converting a single file
```
Input: report.md
Output: report.docx
```

#### Batch conversion
```
Input: file1.md, file2.md, file3.md
Output: file1.docx, file2.docx, file3.docx
```

---

## 📋 Markdown Examples

### Headings
```markdown
# Main heading
## Subheading
### Smaller heading
```

### Text Formatting
```markdown
**Bold text**
*Italic text*
***Bold and Italic***
`Inline code`
```

### Tables
```markdown
| Header 1 | Header 2 | Header 3 |
|-----------|-----------|-----------|
| Cell 1    | Cell 2    | Cell 3    |
| Data 1    | Data 2    | Data 3    |
```

### Lists
```markdown
- Bullet list item
- Second item
- Third item

1. Numbered list
2. Second item
3. Third item
```

### Code Blocks
```python
def hello_world():
    print("Hello, World!")
    return "Success"
```

---

## 🏗️ Project Structure

```
MDtoWORD/
├── 📄 md_to_word_converter.py    # Main GUI script
├── 📋 requirements.txt           # Project dependencies
├── 📖 README.md                  # Documentation (this file)
├── 📄 INSTRUCTION.txt            # Quick guide
├── 🚀 запуск_конвертера.bat      # Quick launcher for Windows (local)
└── 📝 test_example.md            # Markdown example file (local)
```

---

## ⚙️ Technical Details

### Default Formatting
- **Font**: Times New Roman, 12pt
- **Color**: Black (RGB: 0, 0, 0)
- **Code**: Courier New, 10pt
- **Tables**: With borders and formatting

### Supported Formats
- **Input**: `.md` (Markdown)
- **Output**: `.docx` (Microsoft Word)

---

## 🔧 Troubleshooting

### Common Issues

**Q: The program doesn’t start**
```bash
# Check Python version
python --version

# Install dependencies
pip install python-docx
```

**Q: Table formatting is lost**
- Ensure the Markdown table syntax is correct
- Verify that the table includes the separator line `|---|---|`

**Q: Encoding error**
- Make sure `.md` files are saved in UTF-8

### Logs and Debugging
The program displays detailed error messages directly in the interface.
