@echo off
chcp 65001 >nul
echo ========================================
echo   Конвертер Markdown в Word
echo ========================================
echo.

REM Проверка установки Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ОШИБКА: Python не установлен или не добавлен в PATH
    echo Установите Python с https://www.python.org/
    pause
    exit /b 1
)

echo Проверка зависимостей...
pip show python-docx >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Устанавливаю необходимые библиотеки...
    pip install python-docx
    if %errorlevel% neq 0 (
        echo.
        echo ОШИБКА: Не удалось установить зависимости
        pause
        exit /b 1
    )
)

echo.
echo Запуск конвертера...
echo.
python md_to_word_converter.py

if %errorlevel% neq 0 (
    echo.
    echo ОШИБКА: Не удалось запустить конвертер
    pause
)

