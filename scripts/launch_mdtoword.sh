#!/usr/bin/env bash
# Запуск MDtoWORD (скрипт лежит в scripts/, проект — родительская папка)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT" || exit 1

# Предпочтительно Python из conda-окружения mdtoword, иначе системный
if [ -x "$HOME/.conda/envs/mdtoword/bin/python" ]; then
    exec "$HOME/.conda/envs/mdtoword/bin/python" "$PROJECT_ROOT/md_to_word_converter.py"
else
    exec python3 "$PROJECT_ROOT/md_to_word_converter.py"
fi
