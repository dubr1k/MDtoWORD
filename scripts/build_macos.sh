#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
BUILD_VENV="${BUILD_VENV:-$PROJECT_ROOT/.venv-macos-build}"
PYTHON="${PYTHON:-python3}"

cd "$PROJECT_ROOT"

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Сборка MDtoWORD.app поддерживается только на macOS."
    exit 1
fi

if [[ "$(uname -m)" != "arm64" ]]; then
    echo "Эта конфигурация сборки предназначена для Apple Silicon (arm64)."
    exit 1
fi

if [[ ! -x "$BUILD_VENV/bin/python" ]]; then
    "$PYTHON" -m venv "$BUILD_VENV"
fi

"$BUILD_VENV/bin/python" -m pip install --upgrade pip
"$BUILD_VENV/bin/python" -m pip install -r requirements.txt -r requirements-build.txt
"$BUILD_VENV/bin/python" -m PyInstaller --noconfirm --clean MDtoWORD.spec

codesign --force --deep --sign - "dist/MDtoWORD.app"
codesign --verify --deep --strict --verbose=2 "dist/MDtoWORD.app"

echo "Готово: $PROJECT_ROOT/dist/MDtoWORD.app"
