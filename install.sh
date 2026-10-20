#!/usr/bin/env bash
# Установка office-docs на Linux/macOS.
#
# Вся логика — в install.py (он же используется на Windows через install.ps1),
# здесь только поиск Python и передача аргументов:
#
#   ./install.sh                    # ядро (docx/pptx/xlsx/pdf)
#   ./install.sh --with-markitdown  # + markitdown (docx/pptx/pdf -> markdown)
#
# LibreOffice этим скриптом НЕ ставится: он нужен только для конвертации
# форматов и проверки рендера. См. раздел «LibreOffice» в README.md.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

for P in python3 python; do
  if command -v "$P" >/dev/null 2>&1; then
    exec "$P" "$ROOT/install.py" "$@"
  fi
done

echo "Python не найден. Нужен Python 3.10+." >&2
echo "  Debian/Ubuntu : sudo apt install python3 python3-venv" >&2
echo "  Fedora        : sudo dnf install python3" >&2
echo "  macOS         : brew install python" >&2
echo "  Windows       : py -3 install.py   (или .\\install.ps1)" >&2
exit 1
