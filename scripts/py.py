#!/bin/sh
"true" '''\'
# Python-скрипт: sh-трampoline ниже запускает его нужным интерпретатором.
# Порядок: .venv в корне репозитория -> .venv рядом со скриптом ->
# старый пользовательский venv -> python3.
D="$(dirname "$0")"
for P in "$D/../.venv/bin/python" "$D/.venv/bin/python" \
         "$D/../.venv/Scripts/python.exe" "$D/.venv/Scripts/python.exe" \
         "$HOME/.local/share/office-skills/venv/bin/python" \
         "$HOME/.local/share/office-skills/venv/Scripts/python.exe"; do
  [ -x "$P" ] && exec "$P" "$0" "$@"
done
exec python3 "$0" "$@"
''' + ''   # конкатенация -> Python не примет трampoline за docstring модуля
DOC = """Запуск своего кода интерпретатором из venv скилла.

Использование:
    py.py -c "import docx; print(docx.Document('a.docx').paragraphs[0].text)"
    py.py snippet.py аргументы...

Зачем: системный `python` / `py -3` может не иметь docx/pptx/openpyxl/pymupdf, а
этот скрипт сам находит Python из venv скилла (его ставит install.py) и
передаёт ему аргументы как есть. Так ад-хок сниппеты работают одинаково на
Windows и POSIX, без знания путей до venv.

    SK=~/.config/opencode/skills/office-docs      # PowerShell: $SK = "..."
    $SK/scripts/py.py -c "import pymupdf; print(pymupdf.open('doc.pdf').page_count)"
"""
__doc__ = DOC


import os as _os
import sys as _sys

_sys.dont_write_bytecode = True  # не плодить __pycache__ рядом со скриптами скилла
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _common import ensure_runtime as _ensure_runtime  # noqa: E402


import os
import subprocess
import sys

from _common import find_runtime_python, no_window_kwargs  # noqa: E402

_ensure_runtime(())  # самому этому скрипту библиотеки не нужны


def main() -> int:
    argv = sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 2 if not argv else 0
    python = find_runtime_python()
    if not python:
        print(
            "Не нашёл интерпретатор с библиотеками скилла (docx/pptx/openpyxl/pymupdf).\n"
            "Запустите установку: py -3 install.py (Windows) или ./install.sh (Linux/macOS).",
            file=sys.stderr,
        )
        return 1
    return subprocess.run([python] + argv, **no_window_kwargs()).returncode


if __name__ == "__main__":
    sys.exit(main())
