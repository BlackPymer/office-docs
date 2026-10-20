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


import os as _os
import sys as _sys

_sys.dont_write_bytecode = True  # не плодить __pycache__ рядом со скриптами скилла
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _common import ensure_runtime as _ensure_runtime  # noqa: E402

_ensure_runtime(())  # на Windows shebang не работает: берём Python из venv

DOC = """Конвертация документов между форматами (в т.ч. legacy .doc/.ppt/.xls).

Использование:
    office_convert.py <вход> <выход>

Формат определяется по расширению выхода:
  pdf, docx, doc, pptx, ppt, xlsx, xls, odt, ods, odp, rtf, html, txt, epub
  md  — только из docx/html/odt/rtf (через pandoc, точная структура)

Всё, что умеет LibreOffice, он и делает; legacy-форматы (doc/ppt/xls) он
напрямую перегоняет в современные (docx/pptx/xlsx).
"""
__doc__ = DOC


import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True  # не плодить __pycache__ рядом со скриптами скилла
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import find_pandoc, run_soffice  # noqa: E402

SOFFICE_EXT = {
    "pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls",
    "odt", "ods", "odp", "rtf", "html", "txt", "epub", "svg", "png",
}
PANDOC_SOURCES = {"docx", "odt", "rtf", "html", "htm", "doc", "epub"}


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    src = Path(sys.argv[1]).expanduser().resolve()
    dst = Path(sys.argv[2]).expanduser()
    if not src.is_file():
        print(f"Нет файла: {src}", file=sys.stderr)
        return 1
    dst.parent.mkdir(parents=True, exist_ok=True)
    fmt = dst.suffix.lower().lstrip(".")
    if not fmt:
        print("У выхода нет расширения — укажите его (напр. out.pdf)", file=sys.stderr)
        return 2

    if fmt == "md":
        if src.suffix.lower().lstrip(".") not in PANDOC_SOURCES:
            print(
                f"pandoc не читает {src.suffix}; для этого формата используйте "
                "LibreOffice-конвертацию в docx, затем в md",
                file=sys.stderr,
            )
            return 1
        pandoc = find_pandoc()
        if not pandoc:
            print("pandoc не найден (нужен только для конвертации в .md — см. README.md)",
                  file=sys.stderr)
            return 1
        r = subprocess.run(
            [pandoc, "-f", "docx", "-t", "gfm", str(src), "-o", str(dst)],
            capture_output=True,
            text=True,
        )
        if r.returncode:
            print(r.stderr, file=sys.stderr)
            return r.returncode
        print(dst)
        return 0

    if fmt not in SOFFICE_EXT:
        print(f"Неизвестный целевой формат: .{fmt}", file=sys.stderr)
        return 2

    with tempfile.TemporaryDirectory(prefix="office_convert_") as tmp:
        # HTML LibreOffice открывает как Writer/Web, который не умеет
        # экспортировать в OOXML — форсируем текстовый фильтр Writer'а
        extra: list[str] = []
        if src.suffix.lower() in (".html", ".htm") and fmt not in ("pdf", "html", "txt"):
            extra = ["--infilter=HTML (StarWriter)"]
        r = run_soffice(extra + ["--convert-to", fmt, "--outdir", tmp, str(src)])
        if r.returncode != 0:
            print(f"LibreOffice завершился с ошибкой:\n{r.stderr or r.stdout}",
                  file=sys.stderr)
            return 1
        produced = [p for p in Path(tmp).iterdir() if p.suffix.lower() == f".{fmt}"]
        if not produced:
            print(
                "Файл не создан. Возможно, для этого направления нужен другой "
                f"фильтр — см. references/legacy.md\n{r.stdout}{r.stderr}",
                file=sys.stderr,
            )
            return 1
        shutil.copyfile(produced[0], dst)
    print(dst)
    return 0


if __name__ == "__main__":
    sys.exit(main())
