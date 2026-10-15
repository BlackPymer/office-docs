#!/bin/sh
"true" '''\'
# Python-скрипт: sh-трampoline ниже запускает его нужным интерпретатором.
# Порядок: .venv рядом со скриптом -> старый пользовательский venv -> python3.
D="$(dirname "$0")"
for P in "$D/.venv/bin/python" "$HOME/.local/share/office-skills/venv/bin/python"; do
  [ -x "$P" ] && exec "$P" "$0" "$@"
done
exec python3 "$0" "$@"
''' + ''   # конкатенация -> Python не примет трampoline за docstring модуля
"""Структурная валидация OOXML-файлов (docx/pptx/xlsx) перед выдачей результата.

Использование:
    validate.py <файл> [--soffice] [--original <исходный>]

Проверяет:
  1. ZIP целостность.
  2. Все XML внутри — синтаксически корректны (lxml).
  3. Есть [Content_Types].xml.
  4. Все относительные пути из *_rels/*.rels существуют в архиве
     (битые ссылки = Word откроет с «восстановлением»).
  5. Файл открывается «родной» библиотекой (python-docx / python-pptx / openpyxl).
  6. --soffice: LibreOffice реально открывает и конвертирует файл
     (сильнейший практический тест).

Exit code 0 = все проверки пройдены.
"""


import os
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlparse

sys.dont_write_bytecode = True  # не плодить __pycache__ рядом со скриптами скилла
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import convert_with_soffice  # noqa: E402

from lxml import etree  # noqa: E402

FAIL = 0


def fail(msg: str) -> None:
    global FAIL
    FAIL = 1
    print(f"FAIL  {msg}")


def ok(msg: str) -> None:
    print(f"ok    {msg}")


def check_zip(z: zipfile.ZipFile) -> None:
    bad = z.testzip()
    if bad:
        fail(f"повреждённый элемент архива: {bad}")
    else:
        ok("zip целостность")
    names = set(z.namelist())
    if "[Content_Types].xml" not in names:
        fail("нет [Content_Types].xml")
    else:
        ok("[Content_Types].xml есть")
    return_names(names)


NAMES: set[str] = set()


def return_names(names: set[str]) -> None:
    global NAMES
    NAMES = names


def check_xml(z: zipfile.ZipFile) -> None:
    bad = 0
    for n in sorted(NAMES):
        if n.endswith((".xml", ".rels")):
            try:
                etree.fromstring(z.read(n))
            except Exception as e:  # noqa: BLE001
                fail(f"битый XML {n}: {e}")
                bad += 1
    if not bad:
        ok(f"все XML разбираются ({sum(1 for n in NAMES if n.endswith(('.xml', '.rels')))} шт.)")


def check_rels(z: zipfile.ZipFile) -> None:
    missing = 0
    total = 0
    for n in sorted(NAMES):
        if not n.endswith(".rels"):
            continue
        base = os.path.dirname(os.path.dirname(n))  # _rels -> папка источника
        try:
            root = etree.fromstring(z.read(n))
        except Exception:  # noqa: BLE001
            continue
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            target = unquote(rel.get("Target") or "")
            if urlparse(target).scheme or target.startswith("/"):
                continue  # абсолютные/внешние не проверяем
            resolved = os.path.normpath(os.path.join(base, target)).replace(os.sep, "/")
            total += 1
            if resolved not in NAMES:
                fail(f"битая ссылка в {n}: {target} -> {resolved} не найден")
                missing += 1
    if not missing:
        ok(f"ссылки в .rels корректны ({total} шт.)")


def check_native(path: Path) -> None:
    ext = path.suffix.lower()
    try:
        if ext == ".docx":
            import docx

            d = docx.Document(str(path))
            ok(f"python-docx открыл ({len(d.paragraphs)} абз., {len(d.tables)} табл.)")
        elif ext == ".pptx":
            from pptx import Presentation

            p = Presentation(str(path))
            ok(f"python-pptx открыл ({len(p.slides)} слайдов)")
        elif ext in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook

            wb = load_workbook(str(path))
            ok(f"openpyxl открыл (листы: {', '.join(wb.sheetnames)})")
            wb.close()
        else:
            ok(f"для {ext} родная библиотека не проверялась")
    except Exception as e:  # noqa: BLE001
        fail(f"родная библиотека не открыла файл: {e}")


def check_soffice(path: Path) -> None:
    import tempfile

    with tempfile.TemporaryDirectory(prefix="validate_lo_") as tmp:
        try:
            r = convert_with_soffice(str(path), tmp, "pdf", timeout=180)
        except Exception as e:  # noqa: BLE001
            fail(f"LibreOffice недоступен: {e}")
            return
        out = Path(tmp) / (path.stem + ".pdf")
        if r.returncode == 0 and out.exists() and out.stat().st_size > 0:
            ok(f"LibreOffice конвертировал в PDF ({out.stat().st_size} байт)")
        else:
            fail(f"LibreOffice не смог открыть/конвертировать: {r.stderr or r.stdout}")


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if not args:
        print(__doc__)
        return 2
    path = Path(args[0]).expanduser()
    if not path.is_file():
        print(f"Нет файла: {path}", file=sys.stderr)
        return 1

    print(f"== {path.name} ==")
    try:
        with zipfile.ZipFile(path) as z:
            check_zip(z)
            check_xml(z)
            check_rels(z)
    except zipfile.BadZipFile:
        fail("это не ZIP/OOXML файл")
    check_native(path)
    if "--soffice" in flags:
        check_soffice(path)

    print("ИТОГ:", "FAIL" if FAIL else "PASS")
    return FAIL


if __name__ == "__main__":
    sys.exit(main())
