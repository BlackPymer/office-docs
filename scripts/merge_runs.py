#!/bin/sh
"true" '''\'
# Python-скрипт: sh-трampoline ниже запускает его нужным интерпретатором.
# Порядок: .venv в корне репозитория -> .venv рядом со скриптом ->
# старый пользовательский venv -> python3.
D="$(dirname "$0")"
for P in "$D/../.venv/bin/python" "$D/.venv/bin/python" \
         "$HOME/.local/share/office-skills/venv/bin/python"; do
  [ -x "$P" ] && exec "$P" "$0" "$@"
done
exec python3 "$0" "$@"
''' + ''   # конкатенация -> Python не примет трampoline за docstring модуля
DOC = """Склейка фрагментированных text-run'ов в .docx — чтобы текст был находим.

Word разбивает один видимый абзац на много <w:r> (проверка орфографии,
revision id, вставки). Фраза «Привет, мир» часто не существует в
word/document.xml как непрерывная строка, и find/replace молча не находит её.

Скрипт объединяет соседние runs с одинаковым оформлением (идентентичным rPr)
внутри одного абзаца, ничего не меняя визуально.

Использование:
    merge_runs.py <файл.docx> [-o выход.docx]

Без -o перезаписывает файл на месте (сделайте бэкап сами: cp a.docx a.bak.docx).
"""
__doc__ = DOC


import copy
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOC_XML = "word/document.xml"


def rpr_key(run) -> bytes:
    """Сериализация rPr (или None) — по ней решаем, можно ли склеивать."""
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return b""
    return etree.tostring(rpr)


def merge_document_xml(data: bytes) -> tuple[bytes, int]:
    root = etree.fromstring(data)
    merged = 0
    # Склеиваем только прямые runs-дети абзаца (не лезем внутрь w:ins/w:del —
    # там свои правила ревизий)
    for p in root.iter(f"{W}p"):
        changed = True
        while changed:
            changed = False
            children = list(p)
            for a, b in zip(children, children[1:]):
                if a.tag != f"{W}r" or b.tag != f"{W}r":
                    continue
                if rpr_key(a) != rpr_key(b):
                    continue
                # не склеиваем, если у runs есть атрибуты-различия (например w:rsidR)
                if a.attrib or b.attrib:
                    continue
                for child in list(b):
                    a.append(child)
                p.remove(b)
                merged += 1
                changed = True
                break
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), merged


def main() -> int:
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        return 2
    src = Path(argv[0]).expanduser()
    out = None
    if "-o" in argv:
        out = Path(argv[argv.index("-o") + 1])
    if not src.is_file():
        print(f"Нет файла: {src}", file=sys.stderr)
        return 1

    with zipfile.ZipFile(src) as z:
        if DOC_XML not in z.namelist():
            print("Не docx (нет word/document.xml)", file=sys.stderr)
            return 1
        new_xml, merged = merge_document_xml(z.read(DOC_XML))

        dst = out or src
        tmpfd, tmppath = tempfile.mkstemp(suffix=".docx", dir=str(dst.parent))
        os_close = __import__("os").close
        os_close(tmpfd)
        try:
            with zipfile.ZipFile(src) as zin, zipfile.ZipFile(
                tmppath, "w", zipfile.ZIP_DEFLATED
            ) as zout:
                for item in zin.infolist():
                    data = new_xml if item.filename == DOC_XML else zin.read(item.filename)
                    # гиперссылки/сервисные части не сжимаем — как в оригинале
                    if item.filename.endswith(".rels") or item.filename == DOC_XML:
                        zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                        zi.compress_type = zipfile.ZIP_DEFLATED
                        zi.external_attr = item.external_attr
                        zout.writestr(zi, data)
                    else:
                        zout.writestr(item, data)
            shutil.move(tmppath, dst)
        finally:
            if os.path.exists(tmppath):
                os.unlink(tmppath)

    print(f"Склеено runs: {merged} -> {dst}")
    if merged:
        print("Теперь фразу можно искать/заменять обычным grep+edit по document.xml.")
    return 0


import os  # noqa: E402  (используется в finally)

if __name__ == "__main__":
    sys.exit(main())
