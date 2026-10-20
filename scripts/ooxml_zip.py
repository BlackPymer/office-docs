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
DOC = """Работа с ZIP-пакетом OOXML (docx/pptx/xlsx) без unzip и zip.

Использование:
    ooxml_zip.py list   <файл>                    # что внутри (замена unzip -l)
    ooxml_zip.py cat    <файл> <часть>            # часть в stdout (замена unzip -p)
    ooxml_zip.py unpack <файл> [каталог]          # распаковать
    ooxml_zip.py pack   <каталог> <файл>          # собрать архив заново

Зачем свой скрипт: на Windows нет `unzip`/`zip`, а ручная сборка архива
шеллом ломается на главной ловушке OOXML — `zip` только дописывает записи и
не удаляет старые, поэтому удалённая часть остаётся в архиве. Здесь архив
всегда собирается с нуля, а тип сжатия частей берётся из исходного файла
(манифест .ooxml_zip.json во время unpack), чтобы не раздувать уже сжатые
части (картинки, встроенные объекты).

Правьте XML только после unpack и никогда не pretty-print'ьте его.
После pack обязательно проверьте результат: validate.py out.docx --soffice.
"""
__doc__ = DOC


import json
import os
import sys
import zipfile
from pathlib import Path

sys.dont_write_bytecode = True  # не плодить __pycache__ рядом со скриптами скилла
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import ensure_runtime as _ensure_runtime  # noqa: E402

_ensure_runtime(())  # только стандартная библиотека, но пусть Windows найдёт venv

MANIFEST = ".ooxml_zip.json"
# Мусор, который не должен попадать в собранный пакет.
JUNK_NAMES = {"__MACOSX", ".DS_Store", "Thumbs.db", "desktop.ini", ".git"}
JUNK_SUFFIXES = (".pyc", ".orig", ".rej", ".swp")


def _is_junk(rel: Path) -> bool:
    if rel.name == MANIFEST:
        return True
    if any(part in JUNK_NAMES for part in rel.parts):
        return True
    return rel.name.endswith(JUNK_SUFFIXES)


def _require_file(path: Path) -> int | None:
    """Понятная ошибка вместо стектрейса, если файла нет или это не ZIP."""
    if not path.is_file():
        print(f"Нет файла: {path}", file=sys.stderr)
        return 1
    if not zipfile.is_zipfile(path):
        print(f"Это не ZIP-пакет: {path}", file=sys.stderr)
        return 1
    return None


def cmd_list(path: Path) -> int:
    if (rc := _require_file(path)) is not None:
        return rc
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            kind = "stored" if info.compress_type == zipfile.ZIP_STORED else "deflate"
            print(f"{info.file_size:>9}  {kind:<7}  {info.filename}")
    return 0


def cmd_cat(path: Path, part: str) -> int:
    if (rc := _require_file(path)) is not None:
        return rc
    part = part.replace("\\", "/")
    with zipfile.ZipFile(path) as z:
        try:
            data = z.read(part)
        except KeyError:
            names = [n for n in z.namelist() if part.strip("/") in n]
            hint = f"\nПохожие части: {', '.join(names[:10])}" if names else ""
            print(f"В архиве нет части {part}.{hint}", file=sys.stderr)
            return 1
    sys.stdout.write(data.decode("utf-8", errors="replace"))
    return 0


def cmd_unpack(path: Path, dest: Path | None) -> int:
    if (rc := _require_file(path)) is not None:
        return rc
    dest = dest or Path(path.stem + ".unpacked")
    with zipfile.ZipFile(path) as z:
        manifest = {
            info.filename: info.compress_type
            for info in z.infolist()
            if not info.is_dir()
        }
        z.extractall(dest)
    (dest / MANIFEST).write_text(json.dumps(manifest, indent=1), encoding="utf-8")
    print(dest)
    return 0


def cmd_pack(src_dir: Path, out: Path) -> int:
    if not src_dir.is_dir():
        print(f"Нет каталога: {src_dir}", file=sys.stderr)
        return 1
    files = [
        p for p in sorted(src_dir.rglob("*")) if p.is_file() and not _is_junk(p.relative_to(src_dir))
    ]
    if not any(p.relative_to(src_dir).as_posix() == "[Content_Types].xml" for p in files):
        print(
            f"В {src_dir} нет [Content_Types].xml — это не распакованный OOXML-пакет",
            file=sys.stderr,
        )
        return 1

    manifest_path = src_dir / MANIFEST
    manifest: dict[str, int] = {}
    if manifest_path.is_file():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except ValueError:
            manifest = {}

    def order(p: Path) -> tuple[int, str]:
        rel = p.relative_to(src_dir).as_posix()
        names = list(manifest)
        idx = names.index(rel) if rel in names else len(names)
        return (idx, rel)

    if out.exists():  # архив собираем с нуля, иначе останутся удалённые части
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(files, key=order):
            rel = p.relative_to(src_dir).as_posix()
            kind = manifest.get(rel, zipfile.ZIP_DEFLATED)
            info = zipfile.ZipInfo(rel)
            info.compress_type = kind
            info.external_attr = 0o600 << 16
            z.writestr(info, p.read_bytes())
    print(f"{out} ({out.stat().st_size} байт, частей: {len(files)})")
    return 0


def main() -> int:
    argv = sys.argv[1:]
    if not argv or argv[0].startswith("-"):
        print(__doc__)
        return 2
    cmd, rest = argv[0], argv[1:]
    if cmd == "list" and len(rest) == 1:
        return cmd_list(Path(rest[0]).expanduser())
    if cmd == "cat" and len(rest) == 2:
        return cmd_cat(Path(rest[0]).expanduser(), rest[1])
    if cmd == "unpack" and 1 <= len(rest) <= 2:
        return cmd_unpack(
            Path(rest[0]).expanduser(), Path(rest[1]).expanduser() if len(rest) == 2 else None
        )
    if cmd == "pack" and len(rest) == 2:
        return cmd_pack(Path(rest[0]).expanduser(), Path(rest[1]).expanduser())
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main())
