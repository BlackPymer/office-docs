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
DOC = """Визуальная проверка документа: любой файл -> PDF -> JPG постранично.

Использование:
    render_check.py <файл> [префикс] [dpi]

Результат: картинки <префикс>-01.jpg, <префикс>-02.jpg, ... в текущем каталоге.
Их нужно открыть через Read и посмотреть глазами — это обязательный шаг
проверки после любых правок документа.

dpi по умолчанию 110 (читаемый текст, не слишком тяжёлые картинки).

PDF делается LibreOffice-ом (или берётся как есть, если на входе PDF), страницы
рисует PyMuPDF — poppler/pdftoppm не нужен. Если pdftoppm всё же есть в PATH,
он используется как запасной вариант. Работает одинаково на Windows и POSIX.
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
from _common import (  # noqa: E402
    convert_with_soffice,
    ensure_runtime as _ensure_runtime,
    find_pdftoppm,
    no_window_kwargs,
)

# PyMuPDF нужен только если нет pdftoppm; просить его раньше времени незачем.
_ensure_runtime(() if find_pdftoppm() else ("pymupdf",))


def to_pdf(src: Path, tmp: Path) -> Path:
    """Сделать PDF из любого офисного файла (PDF возвращается как есть)."""
    if src.suffix.lower() == ".pdf":
        dst = tmp / "doc.pdf"
        shutil.copyfile(src, dst)
        return dst

    try:
        r = convert_with_soffice(str(src), str(tmp), "pdf", timeout=180)
    except Exception as e:  # noqa: BLE001 — LibreOffice отсутствует или повис
        raise SystemExit(f"Конвертация в PDF не удалась: {e}")
    produced = sorted(tmp.glob("*.pdf"))
    if r.returncode != 0 and not produced:
        raise SystemExit(
            f"Конвертация в PDF не удалась: {src}\n{(r.stderr or r.stdout or '').strip()}"
        )
    if not produced:
        raise SystemExit(
            f"PDF не появился (фильтр не поддерживает формат?): {src}\n"
            f"{(r.stdout or '').strip()}"
        )
    # LibreOffice мог выдать имя с другим окончанием — предпочитаем совпадение по stem
    for p in produced:
        if p.stem == src.stem:
            return p
    return produced[0]


def rasterize_pymupdf(pdf: Path, prefix: str, dpi: int) -> int:
    try:
        import pymupdf as fitz
    except ImportError:  # старые сборки знают модуль только как fitz
        import fitz  # type: ignore[no-redef]

    doc = fitz.open(str(pdf))
    try:
        total = doc.page_count
        for i in range(total):
            pix = doc.load_page(i).get_pixmap(dpi=dpi)
            out = Path(f"{prefix}-{i + 1:02d}.jpg")
            out.write_bytes(pix.tobytes("jpeg", jpg_quality=85))
            print(out.name)
    finally:
        doc.close()
    return total


def rasterize_poppler(pdftoppm: str, pdf: Path, prefix: str, dpi: int) -> int:
    with tempfile.TemporaryDirectory(prefix="render_check_") as tmp:
        target = os.path.join(tmp, "page")
        r = subprocess.run(
            [pdftoppm, "-jpeg", "-r", str(dpi), str(pdf), target],
            capture_output=True,
            text=True,
            timeout=300,
            **no_window_kwargs(),
        )
        pages = sorted(Path(tmp).glob("page-*.jpg"))
        if r.returncode != 0 or not pages:
            raise SystemExit(f"pdftoppm не смог разобрать PDF: {r.stderr.strip()}")
        for i, img in enumerate(pages, 1):
            out = Path(f"{prefix}-{i:02d}.jpg")
            shutil.copyfile(img, out)
            print(out.name)
        return len(pages)


def main() -> int:
    argv = sys.argv[1:]
    if not argv or argv[0].startswith("-"):
        print(__doc__)
        return 2
    src = Path(argv[0]).expanduser()
    if not src.is_file():
        print(f"Нет файла: {src}", file=sys.stderr)
        return 1
    prefix = argv[1] if len(argv) > 1 else src.stem
    try:
        dpi = int(argv[2]) if len(argv) > 2 else 110
    except ValueError:
        print(f"dpi должен быть числом, а не {argv[2]!r}", file=sys.stderr)
        return 2

    pdftoppm = find_pdftoppm()
    has_pymupdf = True
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        try:
            import fitz  # noqa: F401
        except ImportError:
            has_pymupdf = False
    if not pdftoppm and not has_pymupdf:
        print(
            "Нет ни PyMuPDF, ни pdftoppm — нечем нарисовать страницы. "
            "Поставьте зависимости скилла (install.py / install.sh).",
            file=sys.stderr,
        )
        return 1

    with tempfile.TemporaryDirectory(prefix="office_render_") as tmp:
        pdf = to_pdf(src.resolve(), Path(tmp))
        total = (
            rasterize_pymupdf(pdf, prefix, dpi)
            if has_pymupdf
            else rasterize_poppler(pdftoppm or "", pdf, prefix, dpi)
        )

    if not total:
        print("Страниц не найдено", file=sys.stderr)
        return 1
    print(f"Страниц: {total} (dpi={dpi}) — посмотрите их через Read")
    return 0


if __name__ == "__main__":
    sys.exit(main())
