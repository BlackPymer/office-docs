#!/usr/bin/env bash
# Установка office-docs: создаёт виртуальное окружение и ставит зависимости.
#
#   ./install.sh                    # ядро (docx/pptx/xlsx/pdf)
#   ./install.sh --with-markitdown  # + markitdown (docx/pptx/pdf -> markdown, тянет ~300 МБ)
#
# LibreOffice НЕ ставится этим скриптом: он нужен только для конвертации форматов
# и для проверки рендера. См. раздел «LibreOffice» в README.md.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$ROOT/.venv"
MIN_MAJOR=10   # нужен python3 >= 3.10 (аннотации вида str | None)

WITH_MARKITDOWN=0
for arg in "$@"; do
  case "$arg" in
    --with-markitdown) WITH_MARKITDOWN=1 ;;
    -h|--help) sed -n '2,8p' "$0"; exit 0 ;;
    *) echo "Неизвестный аргумент: $arg (есть только --with-markitdown)" >&2; exit 2 ;;
  esac
done

echo "==> Проверяю Python"
if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 не найден. Установите Python 3.10+ (apt install python3 python3-venv)" >&2
  exit 1
fi
PYVER=$(python3 -c 'import sys; print(f"{sys.version_info[0]}.{sys.version_info[1]}")')
PYMAJ=${PYVER%%.*}; PYMIN=${PYVER##*.}
if (( PYMAJ < 3 || (PYMAJ == 3 && PYMIN < MIN_MAJOR) )); then
  echo "Нужен Python >= 3.10, а найден $PYVER. Обновите python3." >&2
  exit 1
fi
echo "    python3 $PYVER — ок"

echo "==> Создаю venv: $VENV"
python3 -m venv "$VENV"
VPY="$VENV/bin/python"

echo "==> Ставлю зависимости (ядро)"
"$VPY" -m pip install --upgrade pip -q
"$VPY" -m pip install -q -r "$ROOT/requirements.txt"
if (( WITH_MARKITDOWN )); then
  echo "==> Ставлю markitdown (тяжёлый пакет)"
  "$VPY" -m pip install -q "markitdown==0.1.8"
fi

echo "==> Проверяю импорты"
"$VPY" - <<'PY'
import importlib, sys
mods = ["docx", "pptx", "openpyxl", "pymupdf", "pypdf", "PIL", "lxml"]
try:
    import markitdown  # noqa: F401
    mods.append("markitdown")
except ImportError:
    pass
bad = []
for m in mods:
    try:
        importlib.import_module(m)
    except Exception as e:  # noqa: BLE001
        bad.append(f"{m}: {e}")
if bad:
    print("НЕ ИМПОРТИРУЮТСЯ:", *bad, sep="\n  ", file=sys.stderr)
    sys.exit(1)
print("    все библиотеки на месте:", ", ".join(mods))
PY

echo "==> Проверяю скрипты (компиляция)"
( cd "$ROOT/scripts" && "$VPY" -m py_compile ./*.py )

echo "==> Ищу LibreOffice (soffice)"
if SOFF="$("$VPY" -c "
import sys; sys.path.insert(0, '$ROOT/scripts')
from _common import find_soffice
print(find_soffice() or '')
")" && [[ -n "$SOFF" ]]; then
  echo "    найден: $SOFF"
else
  echo "    НЕ НАЙДЕН — конвертация форматов и проверка рендера работать не будут."
  echo "    Debian/Ubuntu : sudo apt install libreoffice"
  echo "    Fedora        : sudo dnf install libreoffice-core"
  echo "    macOS         : brew install --cask libreoffice"
  echo "    или задайте   : export SOFFICE=/путь/к/soffice"
fi

echo
echo "Готово. Скрипты запускаются напрямую, интерпретатор они находят сами:"
echo "  $ROOT/scripts/validate.py файл.docx --soffice"
echo "  $ROOT/scripts/render_check.sh файл.docx"
