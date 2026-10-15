#!/usr/bin/env bash
# Визуальная проверка документа: любой файл -> PDF -> JPG постранично.
#
# Использование:
#   render_check.sh <файл> [префикс] [dpi]
#
# Результат: картинки <префикс>-01.jpg, <префикс>-02.jpg, ... в текущем
# каталоге. Их нужно открыть через Read и посмотреть глазами — это обязательный
# шаг проверки после любых правок документа.
#
# dpi по умолчанию 110 (читаемый текст, не слишком тяжёлые картинки).
set -euo pipefail

find_soffice() {
  if [[ -n "${SOFFICE:-}" && -x "${SOFFICE:-}" ]]; then echo "$SOFFICE"; return 0; fi
  local HERE c
  HERE=$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)
  # prefix/ внутри репозитория (так ставит install.sh), потом старое
  # пользовательское место, потом PATH.
  for c in "$HERE"/prefix/opt/libreoffice*/program/soffice \
           "$HOME"/.local/share/office-skills/prefix/opt/libreoffice*/program/soffice; do
    [[ -x "$c" ]] && { echo "$c"; return 0; }
  done
  command -v soffice || command -v libreoffice || true
}

[[ $# -ge 1 ]] || { echo "Использование: render_check.sh <файл> [префикс] [dpi]" >&2; exit 2; }

SRC=$1
PREFIX=${2:-$(basename "${SRC%.*}")}
DPI=${3:-110}

[[ -f "$SRC" ]] || { echo "Нет файла: $SRC" >&2; exit 1; }

SOFFICE_BIN=$(find_soffice)
[[ -n "$SOFFICE_BIN" ]] || { echo "soffice не найден (см. SKILL.md)" >&2; exit 1; }

TMP=$(mktemp -d /tmp/office_render.XXXXXX)
trap 'rm -rf "$TMP"' EXIT

# 1) -> PDF (если исходник уже PDF, копируем как есть)
if [[ "${SRC,,}" == *.pdf ]]; then
  cp "$SRC" "$TMP/doc.pdf"
else
  SAL_USE_VCLPLUGIN=svp timeout 180 "$SOFFICE_BIN" \
    -env:UserInstallation="file://$TMP/profile" \
    --headless --norestore --nolockcheck \
    --convert-to pdf --outdir "$TMP" "$SRC" >/dev/null 2>&1 \
    || { echo "Конвертация в PDF не удалась: $SRC" >&2; exit 1; }
  base=$(basename "$SRC"); base=${base%.*}
  [[ -f "$TMP/$base.pdf" ]] || { echo "PDF не появился (фильтр не поддерживает формат?)" >&2; exit 1; }
  mv "$TMP/$base.pdf" "$TMP/doc.pdf"
fi

# 2) PDF -> JPG (имя с ведущим нулём по ширине числа страниц)
pdftoppm -jpeg -r "$DPI" "$TMP/doc.pdf" "$TMP/page" >/dev/null 2>&1

i=0
for img in "$TMP"/page-*.jpg; do
  [[ -e "$img" ]] || { echo "Страниц не найдено" >&2; exit 1; }
  i=$((i+1))
  n=$(printf '%02d' "$i")
  cp "$img" "${PREFIX}-${n}.jpg"
  echo "${PREFIX}-${n}.jpg"
done
echo "Страниц: $i (dpi=$DPI) — посмотрите их через Read"
