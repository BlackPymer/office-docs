# office-docs

Скилл для агентов: чтение, правка, создание и проверка документов —
Word (.docx/.dotx/.doc), PowerPoint (.pptx/.potx/.ppt), Excel (.xlsx/.xlsm/.xls),
PDF (.pdf), плюс конвертация между ними.

Это не библиотека-обёртка, а набор скриптов с рабочей инструкцией (`SKILL.md`),
рассчитанной на то, что их выполняет AI-агент. Скрипты маленькие и читаются
целиком — правьте под себя.

## Что внутри

| Файл | Что делает |
|---|---|
| `SKILL.md` | Инструкция для агента: какой подход выбрать, какие грабли есть |
| `scripts/validate.py` | Проверка файла: ZIP/XML/ссылки/открываемость. Печатает `ИТОГ: PASS/FAIL` |
| `scripts/render_check.sh` | Файл → PDF → JPG постранично (визуальная проверка вёрстки) |
| `scripts/office_convert.py` | Конвертация: всё → .pdf, docx→odt/html/txt/md, pptx→odp/html, xlsx→ods/html, legacy (.doc/.ppt/.xls) → современные. pdf→docx не работает — см. `references/legacy.md` |
| `scripts/merge_runs.py` | Склейка фрагментированных runs в docx (нужно перед поиском/заменой текста) |
| `scripts/pptx_outline.py` | Текстовая выгрузка всех слайдов и фигур презентации |
| `scripts/pptx_dup_slide.py` | Дублирование слайда с правильным bookkeeping'ом |
| `scripts/xlsx_dump.py` | Содержимое книги: значения или формулы, ошибки помечает `!!` |
| `scripts/xlsx_recalc.py` | Пересчёт формул через LibreOffice + JSON-отчёт об ошибках |
| `scripts/pdf_tool.py` | PDF: info, text, формы, fill, merge, split, rotate, replace |
| `scripts/_common.py` | Поиск LibreOffice и запуск headless с временным профилем |
| `tests/smoke.py` | Самодостаточный smoke-тест: генерирует образцы и прогоняет все скрипты |
| `references/*.md` | Подробности по каждому формату (читать по мере надобности) |

## Установка

Нужен Python **3.10 или новее**.

```bash
git clone <url>
cd office-docs
./install.sh
```

Скрипт создаёт `.venv/` **внутри папки скилла** и ставит зависимости
(версии зафиксированы в `requirements.txt`).

Дополнительно, если нужен `markitdown` (docx/pptx/pdf → markdown):

```bash
./install.sh --with-markitdown
```

Это примерно +300 МБ (`onnxruntime`, `magika`) — ставьте только если реально
пользуетесь.

### Установка под агентом (краткая версия)

Если вы ставите скилл агенту, достаточно:

```bash
cd <папка скилла> && ./install.sh && tests/smoke.py
```

`tests/smoke.py` генерирует свои образцы документов и прогоняет по ним все
скрипты. Если последняя строка — `ИТОГ: PASS`, скилл готов. Если что-то
упало, строка `FAIL` скажет, что именно.

### Проверка вручную

```bash
tests/smoke.py
```

От LibreOffice зависит примерно половина проверок; без него они помечаются
`SKIP`, а не роняют тест.

### Интерпретатор находят сами

Скрипты запускаются напрямую (`scripts/xlsx_dump.py файл.xlsx`) — shebang сам
находит нужный Python в таком порядке:

1. `.venv/bin/python` рядом со скриптом (создаёт `install.sh`);
2. `~/.local/share/office-skills/venv/bin/python` (старое место, если уже есть);
3. `python3` из PATH.

Третий вариант может не сработать, если библиотеки не ставились вообще — тогда
выполните `./install.sh`.

## LibreOffice

Нужен для конвертации форматов и для проверки рендера. Поиск идёт в таком
порядке (настраивается в `scripts/_common.py`):

1. переменная `SOFFICE=/путь/к/soffice`;
2. `prefix/opt/libreoffice*/program/soffice` **внутри этой папки**;
3. `~/.local/share/office-skills/prefix/opt/libreoffice*/program/soffice`;
4. `soffice`/`libreoffice` в PATH.

Поставить проще всего системным пакетом:

```bash
sudo apt install libreoffice        # Debian/Ubuntu
sudo dnf install libreoffice-core   # Fedora
brew install --cask libreoffice     # macOS
```

Без LibreOffice работают: чтение docx/pptx/xlsx/pdf, правки через
`python-docx`/`openpyxl`, `pdf_tool.py`, `xlsx_dump.py`, `pptx_outline.py`.
Не работают: `office_convert.py`, `render_check.sh`, `xlsx_recalc.py`,
`validate.py --soffice`.

Также нужны `pdftoppm` (для `render_check.sh`) и `pandoc` (для docx→md):

```bash
sudo apt install poppler-utils pandoc
```

## Быстрый старт

```bash
# прочитать презентацию
scripts/pptx_outline.py deck.pptx

# прочитать таблицу, включая формулы
scripts/xlsx_dump.py book.xlsx --formulas

# отредактировать и проверить, что ничего не сломалось
cp in.docx out.docx
# ... правка ...
scripts/validate.py out.docx --soffice
scripts/render_check.sh out.docx      # -> out-01.jpg, out-02.jpg, ...
```

Обязательный цикл проверки после любой правки описан в `SKILL.md`.
Пропускать его нельзя: половина битых документов — это «я всё поправил, доверьте».

## Что стоит знать перед правками

1. Всегда работайте с копией, исходник не трогайте.
2. OOXML — это ZIP с XML внутри. `unzip` → правки → `zip`, архив собирать
   **изнутри** распакованной папки.
3. Текст в docx часто нарезан на мелкие куски — сначала `merge_runs.py`.
4. `python-pptx` не умеет дублировать слайды, `openpyxl` ломает чужие диаграммы
   при сохранении сложных книг.
5. PDF — фиксированный рендер, а не модель документа: «переписать абзац посреди
   макета» невозможно без потери вёрстки.

Подробности — в `references/`.

## Ограничения

- **PDF не редактируется «как Word».** Замена текста делается локальной подменой
  области (закраска + вставка) и годится для коротких строк: номер, дата, ФИО.
- **Сканы без текстового слоя не редактируются** — сначала OCR.
- PyMuPDF — **AGPL / коммерческая** лицензия. Локально и для своих задач — ок,
  но если вы строите на нём свой продукт и распространяете его, нужна лиценензия
  или замена. Подробнее в `references/pdf.md`.

## Лицензия

MIT — см. [LICENSE](LICENSE). Берите, меняйте, используйте; упоминание автора
сохраняйте.
