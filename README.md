# office-docs

Скилл для агентов: чтение, правка, создание и проверка документов —
Word (.docx/.dotx/.doc), PowerPoint (.pptx/.potx/.ppt), Excel (.xlsx/.xlsm/.xls),
PDF (.pdf), плюс конвертация между ними.

Работает на **Linux, macOS и Windows**: скрипты сами находят LibreOffice и
интерпретатор из venv скилла, поэтому их можно запускать напрямую (`./script.py`)
или любым `py -3 script.py` — зависимости подхватятся.

Это не библиотека-обёртка, а набор скриптов с рабочей инструкцией (`SKILL.md`),
рассчитанной на то, что их выполняет AI-агент. Скрипты маленькие и читаются
целиком — правьте под себя.

## Что внутри

| Файл | Что делает |
|---|---|
| `SKILL.md` | Инструкция для агента: какой подход выбрать, какие грабли есть |
| `scripts/validate.py` | Проверка файла: ZIP/XML/ссылки/открываемость. Печатает `ИТОГ: PASS/FAIL` |
| `scripts/render_check.py` | Файл → PDF → JPG постранично (визуальная проверка вёрстки). Основной скрипт, работает везде |
| `scripts/render_check.sh` | Обёртка над `render_check.py` для Linux/macOS (для совместимости) |
| `scripts/ooxml_zip.py` | `list`/`cat`/`unpack`/`pack` для ZIP-пакета OOXML — замена `unzip`/`zip`, которых нет на Windows |
| `scripts/py.py` | Запуск своего кода интерпретатором из venv скилла |
| `scripts/office_convert.py` | Конвертация: всё → .pdf, docx→odt/html/txt/md, pptx→odp/html, xlsx→ods/html, legacy (.doc/.ppt/.xls) → современные. pdf→docx не работает — см. `references/legacy.md` |
| `scripts/merge_runs.py` | Склейка фрагментированных runs в docx (нужно перед поиском/заменой текста) |
| `scripts/pptx_outline.py` | Текстовая выгрузка всех слайдов и фигур презентации |
| `scripts/pptx_dup_slide.py` | Дублирование слайда с правильным bookkeeping'ом |
| `scripts/xlsx_dump.py` | Содержимое книги: значения или формулы, ошибки помечает `!!` |
| `scripts/xlsx_recalc.py` | Пересчёт формул через LibreOffice + JSON-отчёт об ошибках |
| `scripts/pdf_tool.py` | PDF: info, text, формы, fill, merge, split, rotate, replace |
| `scripts/_common.py` | Поиск LibreOffice и venv-Python, запуск headless с временным профилем, ре-экзек |
| `install.py` | Установщик (venv + зависимости), общий для всех ОС |
| `install.sh`, `install.ps1` | Обёртки для Linux/macOS и Windows |
| `tests/smoke.py` | Самодостаточный smoke-тест: генерирует образцы и прогоняет все скрипты |
| `references/*.md` | Подробности по каждому формату (читать по мере надобности) |

## Установка

Нужен Python **3.10 или новее**. Вся логика установки в `install.py` — он
одинаков для всех ОС:

```bash
./install.sh                  # Linux/macOS — просто обёртка над install.py
python install.py --help      # все варианты
```

Windows:

```powershell
.\install.ps1                 # или: py -3 install.py
```

Скрипт создаёт venv **внутри папки скилла** (`.venv/`), ставит зависимости из
`requirements.txt` (версии зафиксированы), проверяет импорты и компиляцию,
затем ищет LibreOffice и подсказывает, как его поставить.

Куда именно ставить venv, можно выбрать:

```bash
python install.py --venv ~/.local/share/office-skills/venv
```

Дополнительно, если нужен `markitdown` (docx/pptx/pdf → markdown):

```bash
./install.sh --with-markitdown      # это примерно +300 МБ
```

### Установка под агентом (краткая версия)

```bash
cd <папка скилла> && ./install.sh && tests/smoke.py         # Linux/macOS
py -3 install.py ; py -3 tests\smoke.py                     # Windows
```

`tests/smoke.py` генерирует свои образцы документов и прогоняет по ним все
скрипты. Если последняя строка — `ИТОГ: PASS`, скилл готов. Если что-то
упало, строка `FAIL` скажет, что именно.

### Интерпретатор находят сами

Скрипты можно запускать напрямую (`scripts/xlsx_dump.py файл.xlsx`), Windows —
`py -3 scripts\xlsx_dump.py файл.xlsx`. Нужный Python ищется в таком порядке:

1. `OFFICE_DOCS_PYTHON` — если задан;
2. `.venv/bin/python` и `.venv/Scripts/python.exe` в корне репозитория (создаёт `install.py`);
3. то же рядом со скриптом;
4. venv прошлых установок: `~/.local/share/office-skills/venv/bin/python`,
   `%LOCALAPPDATA%\office-skills\venv\Scripts\python.exe`;
5. `python3`/`python` из PATH.

На Linux/macOS первый шаг делает shebang-трampoline в начале файла, на Windows
shebang не работает — поэтому каждый скрипт при запуске сам проверяет, что
библиотеки импортируются, и при необходимости перезапускается найденным
интерпретатором. Если библиотек нет нигде, вы получите понятное сообщение с
просьбой выполнить установку, а не стек исключений.

Свой код (сниппеты на python-docx/openpyxl/pymupdf) запускайте так — тогда
библиотеки точно доступны:

```bash
scripts/py.py -c "import docx; print(docx.Document('a.docx').paragraphs[0].text)"
```

## LibreOffice

Нужен для конвертации форматов и для проверки рендера (кроме PDF на входе —
там рендер идёт только через PyMuPDF). Поиск идёт в таком порядке
(настраивается в `scripts/_common.py`):

1. переменная `SOFFICE=путь/к/soffice`;
2. `prefix/opt/libreoffice*/program/soffice` **внутри папки скилла** (Linux/macOS);
3. `prefix/LibreOffice*/program/soffice.exe` **внутри папки скилла** (Windows);
4. `~/.local/share/office-skills/prefix/...`, `%LOCALAPPDATA%\office-skills\prefix\...`;
5. системные установки: `%ProgramFiles%\LibreOffice\program\soffice.exe`,
   `%LocalAppData%\Programs\LibreOffice\program\soffice.exe`,
   `%USERPROFILE%\scoop\apps\libreoffice\current\program\soffice.exe`,
   `/Applications/LibreOffice.app/Contents/MacOS/soffice`;
6. `soffice`/`libreoffice` из PATH (в т.ч. shim от Chocolatey).

Поставить проще всего системным пакетом:

```bash
sudo apt install libreoffice        # Debian/Ubuntu
sudo dnf install libreoffice-core   # Fedora
brew install --cask libreoffice     # macOS
```

```powershell
winget install TheDocumentFoundation.LibreOffice   # Windows
choco install libreoffice-fresh                    # если есть Chocolatey
scoop install libreoffice                          # если есть Scoop
# или распакуйте портативную сборку в prefix\ рядом со скиллом
```

Без LibreOffice работают: чтение docx/pptx/xlsx/pdf, правки через
`python-docx`/`openpyxl`, `pdf_tool.py`, `xlsx_dump.py`, `pptx_outline.py`,
`ooxml_zip.py`, `render_check.py` для PDF на входе.
Не работают: `office_convert.py`, `xlsx_recalc.py`, `validate.py --soffice`,
`render_check.py` для docx/pptx/xlsx.

`pdftoppm` (poppler) больше **не требуется**: страницы рисует PyMuPDF. Poppler
используется, только если он есть в PATH. `pandoc` нужен лишь для конвертации
в Markdown:

```bash
sudo apt install pandoc        # Linux
winget install --id JohnMacFarlane.Pandoc   # Windows
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
scripts/render_check.py out.docx      # -> out-01.jpg, out-02.jpg, ...

# правка XML внутри docx/pptx (в т.ч. на Windows, где нет unzip/zip)
scripts/ooxml_zip.py unpack out.docx
scripts/ooxml_zip.py pack unpacked out.docx
```

Обязательный цикл проверки после любой правки описан в `SKILL.md`.
Пропускать его нельзя: половина битых документов — это «я всё поправил, доверьте».

## Что стоит знать перед правками

1. Всегда работайте с копией, исходник не трогайте.
2. OOXML — это ZIP с XML внутри. Распаковка и обратная сборка — только через
   `ooxml_zip.py`: архив собирается с нуля, поэтому удалённые части не
   «воскресают», а тип сжатия частей берётся из исходного файла.
3. Текст в docx часто нарезан на мелкие куски — сначала `merge_runs.py`.
4. `python-pptx` не умеет дублировать слайды, `openpyxl` ломает чужие диаграммы
   при сохранении сложных книг.
5. PDF — фиксированный рендер, а не модель документа: «переписать абзац посреди
   макета» невозможно без потери вёрстки.

Подробности — в `references/`.

## Windows: детали

- Скрипты запускаются как `py -3 "путь\scripts\validate.py" файл.docx --soffice`.
  `py` — лаунчер из python.org; если его нет, подойдёт `python`.
- `unzip`/`zip`/`cp`/`find`/`chmod` в Windows нет — используйте `ooxml_zip.py`,
  PowerShell (`Copy-Item`) и `py.py` вместо shell-сборки XML.
- Кириллица в пайпах требует UTF-8: `$env:PYTHONIOENCODING="utf-8"`.
- Проверка установки: `py -3 tests\smoke.py`. Windows-ветки путей в нём
  тестируются принудительно (`OFFICE_DOCS_FORCE_PLATFORM=windows`), поэтому
  smoke показывает, что логика поиска LibreOffice/venv работает и для Windows.

## Ограничения

- **PDF не редактируется «как Word».** Замена текста делается локальной подменой
  области (закраска + вставка) и годится для коротких строк: номер, дата, ФИО.
- **Сканы без текстового слоя не редактируются** — сначала OCR.
- PyMuPDF — **AGPL / коммерческая** лицензия. Локально и для своих задач — ок,
  но если вы строите на нём свой продукт и распространяете его, нужна лицензия
  или замена. Подробнее в `references/pdf.md`.

## Лицензия

MIT — см. [LICENSE](LICENSE). Берите, меняйте, используйте; упоминание автора
сохраняйте.
