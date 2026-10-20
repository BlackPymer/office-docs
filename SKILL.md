---
name: office-docs
description: "Редактирование и создание документов — Word (.docx, .dotx, .doc), PowerPoint (.pptx, .potx, .ppt), Excel (.xlsx, .xlsm, .xls), PDF (.pdf), а также конвертация между ними. Работает на Linux/macOS и Windows. Используй при любом упоминании файлов этих форматов: чтение, правка, создание отчётов/презентаций/таблиц, заполнение pdf-форм, объединение и разбиение pdf, замена текста в документе, преобразование legacy-файлов (.doc/.ppt/.xls), проверка качества вёрстки. Триггеры: «ворд», «документ», «docx», «слайды», «презентация», «pptx», «таблица», «excel», «xlsx», «pdf-форма», «отчёт», «договор», «бэкап в pdf». Скрипты и документация лежат в этой папке (scripts/, references/)."
---

# Office-документы: чтение, правка, создание, проверка

## Среда (одноразово понять, дальше не гадать)

| Что | Где это берётся |
|---|---|
| Python со всеми библиотеками | venv скилла: `~/.local/share/office-skills/venv/bin/python` или `<папка скилла>/.venv/bin/python` (Linux/macOS), `<папка скилла>\.venv\Scripts\python.exe` (Windows). Ставится `install.py` / `install.sh` / `install.ps1`. Библиотеки: docx, pptx, openpyxl, pymupdf, pypdf, lxml, markitdown (опционально) |
| LibreOffice | система или портативная сборка: `prefix/opt/libreoffice*/program/soffice` (Linux/macOS), `prefix/LibreOffice*/program/soffice.exe` (Windows), `%ProgramFiles%\LibreOffice\program\soffice.exe`, `/Applications/LibreOffice.app/Contents/MacOS/soffice`, PATH. Скрипты находят его сами; можно задать `SOFFICE` |
| pandoc | нужен только для конвертации в `.md`: PATH, `~/.local/bin/pandoc`, `%LOCALAPPDATA%\Pandoc\pandoc.exe` |
| poppler/pdftoppm | **не нужен**: страницы рисует PyMuPDF из зависимостей (pdftoppm используется, только если он есть) |

**Как запускать скрипты этой папки.** На Linux/macOS — напрямую, `+x` и она сама находит нужный Python:

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/pptx_outline.py deck.pptx
```

На Windows shebang не работает, поэтому запускайте через `py -3` (или `python`) — интерпретатор скрипты подменяют сами:

```powershell
$SK = "$env:USERPROFILE\.config\opencode\skills\office-docs"
py -3 "$SK\scripts\pptx_outline.py" deck.pptx
```

**Свой код с библиотеками скилла** — через `scripts/py.py`, он сам находит venv (одинаково на всех платформах):

```bash
$SK/scripts/py.py -c "import docx; d=docx.Document('a.docx'); print([p.text for p in d.paragraphs])"
```

Прямой вызов `python -c "import docx"` системным Python упадёт на `ModuleNotFoundError` — библиотек там нет.

## Выбор подхода

| Задача | Подход |
|---|---|
| **Прочитать** docx/pptx/xlsx/pdf | `pandoc file.docx -t gfm` · `$SK/scripts/pptx_outline.py deck.pptx` · `$SK/scripts/xlsx_dump.py f.xlsx` · `markitdown f.pdf` |
| **Создать** новый .docx | Скрипт на `python-docx` (запускать через `py.py`) |
| **Создать** новый .pptx | Скрипт на `python-pptx` |
| **Создать** новый .xlsx | Скрипт на `openpyxl` → обязательно `$SK/scripts/xlsx_recalc.py out.xlsx` |
| **Создать** PDF | Напрямую через `pymupdf`/`reportlab`, **или** создать docx/odt и `$SK/scripts/office_convert.py in.docx out.pdf` |
| **Отредактировать** существующий docx | Точечно python-docx (стили/фигуры сохраняются) **или** `ooxml_zip.py unpack` → правка `word/document.xml` → `ooxml_zip.py pack` (см. references/word.md) |
| **Отредактировать** существующий pptx | `ooxml_zip.py unpack` → правка `ppt/slides/slideN.xml` → `ooxml_zip.py pack`; дубль слайда — `pptx_dup_slide.py` (см. references/powerpoint.md) |
| **Отредактировать** xlsx | `openpyxl` (`keep_vba=True` для .xlsm) → `xlsx_recalc.py` |
| **PDF** | `pdf_tool.py` (формы, склейка, ротация, локальная замена текста) — см. references/pdf.md |
| **Legacy** .doc/.ppt/.xls | Сначала `$SK/scripts/office_convert.py a.doc a.docx`, потом работать с modern-форматом |

## Обязательный цикл проверки (не пропускайте)

После любой правки, прежде чем сказать «готово»:

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/validate.py out.docx --soffice      # структура + открываемость
$SK/scripts/render_check.py out.docx            # -> out-01.jpg, out-02.jpg, ...
```

На Windows — то же через `py -3 "$SK\scripts\render_check.py" out.docx`. `render_check.sh` — тонкая обёртка вокруг `render_check.py` для Linux/macOS, отдельной логики в ней нет.

Затем **прочитайте картинки `Read`-ом и посмотрите глазами**: разъехались ли таблицы, нет ли обрезанного текста, правильные ли шрифты и колонки. Для xlsx вместо картинок — `xlsx_dump.py --formulas` + отчёт `xlsx_recalc.py` (там не должно быть `errors_found`).

Отчёт `validate.py` заканчивается строкой `ИТОГ: PASS` или `FAIL`. `FAIL` — не сдавайте, чините то, что он называет.

## Windows: что отличается

1. **Битых путей нет, но нет и unzip/zip.** Для XML-правок используйте `ooxml_zip.py`: `unpack`, `pack`, `list`, `cat` — он же гарантирует, что архив собирается с нуля (главная ловушка сборки `zip`-ом).
2. **Никогда не правьте пути на `/`-стиль** — используйте `py.py` и `pathlib`, а не shell-команды: `cp`, `rm`, `find`, `chmod`, `| head` в Windows-оболочке не те.
3. **Кодировка вывода**: если консоль ругается на кириллицу, ставьте `set PYTHONIOENCODING=utf-8` (или `$env:PYTHONIOENCODING="utf-8"`).
4. **LibreOffice**: `winget install TheDocumentFoundation.LibreOffice` (или choco/scoop). Портативную сборку можно распаковать в `prefix\` рядом со скиллом. Если бинарь в нестандартном месте — `set SOFFICE=C:\путь\к\soffice.exe`.
5. **Проверка**: `py -3 "$SK\tests\smoke.py"` — прогоняет все скрипты на сгенерированных образцах и печатает `ИТОГ: PASS`.

## Правила, на которых ломается 90% правок

1. **Работайте с копией.** `cp in.docx out.docx` (или `Copy-Item` на Windows) и правьте `out.docx`; исходник не трогайте, пока пользователь не попросил именно это.
2. **ZIP-формат:** `ooxml_zip.py unpack out.docx` → правки → `ooxml_zip.py pack unpacked out.docx`. Архив собирается с нуля, тип сжатия частей берётся из исходника. Ручная сборка `zip`-ом опасна: он **дописывает** записи и не удаляет старые — удалённая часть останется в файле.
3. **Не pretty-print'ить** OOXML-XML — отступы и переносы меняют содержимое текстовых узлов.
4. **Не переупорядочивать** элементы `<p:presentation>`/`<w:document>` — порядок частей схемо-зависим.
5. **Текст в docx часто нарезан на куски** (`<w:r>`) — фраза не находится. Сначала `$SK/scripts/merge_runs.py out.docx`, потом правьте.
6. **openpyxl ломает чужие диаграммы/картинки** при сохранении сложных книг — если книга богатая, правьте точечно XML или проверяйте рендером.
7. **python-pptx не дублирует слайды** и `text_frame.text = "..."` убивает форматирование — для текста меняйте `run.text`, для дублей используйте `pptx_dup_slide.py`.
8. **PDF не редактируется «как Word»** — см. честные ограничения в references/pdf.md.

## Скрипты

Путь относительно этой папки, `$SK` = папка скилла. Все они сами находят Python из venv (на Windows — если запускать их любым `py -3`/`python`).

| Скрипт | Что делает |
|---|---|
| `scripts/render_check.py файл [префикс] [dpi]` | файл → PDF → `префикс-01.jpg…` в текущем каталоге (визуальная проверка). Работает на Windows и POSIX |
| `scripts/render_check.sh файл [префикс] [dpi]` | обёртка для Linux/macOS над `render_check.py` (оставлена для совместимости) |
| `scripts/ooxml_zip.py {list,cat,unpack,pack}` | работа с ZIP-пакетом OOXML без `unzip`/`zip` (в т.ч. на Windows) |
| `scripts/py.py -c "..."` | запустить свой код интерпретатором из venv скилла |
| `scripts/office_convert.py вход выход` | конвертация: всё → .pdf, docx→odt/html/txt/md, pptx→odp/html, xlsx→ods/html, legacy (.doc/.ppt/.xls) → современные. **pdf→docx и кросс-конвертации docx/pptx/xlsx не работают** (нет export filter у LibreOffice) — см. references/legacy.md |
| `scripts/validate.py файл [--soffice]` | ZIP/XML/ссылки/открываемость; печатает `ИТОГ: PASS/FAIL` |
| `scripts/merge_runs.py файл.docx [-o out]` | склейка фрагментированных runs (нужно перед find/replace) |
| `scripts/pptx_outline.py deck.pptx` | текстовая выгрузка всех слайдов и фигур |
| `scripts/pptx_dup_slide.py deck.pptx --source N [--after M] -o out.pptx` | дублирование слайда с bookkeeping |
| `scripts/xlsx_dump.py f.xlsx [--formulas] [--sheet X]` | содержимое книги: значения или формулы, ячейки с ошибками помечает `!!` |
| `scripts/xlsx_recalc.py f.xlsx [timeout]` | пересчёт формул через LibreOffice + JSON-отчёт об ошибках |
| `scripts/pdf_tool.py {info,text,forms,fill,merge,split,rotate,replace}` | операции над PDF |

## Документация

Читайте по мере надобности (не загружайте всё сразу):

- `references/word.md` — .docx: чтение, точечные правки, XML, стили, оглавление, проверка
- `references/powerpoint.md` — .pptx: структура пакета, слайды, макеты, дублирование, типичные грабли
- `references/excel.md` — .xlsx: openpyxl, формулы, пересчёт, форматирование, чужие книги
- `references/pdf.md` — PDF: что можно и чего нельзя, формы, замена текста, склейка
- `references/legacy.md` — .doc/.ppt/.xls и прочие конвертации
- `README.md` — установка (Windows/Linux/macOS), где что ищется

## Лицензии

Код скилла — MIT (см. LICENSE), написан с нуля для этой среды. PyMuPDF — AGPL/коммерческая лицензия (локально ок, распространение продукта поверх — согласуйте). `markitdown` — MIT.
