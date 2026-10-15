---
name: office-docs
description: Редактирование и создание документов — Word (.docx, .dotx, .doc), PowerPoint (.pptx, .potx, .ppt), Excel (.xlsx, .xlsm, .xls), PDF (.pdf), а также конвертация между ними. Используй при любом упоминании файлов этих форматов: чтение, правка, создание отчётов/презентаций/таблиц, заполнение pdf-форм, объединение и разбиение pdf, замена текста в документе, преобразование legacy-файлов (.doc/.ppt/.xls), проверка качества вёрстки. Триггеры: «ворд», «документ», «docx», «слайды», «презентация», «pptx», «таблица», «excel», «xlsx», «pdf-форма», «отчёт», «договор», «бэкап в pdf». Скрипты и документация лежат в этой папке (scripts/, references/).
---

# Office-документы: чтение, правка, создание, проверка

## Среда (одноразово понять, дальше не гадать)

| Что | Путь/команда |
|---|---|
| Python со всеми библиотеками | `~/.local/share/office-skills/venv/bin/python` (docx, pptx, openpyxl, pymupdf, pypdf, lxml, markitdown). В клонах репозитория это `$SK/.venv/bin/python` — создаёт `install.sh` |
| LibreOffice | `~/.local/share/office-skills/prefix/opt/libreoffice26.2/program/soffice` — скрипты находят его сами (ищут также `$SK/prefix/opt/`, переменную `SOFFICE` и PATH); см. README.md |
| pandoc | `~/.local/bin/pandoc` (в PATH) |
| pdftoppm | в PATH (pdf → jpg для визуальной проверки) |

**Скрипты этой папки запускайте напрямую** (у них shebang на venv-python и +x):

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/pptx_outline.py deck.pptx
$SK/scripts/validate.py out.docx --soffice
$SK/scripts/render_check.sh out.docx      # печатает пути к jpg-страницам
```

Прямые вызовы библиотек тоже пишите venv-python (`~/.local/share/office-skills/venv/bin/python -c "..."`), иначе `import docx` не найдётся в системном python.

**Сами скрипты запускайте напрямую без python** (`$SK/scripts/xlsx_dump.py f.xlsx`) — их shebang сам находит нужный интерпретатор. Прямой вызов через `python3` без venv упадёт на `import docx`.

## Выбор подхода

| Задача | Подход |
|---|---|
| **Прочитать** docx/pptx/xlsx/pdf | `pandoc file.docx -t gfm` · `$SK/scripts/pptx_outline.py deck.pptx` · `$SK/scripts/xlsx_dump.py f.xlsx` · `markitdown f.pdf` |
| **Создать** новый .docx | Скрипт на `python-docx` (venv) |
| **Создать** новый .pptx | Скрипт на `python-pptx` (venv) |
| **Создать** новый .xlsx | Скрипт на `openpyxl` → обязательно `$SK/scripts/xlsx_recalc.py out.xlsx` |
| **Создать** PDF | Напрямую через `pymupdf`/`reportlab`, **или** создать docx/odt и `$SK/scripts/office_convert.py in.docx out.pdf` |
| **Отредактировать** существующий docx | Точечно python-docx (стили/фигуры сохраняются) **или** unzip → правка `word/document.xml` → zip (см. references/word.md) |
| **Отредактировать** существующий pptx | unzip → правка `ppt/slides/slideN.xml` → zip; дубль слайда — `pptx_dup_slide.py` (см. references/powerpoint.md) |
| **Отредактировать** xlsx | `openpyxl` (`keep_vba=True` для .xlsm) → `xlsx_recalc.py` |
| **PDF** | `pdf_tool.py` (формы, склейка, ротация, локальная замена текста) — см. references/pdf.md |
| **Legacy** .doc/.ppt/.xls | Сначала `$SK/scripts/office_convert.py a.doc a.docx`, потом работать с modern-форматом |

## Обязательный цикл проверки (не пропускайте)

После любой правки, прежде чем сказать «готово»:

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/validate.py out.docx --soffice      # структура + открываемость
$SK/scripts/render_check.sh out.docx            # -> out-01.jpg, out-02.jpg, ...
```

Затем **прочитайте картинки `Read`-ом и посмотрите глазами**: разъехались ли таблицы, нет ли обрезанного текста, правильные ли шрифты и колонки. Для xlsx вместо картинок — `xlsx_dump.py --formulas` + отчёт `xlsx_recalc.py` (там не должно быть `errors_found`).

Отчёт `validate.py` заканчивается строкой `ИТОГ: PASS` или `FAIL`. `FAIL` — не сдавайте, чините то, что он называет.

## Правила, на которых ломается 90% правок

1. **Работайте с копией.** `cp in.docx out.docx` и правьте `out.docx`; исходник не трогайте, пока пользователь не попросил именно это.
2. **ZIP-формат:** `unzip` → правки → `zip` (архив собирать **изнутри** распакованной папки: `(cd unpacked && zip -Xr ../out.docx .)`; если не сделать `rm -f ../out.docx`, удалённые части выживут в архиве).
3. **Не pretty-print'ить** OOXML-XML — отступы и переносы меняют содержимое текстовых узлов.
4. **Не переупорядочивать** элементы `<p:presentation>`/`<w:document>` — порядок частей схемо-зависим.
5. **Текст в docx часто нарезан на куски** (`<w:r>`) — фраза не находится. Сначала `$SK/scripts/merge_runs.py out.docx`, потом правьте.
6. **openpyxl ломает чужие диаграммы/картинки** при сохранении сложных книг — если книга богатая, правьте точечно XML или проверяйте рендером.
7. **python-pptx не дублирует слайды** и `text_frame.text = "..."` убивает форматирование — для текста меняйте `run.text`, для дублей используйте `pptx_dup_slide.py`.
8. **PDF не редактируется «как Word»** — см. честные ограничения в references/pdf.md.

## Скрипты

Путь относительно этой папки, `$SK` = `~/.config/opencode/skills/office-docs`:

| Скрипт | Что делает |
|---|---|
| `scripts/render_check.sh файл [префикс] [dpi]` | файл → PDF → `префикс-01.jpg…` в текущем каталоге (визуальная проверка) |
| `scripts/office_convert.py вход выход` | конвертация: pdf↔docx↔pptx↔xlsx, legacy (.doc/.ppt/.xls), html/rtf/txt, docx→md |
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
- `references/powerpoint.md` — .pptx: структура пакета, слайды, макеты, дублирование, типовые грабли
- `references/excel.md` — .xlsx: openpyxl, формулы, пересчёт, форматирование, чужие книги
- `references/pdf.md` — PDF: что можно и чего нельзя, формы, замена текста, склейка
- `references/legacy.md` — .doc/.ppt/.xls и прочие конвертации

## Лицензии

Код скилла — MIT (см. LICENSE), написан с нуля для этой среды. PyMuPDF — AGPL/коммерческая лицензия (локально ок, распространение продукта поверх — согласуйте). `markitdown` — MIT.
