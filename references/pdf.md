# PDF (.pdf)

## Честная карта возможного

PDF — **фиксированный рендер страницы**, а не описание документа. Текст лежит плюс-минус «глифами с координатами», вёрстка воспроизводится, а не пересчитывается.

| Задача | Реально? |
|---|---|
| Извлечь текст/таблицы/картинки/метаданные/TOC | ✅ отлично |
| Заполнить AcroForm-поля, поставить галочки | ✅ отлично |
| Объединить, разбить, повернуть, переставить страницы | ✅ отлично |
| Подсветка/аннотации/печати/водяные знаки | ✅ отлично |
| Локально заменить короткую строку (номер, дата, сумма) | ⚠️ аккуратно, `pdf_tool.py replace` |
| Переписать абзац посреди сложного макета | ❌ вёрстка поедет; правьте исходник (docx → pdf) |
| Отредактировать скан без текстового слоя | ❌ нужен OCR (tesseract **не установлен**) |
| Восстановить «источник» документа | ❌ этого файла не существует |

**Главное правило:** если исходник документа у вас есть (docx/odt/презентация) — правьте его и конвертируйте заново: `$SK/scripts/office_convert.py edited.docx out.pdf`. PDF-правки — только когда источника нет.

## Чтение

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/pdf_tool.py info  doc.pdf         # страницы, поля, аннотации, TOC, метаданные
$SK/scripts/pdf_tool.py text doc.pdf --pages 1-5
$SK/scripts/pdf_tool.py forms doc.pdf         # имена/типы/значения полей
markitdown doc.pdf                            # Markdown (для длинных документов в модель)
```

Таблицы из PDF:

```python
import pymupdf
doc = pymupdf.open("doc.pdf")
for page in doc:
    for t in page.find_tables():
        print(t.to_markdown())   # или t.to_pandas()
```

Если `text` возвращает кракозябры/пусто — это скан (нужен OCR) или кастомная кодировка шрифта (тогда как запасной вариант `page.get_textpage_ocr()` при установленном tesseract).

## Создание

⚠️ **Кириллица: base-14 шрифты (`helv` и др.) её не содержат** — текст сохранится как точки `···`. Для не-латиницы грузите Unicode-шрифт:

```python
import pymupdf
doc = pymupdf.open()
page = doc.new_page(width=595, height=842)          # A4 в пунктах (pt); 1 дюйм = 72 pt
page.insert_font(fontname="F", fontfile="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
page.insert_text((72, 100), "Договор № 15", fontname="F", fontsize=16)
page.insert_textbox(pymupdf.Rect(72, 120, 520, 300),
                    "Абзац с переносами…", fontname="F", fontsize=11, align=pymupdf.TEXT_ALIGN_LEFT)
doc.save("out.pdf")
```

Таблицы: `page.draw_rect()` + `page.insert_textbox()` (вручную) или сразу `pymupdf.Story` / `reportlab` для длинных документов.

## Формы

```bash
$SK/scripts/pdf_tool.py forms  in.pdf                              # какие поля есть
$SK/scripts/pdf_tool.py fill  in.pdf -o out.pdf \
    --set client_name="ООО «Ромашка»" --set agree=Yes
```
Имена полей **точные** (регистр, пробелы). Незаполненные поля скрипт перечислит в stderr — по ним видно, что искали не там. Поля-чекбоксы: значение `Yes`/`Off`.

## Локальная замена текста

```bash
$SK/scripts/pdf_tool.py replace in.pdf -o out.pdf --find "100000" --with "250 000 руб."
```
Как работает: поиск строки → закраска области белым (redaction) → вставка нового текста (Unicode-шрифт подхватывается сам, кегль подгоняется под ширину).

Ограничения: ищется **строка без переносов** в пределах одной строки; шрифт новой строки — DejaVu Sans, не исходный (визуально может отличаться); длина ограничена шириной исходной области. Ищется по извлечённому тексту — если в исходнике кириллица лежала без Unicode-маппинга, `--find` её не найдёт (проверьте `text` сначала). Всегда: `render_check.sh out.pdf` и посмотрите страницу.

## Прочие операции

```bash
$SK/scripts/pdf_tool.py merge -o all.pdf a.pdf b.pdf c.pdf
$SK/scripts/pdf_tool.py split in.pdf --each parts/       # по файлу на страницу
$SK/scripts/pdf_tool.py split in.pdf -o part.pdf --pages 1-5
$SK/scripts/pdf_tool.py rotate in.pdf -o out.pdf --deg 90 --pages 2
```

Аннотации (подсветка найденного):

```python
import pymupdf
doc = pymupdf.open("in.pdf")
page = doc[0]
quads = page.search_for("важный термин", quads=True)
page.add_highlight_annot(quads)
doc.save("highlight.pdf")
```

## Лицензия

PyMuPDF — **AGPL** (или коммерческая лицензия Artifex). Локальные задачи и внутренние инструменты — ок; если встраиваете в закрытый продукт для распространения — нужна коммерческая лицензия. Альтернатива без AGPL для базовых операций — `pypdf` (установлен в том же venv): merge/split/rotate/метаданные/шифрование.
