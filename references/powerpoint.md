# PowerPoint (.pptx, .potx, .ppt)

## Структура

Тоже ZIP+XML:

| Файл | Содержимое |
|---|---|
| `ppt/presentation.xml` | `<p:sldIdLst>` — **порядок слайдов** (список `r:id`) |
| `ppt/_rels/presentation.xml.rels` | `r:id` → файл слайда |
| `ppt/slides/slideN.xml` | содержимое слайда |
| `ppt/slides/_rels/slideN.xml.rels` | его картинки/графики/макет |
| `ppt/slideLayouts/*`, `ppt/slideMasters/*`, `ppt/theme/*` | макеты и тема |
| `ppt/media/*` | картинки |

Важно: **номер файла слайда ≠ номер слайда в презентации.** Порядок задаёт только `p:sldIdLst`.

XML-namespace: `p:` (presentation), `a:` (drawing — фигуры и текст).

## Чтение

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/pptx_outline.py deck.pptx          # слайды, фигуры, текст, таблицы, заметки
markitdown deck.pptx                           # Markdown, по блоку на слайд
```

## Создание (python-pptx)

```python
from pptx import Presentation
from pptx.util import Inches, Pt
p = Presentation()
p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)   # widescreen; до правки layout!
s = p.slides.add_slide(p.slide_layouts[6])                     # 6 = Blank
tb = s.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(8), Inches(1.2))
tb.text_frame.text = "Заголовок"                               # для простого текста ок
p.save("deck.pptx")
```

Грабли:
- **Индекс макета хрупкий** (`slide_layouts[6]`) — у шаблонов состав другой. Выбирайте по `layout.name`, а не по номеру.
- **Плейсхолдеры шаблона** (`add_slide(prs.slide_layouts[0])`) дают унаследованные шрифты/позиции темы; пустой Blank — свои координаты.
- Эмпирический порядок: сначала `slide_width/height`, потом слайды.
- Текст с форматированием — по `run`: `p = tb.text_frame.paragraphs[0]; r = p.add_run(); r.text = "..."`; `r.font.bold = True`.

## Редактирование

**а) Точечные правки текста (обычный случай):**

```python
from pptx import Presentation
p = Presentation("deck.pptx")
sl = p.slides[1]
for sh in sl.shapes:
    if sh.has_text_frame and sh.text_frame.text == "Старый текст":
        # run-by-run: не трогает форматирование
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                if "Старый" in run.text:
                    run.text = run.text.replace("Старый", "Новый")
p.save("out.pptx")
```

**Ни в коем случае** `text_frame.text = "..."` поверх существующего текста — это сносит все runs в абзаце и обнуляет шрифты/кегль/жирность. Меняйте `run.text`.

**б) unzip → правка `ppt/slides/slideN.xml` → zip** — для структуры: новых текстбоксов, групп, перемещения фигур. Координаты в EMU (914400 = 1 дюйм), отступы/интервалы — в сотых пункта (`spcBef`), размер шрифта — в сотых (`sz="3200"` = 32 пт).

```bash
rm -rf unpacked && unzip -q deck.pptx -d unpacked    # cwd = текущий каталог
# править unpacked/ppt/slides/slide2.xml
(cd unpacked && rm -f ../out.pptx && zip -Xqr ../out.pptx .)   # out.pptx окажется в cwd
```

**в) Дублирование слайда:**

```bash
$SK/scripts/pptx_dup_slide.py deck.pptx --source 2 --after 2 -o out.pptx
```
Копирует часть, чинит `[Content_Types].xml`, `presentation.xml.rels` и `<p:sldIdLst>`; убирает из копии связь на заметки/комментарии (иначе PowerPoint предложит восстановить файл). Копия **делится** картинками/графиками с оригиналом — правка графика на одном слайде изменит оба; заметки докладчика у копии нет.

**г) Удаление/перестановка слайдов** — правьте только `<p:sldIdLst>` в `presentation.xml` (перестановка `<p:sldId>`, удаление записи). Сами файлы слайдов останутся «сиротами» — они безобидны, но если хотите вычистить: удалите `ppt/slides/slideN.xml` + `_rels/slideN.xml.rels`, их Override в `[Content_Types].xml`, Relationship в `presentation.xml.rels` и запись в `p:sldIdLst` — **всё это**, иначе файл сломается.

## Грабли

1. **Текст с ведущими/концевыми пробелами** — `xml:space="preserve"` на `<a:t>`.
2. **Списки:** один пункт = один `<a:p>`; маркер наследуется от макета, свой только через `<a:buChar>`/`<a:buAutoNum>`/`<a:buNone>` в `<a:pPr><a:buFont…>`. Символ `•` в тексте даёт двойной маркер.
3. **Маршрут правки структуры:** сначала дубли/удаления/перестановки, потом содержимое слайдов — иначе продублируете уже отредактированное.
4. python-pptx **не читает** EMF/WMF-картинки (`UnidentifiedImageError` при `add_picture`) — используйте PNG/JPEG, либо XML-путь.
5. Не переупорядочивайте дочерние элементы `<p:presentation>` — схема требует конкретный порядок (`sldIdLst`, `notesMasterIdLst`, `sldMasterIdLst` — не в произвольной последовательности).
6. **Заметки докладчика** — `slide.notes_slide.notes_text_frame.text`, не текстбок на слайде.

## Проверка

```bash
$SK/scripts/validate.py out.pptx --soffice
$SK/scripts/pptx_outline.py out.pptx          # текстовая сверка состава
$SK/scripts/render_check.sh out.pptx deck     # -> deck-01.jpg ...
```
Смотрите картинки: не вылезают ли текстовые блоки за край, нет ли наложений, правильные ли макеты у дублей.
