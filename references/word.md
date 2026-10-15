# Word (.docx, .dotx, .doc)

## Структура

`.docx` — это ZIP с XML внутри. Что реально важно:

| Файл в архиве | Содержимое |
|---|---|
| `word/document.xml` | весь текст, таблицы, стили прямо в тексте |
| `word/styles.xml` | именованные стили (заголовки, «Нормальный», «Список») |
| `word/number.xml` | настройки нумерации/маркированных списков |
| `word/_rels/document.xml.rels` | ссылки на картинки, шрифты, шаблоны |
| `[Content_Types].xml` | какой тип у какой части (без неё Word не откроет) |
| `word/media/*` | картинки |

`w:` — пространство имён WordprocessingML. Текст лежит в `<w:t>` внутри `<w:r>` (run) внутри `<w:p>` (абзац).

## Чтение

```bash
pandoc report.docx -t gfm            # в Markdown с таблицами — основной способ
# быстрый греп без распаковки:
unzip -p report.docx word/document.xml | grep -o '<w:t[^>]*>[^<]*</w:t>' | head
# программно:
~/.local/share/office-skills/venv/bin/python -c "import docx; d=docx.Document('a.docx'); print([p.text for p in d.paragraphs])"
```

`markitdown report.docx` — альтернатива pandoc (тоже даёт Markdown).

## Три уровня редактирования

**1. python-docx — когда правятся абзацы/таблицы/стили, а структура пакета не трогается (обычный случай):**

```python
import docx
d = docx.Document("report.docx")
d.add_heading("Новый раздел", level=1)
d.add_paragraph("Текст", style="List Bullet")
t = d.add_table(rows=2, cols=3); t.style = "Table Grid"
d.add_comment(d.paragraphs[0].runs, "Проверить цифру", author="Редактор", initials="РЕД")
d.save("out.docx")
```

- **Точная правка текста внутри абзаца:** меняйте `run.text`, не `paragraph.text = ...` — присваивание `paragraph.text` выбрасывает все runs и **убивает** жирность/курсив/размер.
- `python-docx` сохраняет всё, что не трогали: картинки, стили, секции, колонтитулы, оглавление.
- Чего он **не умеет**: создавать сноски, менять разбивку на колонки, править fields (TOC/PAGE) напрямую — для этого XML-уровень.

**2. unzip → правка XML → zip — для всего, что выше python-docx:**

```bash
cp in.docx out.docx                       # работаем с копией, cwd = текущий каталог
rm -rf unpacked && unzip -q out.docx -d unpacked
find unpacked -type l -delete               # внешние docx могут содержать симлинки — не доверять
# править unpacked/word/document.xml (НЕ делать pretty-print/переформатирование!)
~/.config/opencode/skills/office-docs/scripts/merge_runs.py out.docx   # см. ниже — сначала склейка
(cd unpacked && rm -f ../out.docx && zip -Xqr ../out.docx .)   # zip ИЗНУТРИ unpacked; out.docx окажется в cwd
~/.config/opencode/skills/office-docs/scripts/validate.py out.docx --soffice
```

Почему `rm -f ../out.docx` обязателен: zip дописывает новые записи, но **не удаляет** старые — удалённый вручную кусок останется в архиве.

**3. Пересборка с нуля** — если вёрстка не важна (письмо, бейдж): скрипт на python-docx быстрее и надёжнее XML-правок.

## Грабли

1. **Фраза не находится в XML.** Word нарезает один абзац на много `<w:r>` (проверка орфографии, ревизии, вставки). «Это вступительный абзац» может физически лежать кусками. Перед грепом/заменой:
   ```bash
   ~/.config/opencode/skills/office-docs/scripts/merge_runs.py out.docx   # склеивает соседние runs с одинаковым rPr
   ```
   Склеивание визуально нейтрально (текст и форматирование не меняются).
2. **Не pretty-print'ить и не переупорядочивать** XML — `<w:body>` заканчивается `<w:sectPr>`, его порядок схемо-зависим, отступы в `<w:t>` становятся текстом.
3. **Оглавление (TOC)** — это поле `<w:fldSimple w:instr="TOC \o &quot;1-3&quot; ...">` или сложное поле. Word показывает его «обновите поле»; вставка:
   ```xml
   <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
   <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
   <w:r><w:t>Оглавление — обновите поле (F9)</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
   ```
   Заголовки **обязательно** должны быть стилями `Heading 1..9` (`w:pStyle w:val="Heading1"`), иначе пункт в оглавление не попадёт.
4. **Картинки.** python-docx `add_picture()` не понимает EMF/WMF/SVG (используются в шаблонах) — исключение `UnidentifiedImageError`. В таком случае копируйте картинку на XML-уровне (part `word/media/...` + relationship + `<w:drawing>` из соседнего абзаца-образца).
5. **Горизонтальная линейка** — это не таблица, а абзац с нижней границей:
   ```xml
   <w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/></w:pBdr></w:pPr>
   ```
6. **Размер страницы** задаётся в `<w:sectPr>` в DXA (1440 = 1 дюйм). US Letter: `w:pgSz w:w="12240" w:h="15840"`; A4: `11906x16838`.
7. **Текст с ведущими/концевыми пробелами** требует `xml:space="preserve"` на `<w:t>`.
8. **Списки** не создавайте символом `•` — используйте стиль/`w:numPr` (иначе получите двойные маркеры).
9. **Кириллица/шрифты:** у шрифта в `<w:rFonts w:ascii= w:hAnsi= w:cs=>` — при несоответствии русский текст падает в тему-шрифт и «плывёт». Проверяйте рендером, а не на глаз по XML.

## Проверка

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/validate.py out.docx --soffice   # ZIP/XML/ссылки/открываемость
$SK/scripts/render_check.sh out.docx         # -> out-01.jpg, ...
```
Картинки посмотрите через `Read`: таблицы, колонки, разрывы страниц, шрифты. Только после этого «готово».
