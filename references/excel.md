# Excel (.xlsx, .xlsm, .csv)

## Чтение

```bash
SK=~/.config/opencode/skills/office-docs
$SK/scripts/xlsx_dump.py sales.xlsx                    # значения всех листов
$SK/scripts/xlsx_dump.py sales.xlsx --formulas --sheet Продажи
markitdown sales.xlsx                                  # Markdown по листам (без координат ячеек)
```

**Двухпроходное чтение (значения + формулы) — норма:**

```python
from openpyxl import load_workbook
wf = load_workbook("m.xlsx")                 # формулы как строки "=B2*C2"
wv = load_workbook("m.xlsx", data_only=True) # сохранённые результаты (2000)
```
`data_only=True` возвращает `None` для формул, если файл не пересчитывался после записи openpyxl.

## Создание и правка (openpyxl)

```python
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook(); ws = wb.active; ws.title = "Продажи"
ws["A1"] = "Товар"; ws["A1"].font = Font(bold=True, name="Arial", size=11)
ws["D2"] = "=B2*C2"            # ФОРМУЛА строкой — никогда не compute вручную
ws.column_dimensions["A"].width = 24
ws.freeze_panes = "A2"
wb.save("out.xlsx")
```

Правки существующей книги:

```python
wb = load_workbook("in.xlsx")
ws = wb["Лист1"]
ws["C5"] = 42
ws["C5"].number_format = '# ##0.00'
wb.save("out.xlsx")
```

## Пересчёт — обязательный шаг

```bash
$SK/scripts/xlsx_recalc.py out.xlsx        # JSON: status / total_formulas / total_errors
$SK/scripts/xlsx_dump.py out.xlsx          # значения должны быть числами, не None
```

openpyxl пишет формулы **без вычисленных значений**: пока не пересчитали, pandas/`data_only=True`/превьюеры видят пустоту. `status: errors_found` — не сдавать, правьте ячейки, что в `error_summary`.

**Что доказывает зелёный пересчёт:** формулы *вычисляются*, но не что *верны*. Смещение на строку (`=B2*C3`) даёт чистый файл с неправильными числами. Проверяйте 2–3 ячейки против ручного расчёта до того, как строить всю сетку формул.

## Грабли

1. **openpyxl разрушает то, что не понимает.** При `load_workbook` → `save` теряются: диаграммы, картинки, pivot tables, условное форматирование некоторых типов, некоторые элементы стиля темы. Если книга «богатая» — правьте точечно XML (см. references/word.md, общий ZIP-путь) или не трогайте лишнее. Проверяйте рендером: `render_check.sh out.xlsx`.
2. **`.xlsm`** — только `load_workbook(path, keep_vba=True)`, иначе макросы исчезнут.
3. **Внешние ссылки** на другие книги (`[1]Sheet1!A1`) при пересчёте LibreOffice превращаются в `#NAME?` — не пересчитывайте такие файлы (скрипт это не предотвращает; убедитесь, что `externalLinks/` в архиве отсутствует: `unzip -l out.xlsx | grep external`).
4. **Формулы должны ссылаться на ячейки**, а не содержать посчитанный результат: `=SUM(B2:B9)`, не `12345`.
5. **Шрифт** — задайте на используемый диапазон (`Arial`/`Times New Roman`), иначе просмотрщик подставит свой.
6. **Ошибки в ячейках** openpyxl отдаёт строками: `#DIV/0!`, `#REF!`, `#NAME?`, `#VALUE!`, `#N/A`, `#NULL!`, `#NUM!`. `#REF!` обычно означает удалённую строку/столбец, на которые ссылается формула.
7. Слияние ячеек (`merge_cells`) ломает многие формулы/фильтры — используйте только для заголовков.
8. **CSV → xlsx:** `pandas.read_csv(sep=";", encoding="utf-8-sig")` (для русских Excel-экспортов типичен `;` и BOM), дальше `df.to_excel` или запись по ячейкам с правильными типами (числа — числами, иначе сортировка сломается).

## Проверка

```bash
$SK/scripts/validate.py out.xlsx --soffice
$SK/scripts/xlsx_recalc.py out.xlsx          # status: success
$SK/scripts/render_check.sh out.xlsx sheet   # -> sheet-01.jpg… (лист как страница PDF)
```
На рендере смотрите: ширину колонок, переносы, обрезанные числа, что формулы дали ожидаемые значения.
