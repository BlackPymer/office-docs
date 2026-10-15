#!/bin/sh
"true" '''\'
D="$(dirname "$0")"
for P in "$D/../.venv/bin/python" "$HOME/.local/share/office-skills/venv/bin/python"; do
  [ -x "$P" ] && exec "$P" "$0" "$@"
done
exec python3 "$0" "$@"
''' + ''   # конкатенация -> Python не примет трampoline за docstring модуля
"""Самодостаточный smoke-тест office-docs.

Не зависит от чужих файлов: генерирует свои образцы (docx/xlsx/pptx/pdf) и
прогоняет по ним все скрипты скилла. Запускайте после установки:

    tests/smoke.py

Ожидаемый итог — последняя строка `ИТОГ: PASS`. Любая FAIL-строка выше
означает, что что-то сломано или не установлено (см. сообщение).

Нужен установленный `install.sh`-ом venv. LibreOffice проверяется, но его
отсутствие не валит тест — только помечается как SKIP.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import pathlib

SKILL = pathlib.Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

passed: list[str] = []
failed: list[str] = []
skipped: list[str] = []


def check(name: str, proc: subprocess.CompletedProcess, want=(0,), contains=None) -> None:
    ok = proc.returncode in want
    if ok and contains:
        ok = contains in proc.stdout
    if ok:
        passed.append(name)
    else:
        failed.append(name)
        print(f"  FAIL {name}: rc={proc.returncode}")
        if proc.stdout.strip():
            print(f"       stdout: {proc.stdout.strip()[:400]}")
        if proc.stderr.strip():
            print(f"       stderr: {proc.stderr.strip()[:400]}")


def sh(name: str, *args: str, want=(0,), contains=None) -> subprocess.CompletedProcess:
    p = subprocess.run([str(a) for a in args], capture_output=True, text=True)
    check(name, p, want, contains)
    return p


def main() -> int:
    work = pathlib.Path(tempfile.mkdtemp(prefix="office_smoke_"))
    os.chdir(work)
    print(f"Рабочий каталог: {work}")

    # --- образцы -------------------------------------------------------
    import docx, openpyxl, pymupdf
    from pptx import Presentation
    from pptx.util import Inches

    d = docx.Document()
    d.add_heading("Smoke test", 0)
    d.add_paragraph("Абзац с кириллицей для проверки шрифтов.")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Имя", "Значение"
    t.cell(1, 0).text, t.cell(1, 1).text = "x", "42"
    d.save("in.docx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A3"], ws["B1"] = "x", 3, 4, "=A2+A3"
    wb.save("book.xlsx")

    prs = Presentation()
    for i in (1, 2):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        box.text_frame.text = f"Слайд {i}"
    prs.save("deck.pptx")

    pdf = pymupdf.open()
    for i in (1, 2):
        page = pdf.new_page()
        page.insert_text((72, 72), f"Page {i}")
    pdf.save("doc.pdf")
    pdf.close()

    # --- Excel ---------------------------------------------------------
    print("\n== Excel ==")
    sh("xlsx_dump", f"{SCRIPTS}/xlsx_dump.py", "book.xlsx")
    sh("xlsx_dump --formulas", f"{SCRIPTS}/xlsx_dump.py", "book.xlsx",
       "--formulas", contains="=A2+A3")
    sh("xlsx_recalc", f"{SCRIPTS}/xlsx_recalc.py", "book.xlsx", contains='"total_errors": 0')

    # --- PowerPoint ----------------------------------------------------
    print("== PowerPoint ==")
    sh("pptx_outline", f"{SCRIPTS}/pptx_outline.py", "deck.pptx", contains="Слайд 1")
    sh("pptx_dup_slide", f"{SCRIPTS}/pptx_dup_slide.py", "deck.pptx",
       "--source", "1", "--after", "1", "-o", "dup.pptx")
    sh("дубль дал 3 слайда", f"{SCRIPTS}/pptx_outline.py", "dup.pptx", contains="3 слайд")

    # --- PDF -----------------------------------------------------------
    print("== PDF ==")
    sh("pdf_tool info", f"{SCRIPTS}/pdf_tool.py", "info", "doc.pdf", contains="страниц: 2")
    sh("pdf_tool text", f"{SCRIPTS}/pdf_tool.py", "text", "doc.pdf", contains="Page 1")
    sh("pdf_tool rotate", f"{SCRIPTS}/pdf_tool.py", "rotate", "doc.pdf",
       "-o", "rot.pdf", "--deg", "90")
    sh("pdf_tool merge", f"{SCRIPTS}/pdf_tool.py", "merge", "-o", "m.pdf", "doc.pdf", "doc.pdf")
    sh("pdf_tool split", f"{SCRIPTS}/pdf_tool.py", "split", "doc.pdf",
       "-o", "s.pdf", "--pages", "1-1")

    # --- Word ----------------------------------------------------------
    print("== Word ==")
    sh("merge_runs", f"{SCRIPTS}/merge_runs.py", "in.docx", "-o", "out.docx")
    p = sh("office_convert docx->pdf", f"{SCRIPTS}/office_convert.py", "in.docx", "c.pdf")
    if p.returncode != 0:
        skipped.append("office_convert/render_check/validate --soffice — нужен LibreOffice")
        print("  SKIP: LibreOffice не найден, конвертация и рендер пропущены")
    else:
        sh("office_convert docx->md", f"{SCRIPTS}/office_convert.py", "in.docx", "c.md")
        sh("validate --soffice", f"{SCRIPTS}/validate.py", "in.docx", "--soffice",
           contains="ИТОГ: PASS")
        if shutil.which("pdftoppm"):
            sh("render_check.sh", "bash", f"{SCRIPTS}/render_check.sh", "in.docx")
            imgs = sorted(pathlib.Path(".").glob("in-*.jpg"))
            (passed if imgs else failed).append("render_check дал картинки")
            if not imgs:
                print("  FAIL render_check: jpg не появились")
        else:
            skipped.append("render_check — нет pdftoppm (poppler-utils)")

    # --- деградации, которые не должны быть ошибками -------------------
    print("== Пайпы ==")
    sh("pptx_outline | head -1 не даёт traceback", "bash", "-c",
       f'{SCRIPTS}/pptx_outline.py deck.pptx 2>&1 | head -20 | grep -ci traceback > c.txt; '
       f'[ "$(cat c.txt)" = "0" ]')

    print("== Без аргументов печатают usage и код 2 ==")
    for s in ("validate", "office_convert", "pdf_tool", "xlsx_dump",
              "xlsx_recalc", "pptx_outline", "merge_runs"):
        sh(f"{s}.py без аргументов", f"{SCRIPTS}/{s}.py", want=(2,))

    print()
    for s in skipped:
        print(f"SKIP: {s}")
    print(f"ИТОГ: PASS={len(passed)}  FAIL={len(failed)}  SKIP={len(skipped)}")
    shutil.rmtree(work, ignore_errors=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
