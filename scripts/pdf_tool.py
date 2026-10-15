#!/bin/sh
"true" '''\'
# Python-скрипт: sh-трampoline ниже запускает его нужным интерпретатором.
# Порядок: .venv рядом со скриптом -> старый пользовательский venv -> python3.
D="$(dirname "$0")"
for P in "$D/.venv/bin/python" "$HOME/.local/share/office-skills/venv/bin/python"; do
  [ -x "$P" ] && exec "$P" "$0" "$@"
done
exec python3 "$0" "$@"
''' + ''   # конкатенация -> Python не примет трampoline за docstring модуля
DOC = """Операции над PDF: чтение, формы, склейка/разделение, замена текста.

Использование:
    pdf_tool.py info    <file.pdf>
    pdf_tool.py text    <file.pdf> [--pages 1-5]
    pdf_tool.py forms   <file.pdf>
    pdf_tool.py fill    <file.pdf> -o out.pdf --set ИМЯ=ЗНАЧЕНИЕ [--set ...]
    pdf_tool.py merge   -o out.pdf <a.pdf> <b.pdf> [...]
    pdf_tool.py split   <file.pdf> -o out.pdf --pages 1-5 | --each [каталог]
    pdf_tool.py rotate  <file.pdf> -o out.pdf --deg 90 [--pages 1-5]
    pdf_tool.py replace <file.pdf> -o out.pdf --find "старый текст" --with "новый"

Честные ограничения PDF (подробнее в references/pdf.md):
  - это фиксированный рендер, а не модель документа: «переписать абзац
    посреди макета» невозможно без потери вёрстки; replace делает
    локальную подмену области (закраска + вставка) и годится для коротких
    строк (номер, дата, ФИО);
  - сканы без текстового слоя не редактируются — сначала OCR;
  - у PyMuPDF лицензия AGPL/коммерческая — для локальных задач это ок,
    для распространения своего продукта проконсультируйтесь.
"""
__doc__ = DOC


import argparse
import sys

import pymupdf

ERR_CODES = None


def open_pdf(path: str) -> pymupdf.Document:
    try:
        return pymupdf.open(path)
    except Exception as e:  # noqa: BLE001
        print(f"Не открыть PDF: {e}", file=sys.stderr)
        sys.exit(1)


def pages_sel(spec: str | None, doc: pymupdf.Document) -> list[int]:
    """'1-5,8' -> [0,1,2,3,4,7]"""
    if not spec:
        return list(range(len(doc)))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += list(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    return [i for i in out if 0 <= i < len(doc)]


def cmd_info(a) -> int:
    doc = open_pdf(a.file)
    md = doc.metadata or {}
    forms = sum(1 for p in doc for _ in (p.widgets() or []))
    annots = sum(1 for p in doc for _ in (p.annots() or []))
    print(f"страниц: {len(doc)}")
    print(f"ширинаxвысота первой страницы: {doc[0].rect.width:.0f}x{doc[0].rect.height:.0f} pt")
    print(f"зашифрован: {doc.is_encrypted}")
    print(f"полей формы (AcroForm): {forms}")
    print(f"аннотаций: {annots}")
    print(f"закладки (TOC): {len(doc.get_toc())}")
    for k in ("title", "author", "subject", "producer", "creationDate"):
        if md.get(k):
            print(f"{k}: {md[k]}")
    if doc.get_toc():
        print("TOC:")
        for lvl, title, page in doc.get_toc()[:30]:
            print(f"  {'  ' * (lvl - 1)}{title} -> с. {page}")
    return 0


def cmd_text(a) -> int:
    doc = open_pdf(a.file)
    for i in pages_sel(a.pages, doc):
        print(f"\n===== страница {i + 1} =====")
        print(doc[i].get_text("text").rstrip())
    return 0


def cmd_forms(a) -> int:
    doc = open_pdf(a.file)
    found = 0
    for pno, page in enumerate(doc, 1):
        for w in page.widgets() or []:
            found += 1
            r = w.rect
            print(f"p{pno}  {w.field_name!r}  type={w.field_type_string}  "
                  f"value={w.field_value!r}  rect=({r.x0:.0f},{r.y0:.0f},{r.x1:.0f},{r.y1:.0f})")
    if not found:
        print("Полей формы нет (возможно, это скан — тогда нужен OCR)")
    return 0


def cmd_fill(a) -> int:
    pairs = {}
    for s in a.set or []:
        if "=" not in s:
            print(f"--set ждёт ИМЯ=ЗНАЧЕНИЕ, получено: {s}", file=sys.stderr)
            return 2
        k, v = s.split("=", 1)
        pairs[k] = v
    doc = open_pdf(a.file)
    filled, missing = [], []
    for page in doc:
        for w in page.widgets() or []:
            if w.field_name in pairs:
                w.field_value = pairs.pop(w.field_name)
                w.update()
                filled.append(w.field_name)
    if pairs:
        missing = list(pairs)
    doc.save(a.out, garbage=3, deflate=True)
    print(f"заполнено: {filled}")
    if missing:
        print(f"НЕ найдены поля: {missing} — см. forms {a.file}", file=sys.stderr)
        return 1
    print(f"сохранено: {a.out}")
    return 0


def cmd_merge(a) -> int:
    out = pymupdf.open()
    for f in a.files:
        src = open_pdf(f)
        out.insert_pdf(src)
    out.save(a.out, garbage=3, deflate=True)
    print(f"{len(out)} страниц -> {a.out}")
    return 0


def cmd_split(a) -> int:
    doc = open_pdf(a.file)
    idx = pages_sel(a.pages, doc)
    if a.each:
        import os

        os.makedirs(a.each, exist_ok=True)
        for i in idx:
            n = pymupdf.open()
            n.insert_pdf(doc, from_page=i, to_page=i)
            p = f"{a.each}/{a.file.rsplit('/', 1)[-1][:-4]}-p{i + 1:03d}.pdf"
            n.save(p)
            print(p)
        return 0
    out = pymupdf.open()
    out.insert_pdf(doc, from_page=min(idx), to_page=max(idx))
    out.save(a.out)
    print(f"страницы {min(idx) + 1}-{max(idx) + 1} -> {a.out}")
    return 0


def cmd_rotate(a) -> int:
    doc = open_pdf(a.file)
    for i in pages_sel(a.pages, doc):
        doc[i].set_rotation((doc[i].rotation + a.deg) % 360)
    doc.save(a.out, garbage=3, deflate=True)
    print(f"повёрнуто на {a.deg}° -> {a.out}")
    return 0


UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
)


def pick_font(page) -> tuple[str, "pymupdf.Font"]:
    """base-14 шрифты (helv и др.) не содержат кириллицы — для не-латиницы
    берём Unicode-шрифт с диска, иначе на месте символов появятся точки."""
    import os

    for path in UNICODE_FONT_CANDIDATES:
        if os.path.exists(path):
            page.insert_font(fontname="officefont", fontfile=path)
            return "officefont", pymupdf.Font(fontfile=path)
    return "helv", pymupdf.Font("helv")


def cmd_replace(a) -> int:
    doc = open_pdf(a.file)
    hits = 0
    for page in doc:
        rects = page.search_for(a.find)
        for r in rects:
            hits += 1
            page.add_redact_annot(r, fill=(1, 1, 1))
        if rects:
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
            fontname, font = pick_font(page)
            # подбираем кегль, чтобы строка влезала в исходную ширину
            r0 = rects[0]
            fs = max(6, min(r0.height * 0.85, 12))
            width = font.text_length(a.with_, fontsize=fs)
            if width > r0.width > 0:
                fs = max(6, fs * r0.width / width)
            page.insert_text(pymupdf.Point(r0.x0, r0.y1 - fs * 0.25), a.with_,
                             fontname=fontname, fontsize=fs, color=(0, 0, 0))
    if not hits:
        print(f"Текст не найден: {a.find!r} (ищется построчно, без переносов)", file=sys.stderr)
        return 1
    doc.save(a.out, garbage=3, deflate=True)
    print(f"заменено фрагментов: {hits} -> {a.out}")
    print("Проверьте результат глазами: render_check.sh " + a.out)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("info"); p.add_argument("file"); p.set_defaults(fn=cmd_info)
    p = sub.add_parser("text"); p.add_argument("file"); p.add_argument("--pages"); p.set_defaults(fn=cmd_text)
    p = sub.add_parser("forms"); p.add_argument("file"); p.set_defaults(fn=cmd_forms)
    p = sub.add_parser("fill"); p.add_argument("file"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--set", action="append"); p.set_defaults(fn=cmd_fill)
    p = sub.add_parser("merge"); p.add_argument("-o", "--out", required=True); p.add_argument("files", nargs="+")
    p.set_defaults(fn=cmd_merge)
    p = sub.add_parser("split"); p.add_argument("file"); p.add_argument("-o", "--out")
    p.add_argument("--pages"); p.add_argument("--each", nargs="?", const=".", default=None)
    p.set_defaults(fn=cmd_split)
    p = sub.add_parser("rotate"); p.add_argument("file"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--deg", type=int, default=90); p.add_argument("--pages")
    p.set_defaults(fn=cmd_rotate)
    p = sub.add_parser("replace"); p.add_argument("file"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--find", required=True); p.add_argument("--with", dest="with_", required=True)
    p.set_defaults(fn=cmd_replace)

    a = ap.parse_args()
    if a.cmd == "split" and not (a.out or a.each):
        ap.error("split требует -o out.pdf или --each [каталог]")
    return a.fn(a)


if __name__ == "__main__":
    sys.exit(main())
