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
"""Дублирование слайда в .pptx с корректной регистрацией всех связей.

python-pptx не умеет дублировать слайд (только add_slide(layout) из шаблона),
а копирование файлов руками ломает package bookkeeping. Этот скрипт делает
всё, что нужно новому слайду:

  - копирует ppt/slides/slideN.xml и его .rels (сохраняя r:id внутри XML);
  - убирает из копии связи на notesSlide/комментарии (иначе PowerPoint
    увидит «двойную» обратную ссылку и предложит восстановить файл);
  - регистрирует часть в [Content_Types].xml;
  - добавляет Relationship в ppt/_rels/presentation.xml.rels;
  - добавляет <p:sldId> в <p:sldIdLst> (на нужную позицию).

Использование:
    pptx_dup_slide.py <deck.pptx> --source N [--after M] -o <out.pptx>

N и M — номера слайдов (с 1). --after M вставляет новый слайд сразу за M;
без --after — в конец. -o обязателен, чтобы не перезаписать ввод случайно.

ВАЖНО: новый слайд ссылается на те же части, что и источник (картинки,
графики, шаблон макета). Правка общего графика на одном слайде изменит и
другой; при необходимости вынесите график в отдельную часть или пересоздайте.
"""


import argparse
import os
import sys
import zipfile
from pathlib import Path

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_NOTES = "/notesSlide"
REL_COMMENT = "/comments"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"


def used_rids(xml: bytes) -> set[str]:
    """Все r:id-подобные ссылки, реально используемые внутри XML части."""
    root = etree.fromstring(xml)
    used: set[str] = set()
    for el in root.iter():
        for attr in el.attrib.values():
            pass
        for name, val in el.attrib.items():
            if name.startswith(f"{{{R}}}"):
                used.add(val)
    return used


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--source", type=int, required=True, help="номер дублируемого слайда (с 1)")
    ap.add_argument("--after", type=int, default=None, help="вставить после слайда N (по умолчанию в конец)")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()

    src_path = Path(args.deck).expanduser()
    if not src_path.is_file():
        print(f"Нет файла: {src_path}", file=sys.stderr)
        return 1

    with zipfile.ZipFile(src_path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
        infos = {i.filename: i for i in z.infolist()}

    pres = etree.fromstring(data["ppt/presentation.xml"])
    pres_rels = etree.fromstring(data["ppt/_rels/presentation.xml.rels"])
    rel_map = {r.get("Id"): r.get("Target") for r in pres_rels}

    sld_lst = pres.find(f"{{{P}}}sldIdLst")
    if sld_lst is None:
        print("В presentation.xml нет <p:sldIdLst>", file=sys.stderr)
        return 1
    sld_ids = list(sld_lst)
    if not (1 <= args.source <= len(sld_ids)):
        print(f"Слайда #{args.source} нет (всего {len(sld_ids)})", file=sys.stderr)
        return 1

    src_sldid = sld_ids[args.source - 1]
    src_rid = src_sldid.get(f"{{{R}}}id")
    src_part = "ppt/" + rel_map[src_rid]           # напр. ppt/slides/slide3.xml
    src_part = os.path.normpath(src_part).replace(os.sep, "/")

    # номера частей слайдов в архиве -> новый номер = max+1
    slide_nums = [
        int(n[len("ppt/slides/slide"):-4])
        for n in names
        if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        and n[len("ppt/slides/slide"):-4].isdigit()
    ]
    new_num = max(slide_nums) + 1
    new_part = f"ppt/slides/slide{new_num}.xml"
    new_rels_part = f"ppt/slides/_rels/slide{new_num}.xml.rels"
    src_rels_part = f"ppt/slides/_rels/{Path(src_part).name}.rels"

    if new_part in data:
        print(f"Конфликт: {new_part} уже существует", file=sys.stderr)
        return 1

    # 1) XML нового слайда — копия источника
    data[new_part] = data[src_part]

    # 2) rels: копия без notesSlide/комментариев, на которые нет ссылок в XML
    if src_rels_part in data:
        used = used_rids(data[new_part])
        rels_root = etree.fromstring(data[src_rels_part])
        for rel in list(rels_root):
            t = rel.get("Target", "")
            rtype = rel.get("Type", "")
            if (REL_NOTES in rtype or REL_COMMENT in rtype) and rel.get("Id") not in used:
                rels_root.remove(rel)
        data[new_rels_part] = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # 3) [Content_Types].xml — Override для новой части
    ct_root = etree.fromstring(data["[Content_Types].xml"])
    src_override = None
    for ov in ct_root.findall(f"{{{CT}}}Override"):
        if ov.get("PartName") == "/" + src_part:
            src_override = ov
    if src_override is None:
        print("В [Content_Types].xml нет Override для исходного слайда", file=sys.stderr)
        return 1
    new_ov = etree.SubElement(ct_root, f"{{{CT}}}Override")
    new_ov.set("PartName", "/" + new_part)
    new_ov.set("ContentType", src_override.get("ContentType"))
    data["[Content_Types].xml"] = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # 4) Relationship в presentation.xml.rels
    used_rids_all = {r.get("Id") for r in pres_rels}
    n = 1
    while f"rId{n}" in used_rids_all:
        n += 1
    new_rid = f"rId{n}"
    rel = etree.SubElement(pres_rels, f"{{{PR}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type", REL_SLIDE)
    rel.set("Target", f"slides/slide{new_num}.xml")
    data["ppt/_rels/presentation.xml.rels"] = etree.tostring(
        pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # 5) <p:sldId> в нужную позицию
    new_sldid = etree.Element(f"{{{P}}}sldId")
    used_ids = [int(s.get("id")) for s in sld_ids if (s.get("id") or "").isdigit()]
    new_sldid.set("id", str(max(used_ids + [255]) + 1))
    new_sldid.set(f"{{{R}}}id", new_rid)
    pos = args.source if args.after is None else args.after
    if pos >= len(sld_ids):
        sld_lst.append(new_sldid)
    else:
        sld_lst.insert(pos, new_sldid)
    data["ppt/presentation.xml"] = etree.tostring(
        pres, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # 6) запись архива
    out_path = Path(args.out).expanduser()
    tmp = out_path.with_suffix(out_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n_ in names:
            z.writestr(infos[n_], data[n_])
        z.writestr(new_part, data[new_part])
        if new_rels_part in data:
            z.writestr(new_rels_part, data[new_rels_part])
    os.replace(tmp, out_path)

    print(f"Создан {new_part} (копия {src_part});")
    print(f"слайдов стало: {len(sld_ids) + 1}; позиция: #{pos + 1 if pos < len(sld_ids) else len(sld_ids) + 1}")
    print(f"Записано: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
