# filldoc/cli.py
# -*- coding: utf-8 -*-
import argparse, os, glob, re
from docx import Document
from .utils import convert_doc_to_docx
from .detector import detect_template, suggest_anchors
from .templates import TEMPLATES

def render_mask(mask: str, row: dict) -> str:
    def _repl(m):
        key = m.group(1).strip()
        val = row.get(key)
        return "" if val is None else str(val)
    return re.sub(r"\{\{\s*([^}]+?)\s*\}\}", _repl, mask)

def _process_detect(path: str):
    try:
        src = convert_doc_to_docx(path)
    except Exception as e:
        print(f"[CONVERT-ERROR] {os.path.basename(path)}: {e}")
        return
    doc = Document(src)
    det = detect_template(doc)
    base = os.path.basename(path)
    if det.template_id:
        spec = TEMPLATES[det.template_id]
        print(f"[DETECT] '{base}' → {det.template_id} ({spec.human_name}), score={det.score}")
        for note in det.breakdown:
            print("  -", note)
    else:
        print(f"[DETECT] '{base}' → НЕ ОПОЗНАН (score={det.score})")
        for note in det.breakdown:
            print("  -", note)

def _process_suggest(path: str):
    try:
        src = convert_doc_to_docx(path)
    except Exception as e:
        print(f"[CONVERT-ERROR] {os.path.basename(path)}: {e}")
        return
    doc = Document(src)
    base = os.path.basename(path)
    hints = suggest_anchors(doc)
    print(f"\n[SUGGEST] {base}")
    print("  top_lines:")
    for s in hints["top_lines"]:
        print("   •", s)
    print("  with_colon:")
    for s in hints["with_colon"]:
        print("   •", s)
    print("  uppercase:")
    for s in hints["uppercase"]:
        print("   •", s)
    print("  keywords:")
    for s in hints["keywords"]:
        print("   •", s)
    # рыба для вставки в templates.py
    print("  --- TemplateSpec skeleton ---")
    print("  required = [")
    for s in (hints["uppercase"][:2] + hints["with_colon"][:3] + hints["keywords"][:1]):
        print(f"    \"{s}\",")
    print("  ]")
    print("  optional = [")
    for s in (hints["with_colon"][3:6] + hints["top_lines"][2:4]):
        print(f"    \"{s}\",")
    print("  ]")

def _process_auto(path: str, excel: str, row: int | None, outdir: str, dry_run: bool):
    from .excel import load_master_data
    try:
        from .generator import apply_mapping_to_doc
    except Exception as e:
        raise SystemExit("В filldoc/generator.py должна быть функция apply_mapping_to_doc(doc, data_row, mapping, settings, log, dry_run=False).") from e

    try:
        src = convert_doc_to_docx(path)
    except Exception as e:
        print(f"[CONVERT-ERROR] {os.path.basename(path)}: {e}")
        return
    base = os.path.basename(path)
    doc = Document(src)
    det = detect_template(doc)
    if not det.template_id:
        print(f"[AUTO] '{base}' → не опознан (score={det.score})")
        return
    spec = TEMPLATES[det.template_id]

    rows = load_master_data(excel)
    if row:
        rows = [rows[row-1]]

    os.makedirs(outdir, exist_ok=True)
    for r in rows:
        mask = spec.settings.get("маска_имени_файла", f"{spec.id}_{{{{ФИО}}}}.docx")
        out_name = render_mask(mask, r)
        out_path = os.path.join(outdir, out_name)
        log = []
        d = Document(src)
        apply_mapping_to_doc(d, r, spec.mapping, spec.settings, log, dry_run=dry_run)
        if dry_run:
            print(f"[{spec.id}] Предпросмотр '{out_name}':")
            for line in log:
                print("  -", line)
        else:
            d.save(out_path)
            print(f"[{spec.id}] Сохранено: {out_path}")

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--detect", action="store_true", help="распознать шаблон по 1-й странице")
    p.add_argument("--suggest", action="store_true", help="подсказать якоря по 1-й странице (для заполнения templates.py)")
    p.add_argument("--auto", action="store_true", help="распознать и заполнить")
    p.add_argument("--doc", help="входной .docx|.doc")
    p.add_argument("--input", help="папка с документами")
    p.add_argument("--excel", help="master.xlsx с листом data (для --auto)")
    p.add_argument("--row", type=int)
    p.add_argument("--out", default="./output")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args()

    if args.detect or args.suggest:
        handler = _process_suggest if args.suggest else _process_detect
        if args.input:
            for f in glob.glob(os.path.join(args.input, "*.*")):
                if f.lower().endswith((".docx", ".doc")):
                    handler(f)
        else:
            if not args.doc:
                raise SystemExit("Укажите --doc или --input")
            handler(args.doc)
        return

    if args.auto:
        if not args.excel:
            raise SystemExit("Для --auto нужен --excel (лист 'data').")
        if args.input:
            for f in glob.glob(os.path.join(args.input, "*.*")):
                if f.lower().endswith((".docx", ".doc")):
                    _process_auto(f, args.excel, args.row, args.out, args.dry_run)
        else:
            if not args.doc:
                raise SystemExit("Укажите --doc или --input")
            _process_auto(args.doc, args.excel, args.row, args.out, args.dry_run)
        return

    run_single(args)

if __name__ == "__main__":
    main()
