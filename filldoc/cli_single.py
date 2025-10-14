# -*- coding: utf-8 -*-
# filldoc/cli_single.py
import os, glob, re
from docx import Document
from .utils import convert_doc_to_docx
from .generator import apply_mapping_to_doc
from .excel import load_excel  # должен вернуть ExcelConfig с data_rows, mapping_rows, settings

def _render_mask(mask: str, row: dict) -> str:
    def _repl(m):
        key = m.group(1).strip()
        val = row.get(key)
        return "" if val is None else str(val)
    return re.sub(r"\{\{\s*([^}]+?)\s*\}\}", _repl, mask)

def _process_one_doc(src_path: str, cfg, data_row: dict, outdir: str, dry_run: bool) -> None:
    os.makedirs(outdir, exist_ok=True)
    src = convert_doc_to_docx(src_path)
    d = Document(src)
    log: list[str] = []
    # mapping_rows и settings берём из Excel
    apply_mapping_to_doc(d, data_row, cfg.mapping_rows, cfg.settings, log, dry_run=dry_run)

    # имя файла вывода
    mask = cfg.settings.get("маска_имени_файла") or "{{ФИО}}.docx"
    out_name = _render_mask(mask, data_row) or "output.docx"
    out_path = os.path.join(outdir, out_name)

    if dry_run:
        print(f"[SINGLE] Предпросмотр '{out_name}':")
        for line in log:
            print("  -", line)
    else:
        d.save(out_path)
        print(f"[SINGLE] Сохранено: {out_path}")

def run_single(args):
    """
    Режим «одиночного» заполнения по Excel-маппингу:
      py -m filldoc.cli --doc ".\\input\\file.docx" --excel master.xlsx --row 1 --out ".\\output"
    Либо пакетом:
      py -m filldoc.cli --input ".\\input" --excel master.xlsx --out ".\\output"
    """
    if not args.excel:
        raise SystemExit("Укажите --excel с маппингом и листами (data, mapping, settings).")
    cfg = load_excel(args.excel)

    # Собираем список документов
    docs: list[str] = []
    if args.input:
        for f in glob.glob(os.path.join(args.input, "*.*")):
            if f.lower().endswith((".docx", ".doc")):
                docs.append(f)
    elif args.doc:
        docs.append(args.doc)
    else:
        raise SystemExit("Укажите --doc или --input.")

    # Выбор строк данных
    data_rows = cfg.data_rows
    if args.row is not None:
        idx = int(args.row)
        # считаем, что пользователь передаёт 1-based индекс (удобнее в Excel-контексте)
        if 1 <= idx <= len(data_rows):
            data_rows = [data_rows[idx - 1]]
        else:
            raise SystemExit(f"--row={idx} вне диапазона 1..{len(cfg.data_rows)}")

    # Прогон
    for path in docs:
        base = os.path.basename(path)
        for r in data_rows:
            try:
                _process_one_doc(path, cfg, r, args.out or "./output", args.dry_run)
            except Exception as e:
                print(f"[ERROR] {base}: {e}")
