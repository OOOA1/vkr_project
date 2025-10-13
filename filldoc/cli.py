import argparse, sys
from .generator import fill_one

def main(argv=None):
    p = argparse.ArgumentParser(prog="filldoc", description="Автозаполнение Word по Excel-описанию.")
    p.add_argument("--doc", required=True, help="Путь к .docx")
    p.add_argument("--excel", required=True, help="Путь к .xlsx с листами data/mapping/settings")
    p.add_argument("--out", default="./output", help="Каталог для сохранения результатов")
    g = p.add_mutually_exclusive_group()
    g.add_argument("--all", action="store_true", help="Обработать все строки листа data")
    g.add_argument("--row", type=int, help="Обработать конкретную строку (1-based)")
    p.add_argument("--dry-run", action="store_true", help="Только анализ и сопоставление без записи")

    args = p.parse_args(argv)
    row_index = None if args.all or (args.row is None) else args.row
    code = fill_one(args.doc, args.excel, args.out, row_index=row_index, dry_run=args.dry_run)
    sys.exit(code)

if __name__ == "__main__":
    main()
