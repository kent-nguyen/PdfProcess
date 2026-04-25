import os
import argparse
import importlib
from openpyxl import load_workbook


def find_raw_file(page_dir, page_num):
    path = os.path.join(page_dir, f"raw_{page_num}.xlsx")
    return path if os.path.exists(path) else None


def refine_page(raw_path, out_path, formatters, fixers):
    wb = load_workbook(raw_path)
    ws = wb.active

    row_errors = {}  # {row: {col: original_bad_value}} — populated by formatters, read by fixers
    row_fixes = {}   # {row: [note, ...]} — populated by fixers

    for formatter in formatters:
        formatter(ws, row_errors)

    for fixer in fixers:
        fixer(ws, row_fixes, row_errors)

    # Append Fixed and Notes columns after the last data column
    fixed_col = ws.max_column + 1
    notes_col = ws.max_column + 2
    ws.cell(row=1, column=fixed_col).value = "Fixed"
    ws.cell(row=1, column=notes_col).value = "Notes"

    for row in range(2, ws.max_row + 1):
        if row in row_fixes:
            ws.cell(row=row, column=fixed_col).value = True
            ws.cell(row=row, column=notes_col).value = "; ".join(row_fixes[row])

    wb.save(out_path)

    fixed_count = len(row_fixes)
    print(f"  {fixed_count} row(s) fixed -> {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Refine raw OCR Excel data column by column.")
    parser.add_argument("--input", default="Pages",
                        help="Folder containing page sub-folders (default: Pages)")
    parser.add_argument("--page", type=int, default=None,
                        help="Process only this page number; omit to process all pages")
    parser.add_argument("--bank", default="bidv",
                        help="Bank format to use — selects fixers/<bank>.py (default: bidv)")
    args = parser.parse_args()

    fixer_module = importlib.import_module(f"banks.{args.bank}")
    formatters = getattr(fixer_module, "FORMATTERS", [])
    fixers = getattr(fixer_module, "FIXERS", [])
    print(f"Bank: {args.bank} — {len(formatters)} formatter(s), {len(fixers)} fixer(s) loaded")

    if args.page:
        pages = [args.page]
    else:
        pages = sorted(
            int(p) for p in os.listdir(args.input)
            if os.path.isdir(os.path.join(args.input, p)) and p.isdigit()
        )

    for page_num in pages:
        page_dir = os.path.join(args.input, str(page_num))
        raw_path = find_raw_file(page_dir, page_num)
        if not raw_path:
            print(f"No raw Excel found in {page_dir}, skipping.")
            continue
        out_path = os.path.join(page_dir, f"{page_num}.xlsx")
        print(f"Processing page {page_num} ({os.path.basename(raw_path)})...")
        refine_page(raw_path, out_path, formatters, fixers)


if __name__ == "__main__":
    main()
