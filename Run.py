"""
Run the full PDF-to-Excel pipeline for a range of pages.

Usage:
    python Run.py --page 1-5 --bank bidv
    python Run.py --page 1,3,7-10 --bank bidv
"""

import os
import sys
import argparse
import importlib
import warnings
warnings.filterwarnings("ignore", message=".*pin_memory.*")

import easyocr

from libs.page_range import parse_pages
from ConvertToImages import convert_pages
from Deskew import deskew_page
from SplitTableCells import split_table_cells
from CellsToExcel import cells_to_excel
from RefineData import find_raw_file, refine_page, _print_summary


def _header(label):
    print(f"\n{'='*60}\n  {label}\n{'='*60}")


def main():
    parser = argparse.ArgumentParser(description="Run the full PDF-to-Excel pipeline.")
    parser.add_argument(
        "--page",
        required=True,
        help="Pages to process, e.g. '1' or '1,3,5' or '2-5' or '1-3,7,10-12'.",
    )
    parser.add_argument(
        "--bank",
        default="bidv",
        help="Bank format code (default: bidv).",
    )
    parser.add_argument(
        "--bottom",
        action="store_true",
        help="Detect deskew angle using only lines in the bottom 40%% of the image.",
    )
    args = parser.parse_args()

    pages = parse_pages(args.page)
    print(f"Pipeline: {len(pages)} page(s) — {pages} — bank={args.bank}")

    # Load bank config once — shared by CellsToExcel and RefineData
    bank_module = importlib.import_module(f"banks.{args.bank}")
    column_allowlists = getattr(bank_module, "COLUMN_ALLOWLISTS", {})
    formatters = getattr(bank_module, "FORMATTERS", [])
    fixers = getattr(bank_module, "FIXERS", [])
    garbage_date_cols = getattr(bank_module, "GARBAGE_DATE_COLS", None)
    print(f"Ngân hàng: {args.bank} — {len(formatters)} bộ định dạng, {len(fixers)} bộ sửa lỗi")

    # Initialize EasyOCR reader once — model load is slow, reuse across all pages
    print("Đang khởi tạo EasyOCR reader...")
    reader = easyocr.Reader(["en"])

    page_results = {}

    for page in pages:
        p = str(page)
        page_dir = os.path.join("Pages", p)

        _header(f"Step 1/5 — ConvertToImages  (page {p})")
        convert_pages(p)

        _header(f"Step 2/5 — Deskew           (page {p})")
        deskew_page(page, use_bottom=args.bottom)

        _header(f"Step 3/5 — SplitTableCells  (page {p})")
        img_path = os.path.join(page_dir, f"{p}.png")
        cells_dir = os.path.join(page_dir, "Cells")
        print(f"Đang xử lý trang {page}...")
        count = split_table_cells(img_path, cells_dir)
        print(f"  Đã trích xuất {count} ô -> {cells_dir}")

        _header(f"Step 4/5 — CellsToExcel     (page {p})")
        print(f"Đang xử lý trang {page}...")
        cells_to_excel(page_dir, reader, column_allowlists)

        _header(f"Step 5/5 — RefineData        (page {p})")
        raw_path = find_raw_file(page_dir, page)
        if not raw_path:
            print(f"  Không tìm thấy file Excel thô trong {page_dir}, bỏ qua.")
            continue
        out_path = os.path.join(page_dir, f"{p}.xlsx")
        print(f"Đang xử lý trang {page} ({os.path.basename(raw_path)})...")
        page_results[page] = refine_page(raw_path, out_path, formatters, fixers, garbage_date_cols)

    _print_summary(page_results)

    print(f"\n{'='*60}\n  Done. Output in Pages/<page>/<page>.xlsx\n{'='*60}\n")


if __name__ == "__main__":
    main()
