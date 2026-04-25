import os
import re
import argparse
import easyocr
from openpyxl import Workbook


def ocr_cell(reader, image_path):
    result = reader.readtext(image_path)
    if not result:
        return ""
    return " ".join(item[1] for item in result).strip()


def cells_to_excel(page_dir, reader):
    cells_dir = os.path.join(page_dir, "Cells")
    if not os.path.isdir(cells_dir):
        print(f"  No Cells folder in {page_dir}, skipping.")
        return

    pattern = re.compile(r"row(\d+)_col(\d+)\.png$")
    cell_files = {}
    for fname in os.listdir(cells_dir):
        m = pattern.match(fname)
        if m:
            cell_files[(int(m.group(1)), int(m.group(2)))] = os.path.join(cells_dir, fname)

    if not cell_files:
        print("  No cell images found.")
        return

    wb = Workbook()
    ws = wb.active

    total = len(cell_files)
    for i, (row, col) in enumerate(sorted(cell_files), 1):
        text = ocr_cell(reader, cell_files[(row, col)])
        ws.cell(row=row + 1, column=col + 1, value=text)
        print(f"  [{i}/{total}] row{row:03d}_col{col:03d} -> {repr(text)}")

    page_name = os.path.basename(page_dir)
    out_path = os.path.join(page_dir, f"raw_{page_name}.xlsx")
    wb.save(out_path)
    print(f"  Saved -> {out_path}")


def main():
    parser = argparse.ArgumentParser(description="OCR cell images and write to Excel.")
    parser.add_argument("--input", default="Pages",
                        help="Folder containing page sub-folders (default: Pages)")
    parser.add_argument("--page", type=int, default=None,
                        help="Process only this page number; omit to process all pages")
    args = parser.parse_args()

    reader = easyocr.Reader(["en"])

    if args.page:
        pages = [args.page]
    else:
        pages = sorted(
            int(p) for p in os.listdir(args.input)
            if os.path.isdir(os.path.join(args.input, p)) and p.isdigit()
        )

    for page_num in pages:
        page_dir = os.path.join(args.input, str(page_num))
        print(f"Processing page {page_num}...")
        cells_to_excel(page_dir, reader)


if __name__ == "__main__":
    main()
