"""
Test script for img2table (Option 1) — borderless table extraction.
Usage:
    python test_img2table.py
    python test_img2table.py --image Pages/1/1.png
    python test_img2table.py --image Pages/2/2.png --page 2
"""
import argparse
import subprocess
import sys
import os

# Auto-install img2table if missing
try:
    import img2table
except ImportError:
    print("Installing img2table...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "img2table"])

import cv2
import numpy as np
from img2table.document import Image as Img2TableImage
from img2table.ocr import EasyOCR


def draw_table_overlay(src_path: str, tables, out_path: str):
    """Draw detected cell bounding boxes onto the source image and save."""
    img = cv2.imread(src_path)
    img_h = img.shape[0]
    colors = [
        (0, 120, 255),   # orange-blue
        (0, 200, 0),     # green
        (200, 0, 200),   # magenta
    ]
    for t_idx, table in enumerate(tables):
        color = colors[t_idx % len(colors)]
        tb = table.bbox
        cv2.rectangle(img, (tb.x1, tb.y1), (tb.x2, min(tb.y2, img_h - 1)), color, 3)

        # If this table came from a crop re-run, its cell coords are relative to the
        # crop top — shift them back into full-image space.
        y_off = getattr(table, '_y_offset', 0)

        content_keys = sorted(table.content.keys())
        all_cells = []
        for row_idx in content_keys:
            for cell in table.content[row_idx]:
                if cell is None:
                    continue
                cb = cell.bbox
                ry1, ry2 = cb.y1 + y_off, min(cb.y2 + y_off, img_h - 1)
                cv2.rectangle(img, (cb.x1, ry1), (cb.x2, ry2), color, 1)
                all_cells.append((cb, y_off))

    cv2.imwrite(out_path, img)
    print(f"Overlay saved -> {out_path}")


def print_table(table, table_idx: int):
    df = table.df
    print(f"\n{'='*70}")
    print(f"Table {table_idx + 1}  |  {df.shape[0]} rows × {df.shape[1]} cols"
          f"  |  bbox: {table.bbox}")
    print(f"{'='*70}")
    # Print with index for easy reading
    try:
        print(df.to_string(index=True, max_colwidth=40))
    except Exception:
        print(df)


def main():
    parser = argparse.ArgumentParser(description="Test img2table borderless table extraction.")
    parser.add_argument("--image", default="Pages/1/1.png", help="Path to the page image")
    parser.add_argument("--no-ocr", action="store_true",
                        help="Skip OCR (detect cell positions only, no text)")
    parser.add_argument("--confidence", type=int, default=50,
                        help="Minimum OCR confidence 0-100 (default 50)")
    parser.add_argument("--save-excel", action="store_true",
                        help="Save each table to an Excel file alongside the image")
    args = parser.parse_args()

    if not os.path.exists(args.image):
        print(f"Image not found: {args.image}")
        sys.exit(1)

    print(f"Image: {args.image}")

    # --- OCR setup ---
    if args.no_ocr:
        ocr = None
        print("OCR disabled — cell positions only.")
    else:
        print("Loading EasyOCR (Vietnamese + English)...")
        ocr = EasyOCR(lang=["vi", "en"])

    def extract(src_path):
        doc = Img2TableImage(src=src_path)
        return doc.extract_tables(
            ocr=ocr,
            implicit_rows=True,
            borderless_tables=True,
            min_confidence=args.confidence,
        )

    # --- Extract tables (first pass on full image) ---
    print("Extracting tables (borderless_tables=True)...")
    tables = extract(args.image)

    if not tables:
        print("\nNo tables detected. Try adjusting --confidence or check the image.")
        return

    # --- Re-run on tight crop if last row may be cut off ---
    # Strip both the page header (above the table) AND the large blank footer so
    # img2table's boundary detection extends to the actual last text line.
    img_full = cv2.imread(args.image)
    img_h = img_full.shape[0]
    bottom_pad = 300
    for i, table in enumerate(tables):
        tb = table.bbox
        crop_y1 = max(0, tb.y1 - 50)       # just above the table top
        crop_y2 = min(img_h, tb.y2 + bottom_pad)  # limited tail below
        if crop_y2 >= img_h:
            print(f"  Table {i+1}: already at image bottom, skipping re-run.")
            continue
        print(f"  Table {i+1}: re-running on crop y={crop_y1}..{crop_y2} "
              f"(original y2={tb.y2}, img_h={img_h})...")
        crop = img_full[crop_y1:crop_y2, :]
        tmp_path = args.image.replace(".png", "_tmp_crop.png")
        cv2.imwrite(tmp_path, crop)
        tables_crop = extract(tmp_path)
        os.remove(tmp_path)
        if not tables_crop:
            print(f"  Table {i+1}: re-run found no tables.")
            continue
        rows_before = table.df.shape[0]
        rows_after = tables_crop[0].df.shape[0]
        print(f"  Table {i+1}: first pass={rows_before} rows, re-run={rows_after} rows.")
        if rows_after > rows_before:
            print(f"  Table {i+1}: recovered {rows_after - rows_before} extra row(s).")
            tables[i] = tables_crop[0]
            # Store crop offset so overlay can map coordinates back to full image
            tables[i]._y_offset = crop_y1
        else:
            print(f"  Table {i+1}: re-run found no additional rows "
                  f"(img2table limitation — last row gap exceeds detection threshold).")

    print(f"\nDetected {len(tables)} table(s).")

    # --- Print results ---
    for i, table in enumerate(tables):
        print_table(table, i)

    # --- Overlay visualization ---
    overlay_path = args.image.replace(".png", "_img2table_overlay.png")
    draw_table_overlay(args.image, tables, overlay_path)

    # --- Optional Excel export ---
    if args.save_excel:
        base = os.path.splitext(args.image)[0]
        for i, table in enumerate(tables):
            excel_path = f"{base}_table{i+1}.xlsx"
            table.df.to_excel(excel_path, index=False)
            print(f"Excel saved -> {excel_path}")


if __name__ == "__main__":
    main()
