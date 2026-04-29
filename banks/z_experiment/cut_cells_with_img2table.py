# cut_cells_with_img2table.py
# pip install img2table opencv-python numpy

import argparse
from pathlib import Path

import cv2
import numpy as np
from img2table.document import Image as Img2TableImage


def read_image_unicode(path: str):
    """
    Read image safely, including Unicode paths on Windows.
    """
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Cannot read image: {path}")
    return img


def write_image_unicode(path: str, img):
    """
    Write image safely, including Unicode paths on Windows.
    """
    ext = Path(path).suffix or ".png"
    ok, buf = cv2.imencode(ext, img)
    if not ok:
        raise ValueError(f"Cannot write image: {path}")
    buf.tofile(path)


def clamp(val, low, high):
    return max(low, min(val, high))


def crop_cell(img, x1, y1, x2, y2, padding=0):
    """
    Crop a cell with optional padding.
    """
    h, w = img.shape[:2]

    x1 = clamp(int(x1) - padding, 0, w)
    y1 = clamp(int(y1) - padding, 0, h)
    x2 = clamp(int(x2) + padding, 0, w)
    y2 = clamp(int(y2) + padding, 0, h)

    if x2 <= x1 or y2 <= y1:
        return None

    return img[y1:y2, x1:x2]


def export_cells(
    image_path: str,
    output_dir: str,
    implicit_rows: bool = True,
    implicit_columns: bool = False,
    borderless_tables: bool = False,
    padding: int = 1,
    unique_cells: bool = False,
):
    """
    Detect tables with img2table and crop each cell into a PNG file.

    unique_cells=False:
        save every row/col position as row000_col000.png
        If a merged cell spans multiple logical positions, it may appear duplicated.

    unique_cells=True:
        deduplicate cells by bounding box/value hash
        Better if you want only one file per physical cell region.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Original image for cropping
    original_img = read_image_unicode(image_path)

    # img2table document
    doc = Img2TableImage(image_path)

    tables = doc.extract_tables(
        ocr=None,  # OCR not required for cell coordinates
        implicit_rows=implicit_rows,
        implicit_columns=implicit_columns,
        borderless_tables=borderless_tables,
        min_confidence=50,
    )

    if not tables:
        raise RuntimeError("No tables detected.")

    print(f"Detected {len(tables)} table(s)")

    for table_idx, table in enumerate(tables):
        table_dir = output_dir / f"table_{table_idx:03d}"
        table_dir.mkdir(parents=True, exist_ok=True)

        print(f"\nTable {table_idx}:")
        print(f"  bbox = ({table.bbox.x1}, {table.bbox.y1}, {table.bbox.x2}, {table.bbox.y2})")
        print(f"  rows = {len(table.content)}")

        seen = set()
        saved_count = 0

        for row_idx, row in table.content.items():
            for col_idx, cell in enumerate(row):
                if cell is None or cell.bbox is None:
                    continue

                # Optional dedup for merged cells
                cell_key = (
                    cell.bbox.x1,
                    cell.bbox.y1,
                    cell.bbox.x2,
                    cell.bbox.y2,
                    cell.value,
                )
                if unique_cells and cell_key in seen:
                    continue
                seen.add(cell_key)

                crop = crop_cell(
                    original_img,
                    cell.bbox.x1,
                    cell.bbox.y1,
                    cell.bbox.x2,
                    cell.bbox.y2,
                    padding=padding,
                )

                if crop is None or crop.size == 0:
                    continue

                out_name = f"row{row_idx:03d}_col{col_idx:03d}.png"
                out_path = table_dir / out_name
                write_image_unicode(str(out_path), crop)
                saved_count += 1

        print(f"  saved {saved_count} cell image(s) to: {table_dir}")


def main():
    parser = argparse.ArgumentParser(description="Cut table cells into images using img2table")
    parser.add_argument("image", help="Input image path")
    parser.add_argument("output_dir", help="Output directory")
    parser.add_argument("--padding", type=int, default=1, help="Extra pixels around each crop")
    parser.add_argument(
        "--implicit-rows",
        action="store_true",
        help="Let img2table infer row boundaries when borders are weak",
    )
    parser.add_argument(
        "--implicit-columns",
        action="store_true",
        help="Let img2table infer column boundaries when borders are weak",
    )
    parser.add_argument(
        "--borderless-tables",
        action="store_true",
        help="Try to detect borderless tables",
    )
    parser.add_argument(
        "--unique-cells",
        action="store_true",
        help="Deduplicate merged cells / repeated physical cell regions",
    )

    args = parser.parse_args()

    export_cells(
        image_path=args.image,
        output_dir=args.output_dir,
        implicit_rows=args.implicit_rows,
        implicit_columns=args.implicit_columns,
        borderless_tables=args.borderless_tables,
        padding=args.padding,
        unique_cells=args.unique_cells,
    )


if __name__ == "__main__":
    main()