"""
test_opencv_table.py — Borderless table detection using OpenCV projection profiles.

How it works:
  1. Binarise the table region (Otsu threshold)
  2. Horizontal projection → every visual text line is a row boundary
  3. Vertical projection → column boundaries found by large whitespace gaps
  4. Save an overlay + optional cell crops + optional OCR → Excel

Usage:
    python test_opencv_table.py
    python test_opencv_table.py --image Pages/2/2.png
    python test_opencv_table.py --col-gap 30        # tune column separator width
    python test_opencv_table.py --debug             # save binary + projection images
    python test_opencv_table.py --save-cells        # dump each cell as a PNG
    python test_opencv_table.py --save-excel        # OCR every cell → Excel

Table region defaults are from img2table analysis of Pages/1/1.png.
y2 is set intentionally BELOW img2table's boundary (2804) to capture the last row.
"""

import argparse
import os
import sys

import cv2
import numpy as np

# ── Defaults (Pages/1/1.png) ─────────────────────────────────────────────────
DEFAULT = dict(x1=260, y1=1000, x2=2380, y2=2950)


def binarize(bgr):
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    return binary  # 255 = ink, 0 = paper


def find_bands(proj, threshold_ratio, merge_gap, min_size):
    """
    Return [(start, end), ...] of contiguous ink bands in a 1-D projection vector.
    Bands closer than merge_gap pixels are joined; bands smaller than min_size dropped.
    """
    thresh = float(proj.max()) * threshold_ratio if proj.max() > 0 else 1
    ink = proj > thresh

    bands, start = [], None
    for i, v in enumerate(ink):
        if v and start is None:
            start = i
        elif not v and start is not None:
            bands.append([start, i])
            start = None
    if start is not None:
        bands.append([start, len(ink)])

    merged = []
    for b in bands:
        if merged and b[0] - merged[-1][1] <= merge_gap:
            merged[-1][1] = b[1]
        else:
            merged.append(b)

    return [(s, e) for s, e in merged if e - s >= min_size]


def group_by_gap(bands, gap):
    """Merge adjacent bands whose separation <= gap pixels (column grouping)."""
    if not bands:
        return []
    groups = [list(bands[0])]
    for s, e in bands[1:]:
        if s - groups[-1][1] <= gap:
            groups[-1][1] = e
        else:
            groups.append([s, e])
    return [tuple(g) for g in groups]


def save_debug(binary, row_bands, col_bands, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    h, w = binary.shape

    # Horizontal projection bar chart
    h_proj = np.sum(binary, axis=1)
    h_img = np.zeros((h, 400), np.uint8)
    mx = float(h_proj.max()) or 1
    for y, v in enumerate(h_proj):
        bar = int(v / mx * 399)
        cv2.line(h_img, (0, y), (bar, y), 200, 1)
    for s, e in row_bands:
        cv2.line(h_img, (0, s), (399, s), 255, 1)
        cv2.line(h_img, (0, e), (399, e), 128, 1)
    cv2.imwrite(os.path.join(out_dir, "h_projection.png"), h_img)

    # Vertical projection bar chart
    v_proj = np.sum(binary, axis=0)
    v_img = np.zeros((400, w), np.uint8)
    mx = float(v_proj.max()) or 1
    for x, v in enumerate(v_proj):
        bar = int(v / mx * 399)
        cv2.line(v_img, (x, 399), (x, 399 - bar), 200, 1)
    for s, e in col_bands:
        cv2.line(v_img, (s, 0), (s, 399), 255, 1)
        cv2.line(v_img, (e, 0), (e, 399), 128, 1)
    cv2.imwrite(os.path.join(out_dir, "v_projection.png"), v_img)

    cv2.imwrite(os.path.join(out_dir, "binary.png"), binary)
    print(f"Debug images -> {out_dir}/")


def main():
    parser = argparse.ArgumentParser(
        description="OpenCV projection-based borderless table detector."
    )
    parser.add_argument("--image", default="Pages/1/1.png")
    parser.add_argument("--x1", type=int, default=DEFAULT["x1"], help="Table left   (px)")
    parser.add_argument("--y1", type=int, default=DEFAULT["y1"], help="Table top    (px)")
    parser.add_argument("--x2", type=int, default=DEFAULT["x2"], help="Table right  (px)")
    parser.add_argument("--y2", type=int, default=DEFAULT["y2"], help="Table bottom (px)")
    parser.add_argument(
        "--row-merge-gap", type=int, default=4,
        help="Join text bands within this many px into one row (default 4)"
    )
    parser.add_argument(
        "--col-gap", type=int, default=25,
        help="Whitespace wider than this px marks a column boundary (default 25)"
    )
    parser.add_argument(
        "--row-threshold", type=float, default=0.02,
        help="Row ink-density threshold as fraction of row maximum (default 0.02)"
    )
    parser.add_argument(
        "--col-threshold", type=float, default=0.003,
        help="Column ink-density threshold as fraction of col maximum (default 0.003)"
    )
    parser.add_argument("--padding", type=int, default=2,
                        help="Pixels to trim inside each cell edge (default 2)")
    parser.add_argument("--debug",       action="store_true", help="Save projection debug images")
    parser.add_argument("--save-cells",  action="store_true", help="Save each cell as PNG")
    parser.add_argument("--save-excel",  action="store_true", help="OCR cells and export Excel")
    args = parser.parse_args()

    img = cv2.imread(args.image)
    if img is None:
        print(f"Image not found: {args.image}")
        sys.exit(1)

    img_h, img_w = img.shape[:2]
    x1 = max(0, args.x1);  y1 = max(0, args.y1)
    x2 = min(img_w, args.x2); y2 = min(img_h, args.y2)
    print(f"Image: {args.image}  ({img_w}×{img_h})")
    print(f"Table region: x={x1}..{x2}, y={y1}..{y2}")

    table_crop = img[y1:y2, x1:x2]
    binary = binarize(table_crop)

    # ── Row detection ────────────────────────────────────────────────────────
    h_proj = np.sum(binary, axis=1)
    row_bands = find_bands(
        h_proj,
        threshold_ratio=args.row_threshold,
        merge_gap=args.row_merge_gap,
        min_size=5,
    )
    print(f"\nVisual rows detected: {len(row_bands)}")
    for i, (s, e) in enumerate(row_bands):
        abs_y = y1 + s
        print(f"  row {i:3d}: relative y={s:4d}..{e:4d}  absolute y={abs_y:4d}  height={e-s}px")

    # ── Column detection ─────────────────────────────────────────────────────
    v_proj = np.sum(binary, axis=0)
    col_raw = find_bands(
        v_proj,
        threshold_ratio=args.col_threshold,
        merge_gap=3,
        min_size=5,
    )
    col_bands = group_by_gap(col_raw, gap=args.col_gap)
    print(f"\nColumns detected: {len(col_bands)}  (--col-gap={args.col_gap})")
    for i, (s, e) in enumerate(col_bands):
        print(f"  col {i:2d}: relative x={s:4d}..{e:4d}  absolute x={x1+s:4d}  width={e-s}px")

    # ── Debug ────────────────────────────────────────────────────────────────
    if args.debug:
        debug_dir = os.path.join(os.path.dirname(args.image), "_debug_projection")
        save_debug(binary, row_bands, col_bands, debug_dir)

    # ── Overlay ──────────────────────────────────────────────────────────────
    overlay = img.copy()
    for s, e in col_bands:
        ax1, ax2 = x1 + s, x1 + e
        cv2.line(overlay, (ax1, y1), (ax1, y2), (0, 200,   0), 1)  # green
        cv2.line(overlay, (ax2, y1), (ax2, y2), (0, 200,   0), 1)
    for s, e in row_bands:
        ay1, ay2 = y1 + s, y1 + e
        cv2.line(overlay, (x1, ay1), (x2, ay1), (0, 120, 255), 1)  # orange
        cv2.line(overlay, (x1, ay2), (x2, ay2), (0, 120, 255), 1)

    overlay_path = args.image.replace(".png", "_opencv_overlay.png")
    cv2.imwrite(overlay_path, overlay)
    print(f"\nOverlay saved -> {overlay_path}")
    print("  Orange lines = row boundaries,  Green lines = column boundaries")

    # ── Cell extraction / OCR / Excel ────────────────────────────────────────
    if not (args.save_cells or args.save_excel):
        return

    cells_dir = args.image.replace(".png", "_opencv_cells")
    os.makedirs(cells_dir, exist_ok=True)

    reader = None
    if args.save_excel:
        import easyocr
        print("\nLoading EasyOCR...")
        reader = easyocr.Reader(["vi", "en"])

    pad = args.padding
    data = []

    for r_idx, (rs, re) in enumerate(row_bands):
        row_data = []
        cry1 = y1 + rs + pad
        cry2 = y1 + re - pad
        for c_idx, (cs, ce) in enumerate(col_bands):
            crx1 = x1 + cs + pad
            crx2 = x1 + ce - pad
            cell = img[cry1:cry2, crx1:crx2]
            if cell.size == 0:
                row_data.append(None)
                continue
            if args.save_cells:
                cv2.imwrite(
                    os.path.join(cells_dir, f"row{r_idx:03d}_col{c_idx:02d}.png"),
                    cell,
                )
            if reader is not None:
                texts = reader.readtext(cell, detail=0)
                row_data.append(" ".join(texts).strip() or None)
            else:
                row_data.append(None)
        data.append(row_data)

    if args.save_excel:
        import pandas as pd
        df = pd.DataFrame(data)
        excel_path = args.image.replace(".png", "_opencv_table.xlsx")
        df.to_excel(excel_path, index=False, header=False)
        print(f"Excel saved -> {excel_path}")
        print(f"\n{df.to_string(max_colwidth=35)}")

    if args.save_cells:
        print(f"Cell PNGs saved -> {cells_dir}/")


if __name__ == "__main__":
    main()
