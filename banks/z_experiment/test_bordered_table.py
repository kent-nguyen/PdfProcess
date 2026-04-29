"""
test_bordered_table.py — Extract data from a bordered (grid-line) table image.

Works for BIDV bank statements and similar scanned tables with clear horizontal
and vertical rules. Uses morphological line detection + EasyOCR.

Usage:
    python test_bordered_table.py --image Pages/206/206.png
    python test_bordered_table.py --image Pages/206/206.png --save-excel
    python test_bordered_table.py --image Pages/206/206.png --debug
    python test_bordered_table.py --image Pages/206/206.png --no-ocr   # grid only, fast
"""

import argparse
import os
import sys

import cv2
import numpy as np


# ── Image I/O (Unicode-safe) ──────────────────────────────────────────────────

def read_image(path):
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Cannot read image: {path}")
    return img


def write_image(path, img):
    from pathlib import Path
    ext = Path(path).suffix or ".png"
    ok, buf = cv2.imencode(ext, img)
    if not ok:
        raise ValueError(f"Cannot encode: {path}")
    buf.tofile(path)


# ── Grid detection ────────────────────────────────────────────────────────────

def detect_grid(img, h_min_len_frac=0.015, v_min_len_frac=0.015,
                h_bridge_frac=0.05, v_bridge_frac=0.02,
                threshold_ratio=0.02, gap=10):
    """
    Detect horizontal and vertical table lines via short morphological open
    followed by gap-bridging dilation and projection-profile peak finding.

    Strategy:
      1. Short OPEN  — removes isolated text pixels but keeps any line fragment
                       longer than ~1.5% of the image dimension.
      2. Bridge dil  — merges nearby fragments along the same line into one blob
                       (handles breaks at column/row intersections in scanned docs).
      3. Projection  — sums each row/col; rows with many bridged pixels = line rows.
      4. Peak find   — threshold at 2% of max projection, cluster adjacent peaks.

    h_min_len_frac : min continuous horizontal fragment (fraction of image width).
    v_min_len_frac : min continuous vertical fragment (fraction of image height).
    h_bridge_frac  : horizontal dilation width to bridge inter-column gaps.
    v_bridge_frac  : vertical dilation height to bridge inter-row gaps.
    threshold_ratio: projection value must exceed this fraction of the row maximum.
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape

    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # ── Horizontal lines ──────────────────────────────────────────────────────
    h_len = max(int(w * h_min_len_frac), 20)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1))
    horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=1)
    # Bridge gaps between fragments along the same line
    h_bridge = cv2.getStructuringElement(cv2.MORPH_RECT, (max(int(w * h_bridge_frac), 20), 1))
    horiz = cv2.dilate(horiz, h_bridge, iterations=1)
    # Thicken vertically so thin lines project reliably
    horiz = cv2.dilate(horiz, np.ones((3, 1), np.uint8), iterations=2)

    # ── Vertical lines ────────────────────────────────────────────────────────
    v_len = max(int(h * v_min_len_frac), 20)
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_len))
    vert = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=1)
    # Bridge gaps at row intersections
    v_bridge = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(int(h * v_bridge_frac), 20)))
    vert = cv2.dilate(vert, v_bridge, iterations=1)
    # Thicken horizontally
    vert = cv2.dilate(vert, np.ones((1, 3), np.uint8), iterations=2)

    # ── Projection peak finder ────────────────────────────────────────────────
    def line_centers(mask, axis, threshold_ratio=threshold_ratio):
        proj = np.sum(mask, axis=axis).astype(float)
        if proj.max() == 0:
            return []
        thresh = proj.max() * threshold_ratio
        active = np.where(proj > thresh)[0]
        if len(active) == 0:
            return []
        groups, cur = [], [active[0]]
        for idx in active[1:]:
            if idx - cur[-1] <= gap:
                cur.append(idx)
            else:
                groups.append(int(np.mean(cur)))
                cur = [idx]
        groups.append(int(np.mean(cur)))
        return groups

    h_pos = line_centers(horiz, axis=1)   # horizontal lines → y positions

    # ── Vertical: scan header row only ───────────────────────────────────────
    # Column separators in scanned tables often appear only in the header row,
    # not in data rows. Cropping to the header and running a fresh detection
    # there finds all columns reliably.
    if len(h_pos) >= 2:
        # Use the first inter-line band as the header crop
        header_y1 = max(0, h_pos[0] - 50)   # a bit above the top border
        header_y2 = h_pos[1]                 # bottom of header row
        header_crop = binary[header_y1:header_y2, :]
        hh = header_crop.shape[0]

        # Require 50% of header height — true separators span it, text strokes don't
        v_len_hdr = max(int(hh * 0.50), 10)
        v_kernel_hdr = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_len_hdr))
        vert_hdr = cv2.morphologyEx(header_crop, cv2.MORPH_OPEN, v_kernel_hdr)
        vert_hdr = cv2.dilate(vert_hdr, np.ones((1, 3), np.uint8), iterations=2)

        # Higher threshold than general: 20% of max to reject residual text noise
        v_pos = line_centers(vert_hdr, axis=0, threshold_ratio=0.20)

        # Anchor v_pos to the actual table left/right borders.
        # The horizontal line mask spans the full table width, so its leftmost
        # and rightmost active columns are the true outer borders.
        h_proj_x = np.sum(horiz, axis=0)
        active_x = np.where(h_proj_x > 0)[0]
        if len(active_x):
            left_border  = int(active_x[0])
            right_border = int(active_x[-1])
            if not v_pos or v_pos[0] - left_border > gap:
                v_pos = [left_border] + v_pos
            if not v_pos or right_border - v_pos[-1] > gap:
                v_pos = v_pos + [right_border]

        # Rebuild the full-height vert mask for debug output
        vert = np.zeros_like(binary)
        for x in v_pos:
            vert[:, max(0, x-1):x+2] = 255
    else:
        v_pos = line_centers(vert, axis=0)

    return h_pos, v_pos, horiz, vert


# ── Cell enhancement ──────────────────────────────────────────────────────────

def enhance_cell(cell_bgr):
    """2× upscale + unsharp mask for better OCR on small cells."""
    gray = cv2.cvtColor(cell_bgr, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    up = cv2.resize(gray, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
    blurred = cv2.GaussianBlur(up, (0, 0), sigmaX=1.0)
    return cv2.addWeighted(up, 1.5, blurred, -0.5, 0)


# ── Overlay ───────────────────────────────────────────────────────────────────

def draw_overlay(img, h_pos, v_pos):
    overlay = img.copy()
    img_h, img_w = img.shape[:2]
    for y in h_pos:
        cv2.line(overlay, (0, y), (img_w, y), (0, 140, 255), 1)   # orange
    for x in v_pos:
        cv2.line(overlay, (x, 0), (x, img_h), (0, 210, 0), 1)     # green
    return overlay


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Extract bordered table data from a scanned image."
    )
    parser.add_argument("--image", default="Pages/206/206.png",
                        help="Path to the page image (default: Pages/206/206.png)")
    parser.add_argument("--h-frac", type=float, default=0.015,
                        help="Min horizontal fragment length as fraction of image width (default 0.015)")
    parser.add_argument("--v-frac", type=float, default=0.015,
                        help="Min vertical fragment length as fraction of image height (default 0.015)")
    parser.add_argument("--h-bridge", type=float, default=0.05,
                        help="Horizontal dilation to bridge inter-column gaps (default 0.05)")
    parser.add_argument("--v-bridge", type=float, default=0.02,
                        help="Vertical dilation to bridge inter-row gaps (default 0.02)")
    parser.add_argument("--threshold", type=float, default=0.02,
                        help="Projection threshold as fraction of row max (default 0.02)")
    parser.add_argument("--padding", type=int, default=3,
                        help="Pixels to trim inside each cell edge (default 3)")
    parser.add_argument("--no-ocr", action="store_true",
                        help="Skip OCR — detect grid only (fast)")
    parser.add_argument("--save-excel", action="store_true",
                        help="Save extracted table to Excel")
    parser.add_argument("--save-cells", action="store_true",
                        help="Save each cell as a PNG file")
    parser.add_argument("--cells-dir", default=None,
                        help="Directory to save cell PNGs (default: <image>_cells/ beside the image). "
                             "Set to 'Pages/206/Cells' to feed the pipeline directly.")
    parser.add_argument("--debug", action="store_true",
                        help="Save horizontal/vertical line masks")
    args = parser.parse_args()

    if not os.path.exists(args.image):
        print(f"Image not found: {args.image}")
        sys.exit(1)

    img = read_image(args.image)
    img_h, img_w = img.shape[:2]
    print(f"Image: {args.image}  ({img_w}×{img_h})")

    # ── Detect grid ───────────────────────────────────────────────────────────
    h_pos, v_pos, horiz_mask, vert_mask = detect_grid(
        img,
        h_min_len_frac=args.h_frac,
        v_min_len_frac=args.v_frac,
        h_bridge_frac=args.h_bridge,
        v_bridge_frac=args.v_bridge,
        threshold_ratio=args.threshold,
    )

    h_kernel_px = max(int(img_w * args.h_frac), 20)
    v_kernel_px = max(int(img_h * args.v_frac), 20)
    print(f"H kernel: {h_kernel_px}px ({args.h_frac:.2f}×{img_w})  "
          f"V kernel: {v_kernel_px}px ({args.v_frac:.2f}×{img_h})")
    print(f"Horizontal lines: {len(h_pos)}  y={h_pos}")
    print(f"Vertical lines:   {len(v_pos)}  x={v_pos}")

    if len(h_pos) < 2 or len(v_pos) < 2:
        print("\nNot enough grid lines detected. Try lowering --h-frac / --v-frac.")
        sys.exit(1)

    n_rows = len(h_pos) - 1
    n_cols = len(v_pos) - 1
    print(f"Grid: {n_rows} rows × {n_cols} cols")

    # ── Debug masks ───────────────────────────────────────────────────────────
    if args.debug:
        base = os.path.splitext(args.image)[0]
        write_image(base + "_debug_horiz.png", horiz_mask)
        write_image(base + "_debug_vert.png",  vert_mask)
        write_image(base + "_debug_grid.png",  cv2.add(horiz_mask, vert_mask))
        print("Debug masks saved.")

    # ── Overlay ───────────────────────────────────────────────────────────────
    overlay = draw_overlay(img, h_pos, v_pos)
    overlay_path = os.path.splitext(args.image)[0] + "_grid_overlay.png"
    write_image(overlay_path, overlay)
    print(f"Overlay saved -> {overlay_path}")
    print("  Orange lines = rows,  Green lines = columns")

    if args.no_ocr and not args.save_cells:
        return

    # ── Cell extraction ───────────────────────────────────────────────────────
    if args.cells_dir:
        cells_dir = args.cells_dir
    else:
        cells_dir = os.path.splitext(args.image)[0] + "_cells"
    if args.save_cells:
        os.makedirs(cells_dir, exist_ok=True)
        print(f"Cells -> {cells_dir}", flush=True)

    reader = None
    if not args.no_ocr:
        import easyocr
        print("\nLoading EasyOCR (en)...", flush=True)
        reader = easyocr.Reader(["en"])
        print("EasyOCR ready.", flush=True)

    pad = args.padding
    data = []
    total_cells = n_rows * n_cols

    for r in range(n_rows):
        y1 = h_pos[r]     + pad
        y2 = h_pos[r + 1] - pad
        if y2 <= y1:
            data.append([None] * n_cols)
            continue

        row_data = []
        for c in range(n_cols):
            x1 = v_pos[c]     + pad
            x2 = v_pos[c + 1] - pad
            if x2 <= x1:
                row_data.append(None)
                continue

            cell = img[y1:y2, x1:x2]
            if cell.size == 0:
                row_data.append(None)
                continue

            if args.save_cells:
                cell_path = os.path.join(cells_dir, f"row{r:03d}_col{c:03d}.png")
                write_image(cell_path, enhance_cell(cell))

            if reader is not None:
                try:
                    enhanced = enhance_cell(cell)
                    texts = reader.readtext(enhanced, detail=0)
                    row_data.append(" ".join(texts).strip() or None)
                except Exception as e:
                    print(f"  OCR error row={r} col={c}: {e}", flush=True)
                    row_data.append(None)
            else:
                row_data.append(None)

        done = (r + 1) * n_cols
        if reader is not None:
            print(f"  Row {r+1}/{n_rows} done ({done}/{total_cells} cells)", flush=True)

        data.append(row_data)

    if args.save_cells:
        print(f"Cell PNGs saved -> {cells_dir}/")

    # ── Print & export ────────────────────────────────────────────────────────
    print("\nAll rows processed. Saving...", flush=True)
    base = os.path.splitext(args.image)[0]

    # Always write CSV — no dependencies beyond stdlib
    import csv
    csv_path = base + "_table.csv"
    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(
                [("" if v is None else v) for v in row] for row in data
            )
        print(f"CSV saved  -> {csv_path}", flush=True)
    except Exception as e:
        print(f"CSV save failed: {e}", flush=True)

    # Excel via pandas (optional, may fail on Python 3.13 / missing openpyxl)
    if args.save_excel:
        try:
            import pandas as pd
            print("pandas imported OK.", flush=True)
            df = pd.DataFrame(data)
            print(f"DataFrame: {df.shape[0]} rows × {df.shape[1]} cols", flush=True)
            excel_path = base + "_table.xlsx"
            df.to_excel(excel_path, index=False, header=False)
            print(f"Excel saved -> {excel_path}", flush=True)
        except Exception:
            import traceback
            traceback.print_exc()
            print("Excel export failed — use the CSV instead.", flush=True)


if __name__ == "__main__":
    main()
