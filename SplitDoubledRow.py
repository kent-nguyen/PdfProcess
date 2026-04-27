import cv2
import numpy as np
import os
import argparse


def detect_split_y(cell_img):
    """
    Find the y-coordinate of the internal horizontal divider inside a doubled cell.
    Searches only the middle 60% of the image height to avoid the cell border lines.
    Returns the detected y, or None if nothing found (caller should fall back to midpoint).
    """
    gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape

    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    h_len = max(w // 4, 10)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1))
    horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel, iterations=1)

    projection = np.sum(horizontal, axis=1).astype(float)

    # Only look in the middle 60% to avoid border artifacts
    margin = h // 5
    search = projection[margin : h - margin]

    if np.max(search) == 0:
        return None

    return int(np.argmax(search)) + margin


def split_doubled_row(cells_dir, row_idx):
    # Collect columns for this row
    prefix = f"row{row_idx:03d}_col"
    col_files = sorted(f for f in os.listdir(cells_dir) if f.startswith(prefix) and f.endswith(".png"))
    if not col_files:
        raise FileNotFoundError(f"No cells found for row {row_idx} in {cells_dir}")

    n_cols = len(col_files)
    col_indices = [int(f[len(prefix): len(prefix) + 3]) for f in col_files]

    # Find the last row that exists
    all_row_indices = set()
    for f in os.listdir(cells_dir):
        if f.startswith("row") and f.endswith(".png") and len(f) >= 6:
            try:
                all_row_indices.add(int(f[3:6]))
            except ValueError:
                pass
    max_row = max(all_row_indices)

    print(f"Cells dir : {cells_dir}")
    print(f"Splitting row {row_idx} (columns: {n_cols}, last row: {max_row})")

    # ── Step 1: detect split point from col000 (or first available col) ──────
    first_cell_path = os.path.join(cells_dir, col_files[0])
    first_img = cv2.imread(first_cell_path)
    if first_img is None:
        raise IOError(f"Cannot read {first_cell_path}")

    cell_h = first_img.shape[0]
    split_y = detect_split_y(first_img)
    if split_y is None:
        split_y = cell_h // 2
        print(f"  No divider detected — splitting at midpoint y={split_y}")
    else:
        print(f"  Detected divider at y={split_y} (cell height={cell_h})")

    # ── Step 2: shift rows (row_idx+1 … max_row) down by one, in reverse ─────
    print(f"  Renaming rows {row_idx + 1}–{max_row}  →  {row_idx + 2}–{max_row + 1}")
    for r in range(max_row, row_idx, -1):
        for c in col_indices:
            src = os.path.join(cells_dir, f"row{r:03d}_col{c:03d}.png")
            dst = os.path.join(cells_dir, f"row{r + 1:03d}_col{c:03d}.png")
            if os.path.exists(src):
                os.rename(src, dst)

    # ── Step 3: split each cell in the problematic row ───────────────────────
    print(f"  Splitting {n_cols} cells …")
    for c in col_indices:
        cell_path = os.path.join(cells_dir, f"row{row_idx:03d}_col{c:03d}.png")
        img = cv2.imread(cell_path)
        if img is None:
            print(f"    [WARN] Cannot read {cell_path}, skipping")
            continue

        top = img[:split_y, :]
        bottom = img[split_y:, :]

        cv2.imwrite(cell_path, top)
        new_path = os.path.join(cells_dir, f"row{row_idx + 1:03d}_col{c:03d}.png")
        cv2.imwrite(new_path, bottom)

    print(f"Done. Row {row_idx} split into rows {row_idx} and {row_idx + 1}.")
    print(f"      Rows {row_idx + 2}-{max_row + 1} shifted from original {row_idx + 1}-{max_row}.")


def main():
    parser = argparse.ArgumentParser(
        description="Split a doubled row (two STT entries in one cell row) into two proper rows."
    )
    parser.add_argument(
        "--cells-dir",
        required=True,
        help="Path to the Cells directory, e.g. Pages/97/Cells",
    )
    parser.add_argument(
        "--row",
        type=int,
        required=True,
        help="0-based row index of the problematic doubled row, e.g. 19",
    )
    args = parser.parse_args()

    split_doubled_row(args.cells_dir, args.row)


if __name__ == "__main__":
    main()
