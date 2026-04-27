import cv2
import numpy as np
import os
import argparse


def find_line_positions(line_img, axis, gap_threshold=10):
    """Project a binary line image onto one axis and return the center of each detected line."""
    projection = np.sum(line_img, axis=1 if axis == "h" else 0)
    max_val = np.max(projection)
    if max_val == 0:
        return []

    indices = np.where(projection > max_val * 0.1)[0]
    if len(indices) == 0:
        return []

    groups, current = [], [indices[0]]
    for idx in indices[1:]:
        if idx - current[-1] <= gap_threshold:
            current.append(idx)
        else:
            groups.append(int(np.mean(current)))
            current = [idx]
    groups.append(int(np.mean(current)))
    return groups


def _enhance_cell(cell_bgr):
    """
    Sharpen and upscale a cell crop for better OCR accuracy.

    Steps:
      1. Grayscale
      2. 2× upscale with cubic interpolation so EasyOCR sees larger glyphs
      3. Unsharp mask to sharpen edges without distorting stroke widths
    """
    gray = cv2.cvtColor(cell_bgr, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    up = cv2.resize(gray, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
    blurred = cv2.GaussianBlur(up, (0, 0), sigmaX=1.0)
    return cv2.addWeighted(up, 1.5, blurred, -0.5, 0)


def split_table_cells(image_path, output_dir, padding=2, debug=False):
    img = cv2.imread(image_path)
    if img is None:
        print(f"Không thể đọc: {image_path}")
        return 0

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape

    # Invert so dark lines become white for morphological detection
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Horizontal lines: erode with a wide kernel, then dilate to fill gaps
    h_len = max(w // 25, 20)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1))
    horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel, iterations=1)
    horizontal = cv2.dilate(horizontal, np.ones((3, 1), np.uint8), iterations=1)

    # Vertical lines: erode with a tall kernel, then dilate to fill gaps
    v_len = max(h // 25, 20)
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_len))
    vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel, iterations=1)
    vertical = cv2.dilate(vertical, np.ones((1, 3), np.uint8), iterations=1)

    if debug:
        os.makedirs(output_dir, exist_ok=True)
        cv2.imwrite(os.path.join(output_dir, "_debug_horizontal.png"), horizontal)
        cv2.imwrite(os.path.join(output_dir, "_debug_vertical.png"), vertical)
        cv2.imwrite(os.path.join(output_dir, "_debug_grid.png"), cv2.add(horizontal, vertical))

    h_pos = find_line_positions(horizontal, "h")
    v_pos = find_line_positions(vertical, "v")

    print(f"  Đường ngang: {len(h_pos)} -> y={h_pos}")
    print(f"  Đường dọc:        {len(v_pos)} -> x={v_pos}")

    if len(h_pos) < 2 or len(v_pos) < 2:
        print("  Không phát hiện đủ đường kẻ để tạo thành ô.")
        return 0

    os.makedirs(output_dir, exist_ok=True)

    # Decide how many row intervals to process.
    # If the last interval is much taller than the median data row it is footer/empty
    # space below the table border — skip it. Otherwise include it as a data row.
    row_heights = [h_pos[i + 1] - h_pos[i] for i in range(len(h_pos) - 1)]
    median_height = float(np.median(row_heights[:-1])) if len(row_heights) > 1 else row_heights[0]
    if row_heights[-1] > median_height * 1.8:
        n_rows = len(h_pos) - 2   # last interval is footer space, skip it
    else:
        n_rows = len(h_pos) - 1   # last interval is a real data row, keep it

    count = 0
    for row_idx in range(n_rows):
        y1 = h_pos[row_idx] + padding
        y2 = h_pos[row_idx + 1] - padding
        if y2 <= y1:
            continue
        for col_idx in range(len(v_pos) - 1):
            x1 = v_pos[col_idx] + padding
            x2 = v_pos[col_idx + 1] - padding
            if x2 <= x1:
                continue
            cell = img[y1:y2, x1:x2]
            if cell.size == 0:
                continue
            cv2.imwrite(os.path.join(output_dir, f"row{row_idx:03d}_col{col_idx:03d}.png"), _enhance_cell(cell))
            count += 1

    return count


def main():
    parser = argparse.ArgumentParser(description="Detect table grid lines and split into cell images.")
    parser.add_argument("--input", default="Pages",
                        help="Folder containing page sub-folders from ConvertToImages.py (default: Pages)")
    parser.add_argument("--page", type=int, default=None,
                        help="Process only this page number; omit to process all pages")
    parser.add_argument("--padding", type=int, default=2,
                        help="Pixels to trim from each cell edge to remove line artifacts (default: 2)")
    parser.add_argument("--debug", action="store_true",
                        help="Save intermediate line-detection images alongside cells")
    args = parser.parse_args()

    if args.page:
        pages = [args.page]
    else:
        pages = sorted(
            int(p) for p in os.listdir(args.input)
            if os.path.isdir(os.path.join(args.input, p)) and p.isdigit()
        )

    for page_num in pages:
        img_path = os.path.join(args.input, str(page_num), f"{page_num}.png")
        if not os.path.exists(img_path):
            print(f"Không tìm thấy ảnh: {img_path}, bỏ qua.")
            continue
        out_dir = os.path.join(args.input, str(page_num), "Cells")
        print(f"Đang xử lý trang {page_num}...")
        count = split_table_cells(img_path, out_dir, padding=args.padding, debug=args.debug)
        print(f"  Đã trích xuất {count} ô -> {out_dir}")


if __name__ == "__main__":
    main()
