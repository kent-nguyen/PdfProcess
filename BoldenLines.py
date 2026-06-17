"""
BoldenLines.py — Convert dotted/dashed table borders to solid lines.

Runs after ConvertToImages + Deskew, before SplitTableCells.

Uses HoughLinesP with a generous maxLineGap to bridge the gaps between dots/dashes,
then redraws each detected segment as a solid black stroke. This avoids the text-
smearing problem that morphological CLOSE operations cause.

Usage:
    python BoldenLines.py --page 1
    python BoldenLines.py --page 1-5,8
    python BoldenLines.py          # all pages
    python BoldenLines.py --page 1 --debug
"""

import argparse
import math
import os
from pathlib import Path

import cv2
import numpy as np

from libs.page_range import parse_pages


def read_image(path: str):
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Cannot read image: {path}")
    return img


def write_image(path: str, img):
    ext = Path(path).suffix or ".png"
    ok, buf = cv2.imencode(ext, img)
    if not ok:
        raise ValueError(f"Cannot write image: {path}")
    buf.tofile(path)


def bolden_lines(img, max_gap: int = 20, min_length_frac: float = 0.50,
                 line_thickness: int = 2) -> np.ndarray:
    """
    Detect dotted/dashed horizontal and vertical table borders and
    redraw them as solid black lines using the Hough line transform.

    Args:
        max_gap:          Maximum pixel gap to bridge between dots/dashes (default 20).
        min_length_frac:  Minimum segment length as a fraction of the page dimension
                          (default 0.25). Filters out text — table borders are much longer.
        line_thickness:   Thickness of the redrawn solid lines in pixels (default 2).
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape

    # Gentle blur then Canny so we get clean edges without noise spikes
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    edges = cv2.Canny(blurred, 30, 100)

    min_h_len = int(w * min_length_frac)   # horizontal lines must span this many px
    min_v_len = int(h * min_length_frac)   # vertical lines must span this many px

    segments = cv2.HoughLinesP(
        edges,
        rho=1,
        theta=np.pi / 360,       # 0.5° angular resolution
        threshold=30,
        minLineLength=min(min_h_len, min_v_len),
        maxLineGap=max_gap,
    )

    result = img.copy()
    n_h = n_v = 0

    if segments is not None:
        for seg in segments:
            x1, y1, x2, y2 = seg[0]
            dx, dy = x2 - x1, y2 - y1
            length = math.hypot(dx, dy)
            if length == 0:
                continue
            angle = abs(math.degrees(math.atan2(dy, dx)))

            if angle < 5 and length >= min_h_len:
                # Horizontal: draw at the average y of the two endpoints
                y_mid = (y1 + y2) // 2
                cv2.line(result, (x1, y_mid), (x2, y_mid), (0, 0, 0), line_thickness)
                n_h += 1

            elif 85 < angle < 95 and length >= min_v_len:
                # Vertical: draw at the average x of the two endpoints
                x_mid = (x1 + x2) // 2
                cv2.line(result, (x_mid, y1), (x_mid, y2), (0, 0, 0), line_thickness)
                n_v += 1

    print(f"  Phát hiện: {n_h} đoạn ngang, {n_v} đoạn dọc")
    return result


def bolden_page(page_num: int, input_dir: str = "Pages",
                max_gap: int = 20, min_length_frac: float = 0.25,
                line_thickness: int = 2, debug: bool = False):
    page_dir = os.path.join(input_dir, str(page_num))
    img_path = os.path.join(page_dir, f"{page_num}.png")

    if not os.path.isfile(img_path):
        print(f"  Không tìm thấy ảnh: {img_path}, bỏ qua.")
        return

    img = read_image(img_path)
    result = bolden_lines(img, max_gap=max_gap, min_length_frac=min_length_frac,
                          line_thickness=line_thickness)

    if debug:
        debug_path = os.path.join(page_dir, f"{page_num}_bolden.png")
        write_image(debug_path, result)
        print(f"  Debug: {debug_path}")
    else:
        write_image(img_path, result)
        print(f"  Đã lưu: {img_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert dotted/dashed table borders to solid lines."
    )
    parser.add_argument("--input", default="Pages",
                        help="Folder containing page sub-folders (default: Pages)")
    parser.add_argument("--page", type=str, default=None,
                        help="Pages to process, e.g. '1' or '1,3,5' or '2-5'. Omit for all pages.")
    parser.add_argument("--gap", type=int, default=20,
                        help="Max gap in pixels to bridge between dots/dashes (default: 20).")
    parser.add_argument("--min-length", type=float, default=0.50, dest="min_length",
                        help="Min segment length as fraction of page dimension (default: 0.25).")
    parser.add_argument("--thickness", type=int, default=2,
                        help="Solid line thickness in pixels (default: 2).")
    parser.add_argument("--debug", action="store_true",
                        help="Save result as {page}_bolden.png without overwriting the original.")
    args = parser.parse_args()

    if args.page:
        pages = parse_pages(args.page)
    else:
        pages = sorted(
            int(p) for p in os.listdir(args.input)
            if os.path.isdir(os.path.join(args.input, p)) and p.isdigit()
        )

    for page_num in pages:
        print(f"Đang xử lý trang {page_num}...")
        bolden_page(page_num, args.input,
                    max_gap=args.gap, min_length_frac=args.min_length,
                    line_thickness=args.thickness, debug=args.debug)


if __name__ == "__main__":
    main()
