# pip install opencv-python numpy

import argparse
import math
import os
from pathlib import Path

import cv2
import numpy as np


def read_image(path: str):
    # Supports Unicode paths on Windows
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


def weighted_median(values, weights):
    values = np.asarray(values)
    weights = np.asarray(weights)

    order = np.argsort(values)
    values = values[order]
    weights = weights[order]

    cumulative = np.cumsum(weights)
    cutoff = weights.sum() / 2.0

    return float(values[np.searchsorted(cumulative, cutoff)])


def normalize_line_angle(angle):
    while angle >= 90:
        angle -= 180
    while angle < -90:
        angle += 180
    return angle


def binarize_for_lines(gray):
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.threshold(
        blurred, 0, 255,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )[1]
    return binary


def estimate_skew_from_horizontal_lines(gray, max_skew=8, y_band=None):
    """Detect skew angle from horizontal lines.

    y_band: (min_frac, max_frac) to restrict detection to a vertical slice,
            e.g. (0.6, 0.98) for the bottom 40%. None = full image.
    """
    h, w = gray.shape[:2]
    binary = binarize_for_lines(gray)

    kernel_width = max(40, w // 18)
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_width, 1))

    line_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    line_mask = cv2.dilate(
        line_mask,
        cv2.getStructuringElement(cv2.MORPH_RECT, (3, 1)),
        iterations=1
    )

    lines = cv2.HoughLinesP(
        line_mask,
        rho=1,
        theta=np.pi / 180,
        threshold=max(60, w // 20),
        minLineLength=int(w * 0.25),
        maxLineGap=int(w * 0.04)
    )

    if lines is None:
        return None, line_mask

    angles = []
    weights = []
    top_bottom_margin = int(h * 0.02)

    y_min = int(h * y_band[0]) if y_band else top_bottom_margin
    y_max = int(h * y_band[1]) if y_band else h - top_bottom_margin

    for line in lines[:, 0]:
        x1, y1, x2, y2 = map(int, line)

        mid_y = (y1 + y2) / 2
        if mid_y < y_min or mid_y > y_max:
            continue

        dx = x2 - x1
        dy = y2 - y1
        length = math.hypot(dx, dy)

        if length < w * 0.25:
            continue

        angle = math.degrees(math.atan2(dy, dx))
        angle = normalize_line_angle(angle)

        if abs(angle) <= max_skew:
            angles.append(angle)
            weights.append(length)

    if len(angles) < 2:
        return None, line_mask

    return weighted_median(angles, weights), line_mask


def rotate_same_size(binary, angle):
    h, w = binary.shape[:2]
    center = (w / 2, h / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(
        binary, matrix, (w, h),
        flags=cv2.INTER_NEAREST,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=0
    )
    return rotated


def estimate_skew_from_projection(gray, max_skew=5):
    h, w = gray.shape[:2]

    scale = min(1.0, 1000 / max(h, w))
    small = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale < 1.0 else gray.copy()

    binary = binarize_for_lines(small)

    hh, ww = binary.shape[:2]
    margin_y = int(hh * 0.02)
    margin_x = int(ww * 0.02)
    binary[:margin_y, :] = 0
    binary[-margin_y:, :] = 0
    binary[:, :margin_x] = 0
    binary[:, -margin_x:] = 0

    def score_angle(angle):
        rotated = rotate_same_size(binary, angle)
        row_sum = np.sum(rotated, axis=1, dtype=np.float64)
        return np.sum(np.diff(row_sum) ** 2)

    coarse_angles = np.arange(-max_skew, max_skew + 0.001, 0.5)
    best_coarse = max(coarse_angles, key=score_angle)

    fine_angles = np.arange(best_coarse - 0.6, best_coarse + 0.601, 0.05)
    best_fine = max(fine_angles, key=score_angle)

    return float(best_fine)


def rotate_bound_white(img, angle):
    h, w = img.shape[:2]
    center = (w / 2, h / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)

    cos = abs(matrix[0, 0])
    sin = abs(matrix[0, 1])
    new_w = int((h * sin) + (w * cos))
    new_h = int((h * cos) + (w * sin))

    matrix[0, 2] += (new_w / 2) - center[0]
    matrix[1, 2] += (new_h / 2) - center[1]

    return cv2.warpAffine(
        img, matrix, (new_w, new_h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(255, 255, 255)
    )


def deskew_image(img, use_bottom=False):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    y_band = (0.6, 0.98) if use_bottom else None
    angle, line_mask = estimate_skew_from_horizontal_lines(gray, y_band=y_band)

    if angle is None:
        print("  Không tìm thấy đường kẻ bảng rõ ràng. Dùng phương pháp dự phòng...")
        angle = estimate_skew_from_projection(gray)

    corrected = rotate_bound_white(img, angle)
    return corrected, angle, line_mask


def deskew_page(page_num: int, input_dir: str = "Pages", debug: bool = False, use_bottom: bool = False):
    page_dir = os.path.join(input_dir, str(page_num))
    img_path = os.path.join(page_dir, f"{page_num}.png")

    if not os.path.isfile(img_path):
        print(f"  Không tìm thấy ảnh: {img_path}, bỏ qua.")
        return

    img = read_image(img_path)
    corrected, angle, line_mask = deskew_image(img, use_bottom=use_bottom)
    write_image(img_path, corrected)

    print(f"  Góc hiệu chỉnh: {angle:.3f}° — đã lưu: {img_path}")

    if debug:
        mask_path = os.path.join(page_dir, f"{page_num}_line_mask.png")
        write_image(mask_path, line_mask)
        print(f"  Debug line mask: {mask_path}")


def main():
    parser = argparse.ArgumentParser(description="Deskew scanned page images.")
    parser.add_argument("--input", default="Pages",
                        help="Folder containing page sub-folders (default: Pages)")
    parser.add_argument("--page", type=int, default=None,
                        help="Process only this page number; omit to process all pages")
    parser.add_argument("--debug", action="store_true",
                        help="Save detected line mask alongside each output")
    parser.add_argument("--bottom", action="store_true",
                        help="Detect angle using only lines in the bottom 40%% of the image")
    args = parser.parse_args()

    if args.page:
        pages = [args.page]
    else:
        pages = sorted(
            int(p) for p in os.listdir(args.input)
            if os.path.isdir(os.path.join(args.input, p)) and p.isdigit()
        )

    for page_num in pages:
        print(f"Đang xử lý trang {page_num}...")
        deskew_page(page_num, args.input, args.debug, use_bottom=args.bottom)


if __name__ == "__main__":
    main()
