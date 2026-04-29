# deskew_scan.py
# pip install opencv-python numpy

import argparse
import math
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
    """
    Normalize angle to [-90, 90).
    """
    while angle >= 90:
        angle -= 180
    while angle < -90:
        angle += 180
    return angle


def binarize_for_lines(gray):
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)

    binary = cv2.threshold(
        blurred,
        0,
        255,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )[1]

    return binary


def estimate_skew_from_horizontal_lines(gray, max_skew=8):
    """
    Best for bank statements / tables because it uses long horizontal table lines.
    Returns correction angle in degrees.
    """
    h, w = gray.shape[:2]
    binary = binarize_for_lines(gray)

    # Extract long horizontal lines
    kernel_width = max(40, w // 18)
    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (kernel_width, 1)
    )

    line_mask = cv2.morphologyEx(
        binary,
        cv2.MORPH_OPEN,
        horizontal_kernel,
        iterations=1
    )

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

    for line in lines[:, 0]:
        x1, y1, x2, y2 = map(int, line)

        # Ignore scanner/image borders
        if min(y1, y2) < top_bottom_margin:
            continue
        if max(y1, y2) > h - top_bottom_margin:
            continue

        dx = x2 - x1
        dy = y2 - y1
        length = math.hypot(dx, dy)

        if length < w * 0.25:
            continue

        angle = math.degrees(math.atan2(dy, dx))
        angle = normalize_line_angle(angle)

        # We only want nearly-horizontal lines
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
        binary,
        matrix,
        (w, h),
        flags=cv2.INTER_NEAREST,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=0
    )

    return rotated


def estimate_skew_from_projection(gray, max_skew=5):
    """
    Fallback method.
    It tries several angles and chooses the one where text rows become most horizontal.
    Slower, but useful when table lines are weak.
    """
    h, w = gray.shape[:2]

    scale = min(1.0, 1000 / max(h, w))
    if scale < 1.0:
        small = cv2.resize(
            gray,
            None,
            fx=scale,
            fy=scale,
            interpolation=cv2.INTER_AREA
        )
    else:
        small = gray.copy()

    binary = binarize_for_lines(small)

    # Remove outer scanner borders from scoring
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
        score = np.sum(np.diff(row_sum) ** 2)
        return score

    # Coarse search
    coarse_angles = np.arange(-max_skew, max_skew + 0.001, 0.5)
    coarse_scores = [(score_angle(a), a) for a in coarse_angles]
    best_coarse = max(coarse_scores, key=lambda x: x[0])[1]

    # Fine search around best coarse angle
    fine_angles = np.arange(best_coarse - 0.6, best_coarse + 0.601, 0.05)
    fine_scores = [(score_angle(a), a) for a in fine_angles]
    best_fine = max(fine_scores, key=lambda x: x[0])[1]

    return float(best_fine)


def rotate_bound_white(img, angle):
    """
    Rotate without cropping. New background is white.
    """
    h, w = img.shape[:2]
    center = (w / 2, h / 2)

    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)

    cos = abs(matrix[0, 0])
    sin = abs(matrix[0, 1])

    new_w = int((h * sin) + (w * cos))
    new_h = int((h * cos) + (w * sin))

    matrix[0, 2] += (new_w / 2) - center[0]
    matrix[1, 2] += (new_h / 2) - center[1]

    rotated = cv2.warpAffine(
        img,
        matrix,
        (new_w, new_h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(255, 255, 255)
    )

    return rotated


def deskew_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    angle, line_mask = estimate_skew_from_horizontal_lines(gray)

    if angle is None:
        print("Could not detect strong table lines. Using projection fallback...")
        angle = estimate_skew_from_projection(gray)

    corrected = rotate_bound_white(img, angle)

    return corrected, angle, line_mask


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input scanned image")
    parser.add_argument("output", help="Output deskewed image")
    parser.add_argument("--debug", action="store_true", help="Save detected line mask")

    args = parser.parse_args()

    img = read_image(args.input)

    corrected, angle, line_mask = deskew_image(img)

    write_image(args.output, corrected)

    print(f"Detected correction angle: {angle:.3f} degrees")
    print(f"Saved: {args.output}")

    if args.debug:
        debug_path = str(Path(args.output).with_name(Path(args.output).stem + "_line_mask.png"))
        write_image(debug_path, line_mask)
        print(f"Saved debug line mask: {debug_path}")


if __name__ == "__main__":
    main()