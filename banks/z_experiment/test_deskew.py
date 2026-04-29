"""
test_deskew.py — Correct skew in scanned document images using OpenCV.

Strategy:
  1. Binarise (Otsu) and invert so ink = white
  2. Detect edges with Canny
  3. Find lines via HoughLinesP, keep only near-horizontal ones (|angle| < 15°)
  4. Take the median angle as the skew
  5. Rotate the original image by -angle around its centre
  6. Save the deskewed result

Usage:
    python test_deskew.py
    python test_deskew.py --image Pages/3/3.png
    python test_deskew.py --image Pages/3/3.png --debug
"""

import argparse
import math
import os

import cv2
import numpy as np


def detect_skew_angle(gray: np.ndarray) -> float:
    """Return the estimated skew angle in degrees (positive = clockwise tilt)."""
    # Binarise: ink becomes white
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Detect edges
    edges = cv2.Canny(binary, 50, 150, apertureSize=3)

    # Detect line segments
    lines = cv2.HoughLinesP(
        edges,
        rho=1,
        theta=math.pi / 180,
        threshold=100,
        minLineLength=100,
        maxLineGap=10,
    )
    if lines is None:
        return 0.0

    angles = []
    for x1, y1, x2, y2 in lines[:, 0]:
        if x2 == x1:
            continue  # skip vertical lines
        angle = math.degrees(math.atan2(y2 - y1, x2 - x1))
        # Keep only near-horizontal lines
        if abs(angle) < 15:
            angles.append(angle)

    if not angles:
        return 0.0

    return float(np.median(angles))


def deskew(img: np.ndarray, angle: float) -> np.ndarray:
    """Rotate img by -angle degrees around its centre, keeping full canvas."""
    h, w = img.shape[:2]
    cx, cy = w / 2, h / 2
    M = cv2.getRotationMatrix2D((cx, cy), -angle, 1.0)

    # Expand canvas so corners aren't clipped
    cos, sin = abs(M[0, 0]), abs(M[0, 1])
    new_w = int(h * sin + w * cos)
    new_h = int(h * cos + w * sin)
    M[0, 2] += (new_w - w) / 2
    M[1, 2] += (new_h - h) / 2

    return cv2.warpAffine(
        img, M, (new_w, new_h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )


def main():
    parser = argparse.ArgumentParser(description="Deskew a scanned document image.")
    parser.add_argument("--image", default="Pages/3/3.png")
    parser.add_argument("--debug", action="store_true",
                        help="Save binary + edge images for inspection")
    args = parser.parse_args()

    img = cv2.imread(args.image)
    if img is None:
        print(f"Image not found: {args.image}")
        return

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    angle = detect_skew_angle(gray)
    print(f"Detected skew angle: {angle:.3f}°")

    if abs(angle) < 0.05:
        print("Image is already straight — no rotation needed.")
        return

    corrected = deskew(img, angle)

    stem, ext = os.path.splitext(args.image)
    out_path = f"{stem}_deskewed{ext}"
    cv2.imwrite(out_path, corrected)
    print(f"Deskewed image saved -> {out_path}")

    if args.debug:
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        edges = cv2.Canny(binary, 50, 150, apertureSize=3)
        cv2.imwrite(f"{stem}_binary.png", binary)
        cv2.imwrite(f"{stem}_edges.png", edges)
        print(f"Debug images: {stem}_binary.png, {stem}_edges.png")


if __name__ == "__main__":
    main()
